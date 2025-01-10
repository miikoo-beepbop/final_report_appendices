// Global variables
const projectId = 'e-dragon-428800-f0'; // Google BigQuery Project ID
const datasetName = 'ET_Output'; // Google BigQuery Dataset ID
const tableName = 'mdv2'; // Google BigQuery Table ID for old ET output
const sheetId = '14thocGSc9x-eaVVVMITJ3DWSj3SNUR1uhQQrNzDL4wE' // Google Sheets ID for new ET output
const sheetName = 'MDv2'; // Google Sheets Sheet name

const headers = [
  "timestamp", "done_by_testing", "done_by_tracing", "split_id", "lot_id", "lot_id_suffix", "tool_number", "kevin_probe_test", "for_4wire", "date_code", "date_code_input", "job_type", "machine", "xcut_allowed", "ipc_class", "mass_lam", "resistive", "continuity", "test_voltage", "isolation", "adjacency_used", "mfg_qty", "tested_qty", "first_passed_qty", "open", "short", "sdp", "open_circuit_reason", "qty_reject_080_il_strip", "qty_reject_chema_180_ol_cm_dm_o", "qty_reject_477_ol_cm_dm_o", "qty_reject_473_ol_lpsm_cure", "qty_reject_473_ol_cm_osp", "qty_reject_473_ol_cm_strp", "qty_reject_133_ol_drl", "qty_repair_outerlayer_open", "qty_reject_photoprint_080_ol_pp_dev", "qty_repair_mask_on_pad_finger", "qty_reject_205_ol_sm_cure", "qty_repair_plating_nodule", "qty_reject_178_ol_cm_strp", "qty_repair_legend_on_pad_hole", "qty_reject_233_ol_lpsm_cure", "qty_repair_others_open", "qty_reject_others_open", "qty_repair_false_oc", "short_circuit_reason", "qty_repair_short", "qty_reject_photoprint_083_ol_pp_dev", "qty_repair_scratches", "qty_reject_photoprint_483_ol_pp_dev", "qty_reject_483_chemb_ol_cm_strp", "qty_reject_483_chemc_ol_cm_osp", "qty_repair_cu_residue", "qty_reject_chemb_172_ol_cm_strp", "qty_repair_under_etch", "qty_reject_chemb_076_ol_cm_strp", "qty_reject_083_il_strip", "qty_repair_misregistration", "qty_reject_471_ol_lam", "qty_reject_469_ol_lam", "qty_reject_132_ol_drl", "qty_repair_micro_short", "qty_reject_chemb_083_ol_cm_strp", "qty_reject_lpsm_083_ol_sm_cure", "qty_repair_incomplete_solder_strip", "qty_reject_chemc_256_ol_cm_osp", "qty_repair_feathering", "qty_reject_chemc_250_ol_cm_osp", "qty_repair_incomplete_resist_strip", "qty_reject_chemb_51_ol_cm_strp", "qty_repair_npth_short", "qty_reject_142", "qty_repair_short_others", "qty_reject_short_others", "qty_reject_479_ol_lam", "qty_reject_479_ol_pp_dev", "qty_reject_32_ol_lam", "impedance_fail", "impedance_fail_high_qty", "impedance_fail_low_qty", "final_passed_qty", "full_lot_id"
]

// Transfer data daily from Google Sheets to BigQuery
function dailyTransfer() {
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName)
  const data = sheet.getDataRange().getValues(); // Get all data from the sheet
  const rows = data.slice(2); // Start reading data from the 3rd row onwards

  if (!rows.length) {
    Logger.log("No data found for daily transfer.");
    return;
  }

  // Get yesterday's date in YYYY-MM-DD format
  const yesterday = new Date ();
  yesterday.setDate(yesterday.getDate() - 1); // Go back one day
  const yesterdayString = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Filter rows for yesterday's date
  const timestampIndex = headers.indexOf("timestamp");
  const filteredRows = rows.filter(row => {
    const rowDate = newDate(row[timestampIndex]);
    const rowDateString = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    return rowDateString === yesterdayString;
  })

  if (!filteredRows.length) {
    Logger.log("No data for today found in the sheet.");
    return;
  }

  // Prepare rows for BigQuery
  const bigQueryData = rows.map(row => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = row[index] || null; // Match data columns to the custom headers
    });
    return { json: record };
  });

  // Insert daily data into BigQuery
  try {
    BigQuery.Tabledata.insertAll(BigQuery.Project.getId(), datasetName, tableName, { rows: bigQueryData });
    Logger.log("Daily data transferred successfully.");
  } catch (error) {
    Logger.log("Error during daily data transfer: " + error.message);
  }
}

// Transfers monthly data
function monthlyTransfer() {
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName)
  const data = sheet.getDataRange().getValues(); // Get all data from the sheet
  const rows = data.slice(2); // Start reading data from the 3rd row onwards

  if (!rows.length) {
    Logger.log("No data found for daily transfer.");
    return;
  }

  // Get previous month and year
  const now = new Date();
  now.setDate(1); // Move to the first day of the current month
  now.setMonth(now.getMonth() - 1); // Go back one month
  const previousYear = now.getFullYear();
  const previousMonth = now.getMonth() + 1; // Months are 0-indexed

  // Prepare rows for BigQuery, filter for the current month
  const bigQueryData = rows.map(row => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = row[index];
    });
    return record;
  })
  .filter(record => {
    const dateString = record.timestamp.split(" ")[0]; // Extract only the date part
    const date = new Date(dateString); // Convert to a Date object
    return (
      date.getFullYear() === previousYear &&
      date.getMonth() + 1 === previousMonth
    );
  });

  // Delete existing rows for the current month in BigQuery
  const tableId = `${datasetName}.${tableName}`;
  const deleteQuery = `DELETE FROM \`${tableId}\` WHERE EXTRACT(YEAR FROM timestamp) = ${previousYear} AND EXTRACT(MONTH FROM timestamp) = ${previousMonth}`;
  
  try {
    BigQuery.Jobs.query({ query: deleteQuery }, BigQuery.Project.getId());
    Logger.log("Existing monthly data deleted in BigQuery.");
  } catch (error) {
    Logger.log("Error during monthly data deletion:", error);
    return;
  }

  // Insert monthly data into BigQuery
  if (bigQueryData.length > 0) {
    try {
      BigQuery.Tabledata.insertAll({
        projectId: BigQuery.Project.getId(),
        datasetId: datasetName,
        tableId: tableName,
        rows: bigQueryData.map(record => ({ json: record })),
      });
      Logger.log("Monthly data inserted into BigQuery.");
    } catch (error) {
      Logger.log("Error during monthly data insertion: " + error.message);
    }
  }
}


