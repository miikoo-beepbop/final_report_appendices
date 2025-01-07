function doGet() {
    Logger.log("doGet called");
    var template = HtmlService.createTemplateFromFile('mainform')
        .evaluate()
        .setTitle('INSP-07-FM-07 REV 08 ET Form')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    Logger.log("Template evaluated");
    return template;
  }
  
  function include(filename) {
    Logger.log("Including file: " + filename);
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }
  
  function getEmployeeIds() {
    const spreadsheetId = '1kXb0eMLorTf46t7apG3Kg0wfwDS9ezTjOFdkOf-nbDI';
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Sheet1');
    const data = sheet.getRange('A2:A').getValues();
    // Flatten array and remove empty values
    const ids = data.flat().filter(id => id !== '');
    return JSON.stringify(ids); // Return as JSON
  }
  
  function getFullLotIds(spreadsheetId, sheetName, range) {
    var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }
  
    // Get all values in the full lot id column, skip the header
    var fullLotIds = sheet.getRange(range).getValues();
  
    // Flatten array and remove empty values
    const ids = fullLotIds.flat().filter(id => id !== '');
    return JSON.stringify(ids); // Return as JSON
  }
  
  function fetchFormData(splitId, lotId, lotIdSuffix) {
    try {
      var sheet = SpreadsheetApp.openById('1hFpcDc7CtFx4MHh4jalZkbE041xQvyULyS_Tz47vBUo').getSheetByName('Form1Data');
      if (!sheet) {
        throw new Error('Sheet "Form1Data" not found.');
      }
  
      // Get all rows
      var data = sheet.getDataRange().getValues(); 
      var headers = data[0].concat(data[1]);
      var matchingRow = null;
  
      // Iterate through rows to find a match
      for (var i = 2; i < data.length; i++) {
        const row = data[i];
        const matchesSplitId = !splitId || row[1] === splitId;
        const matchesLotId = row[2] == lotId;
        const matchesLotIdSuffix = !lotIdSuffix || row[3] === lotIdSuffix;
  
        if (matchesSplitId && matchesLotId && matchesLotIdSuffix) {
          matchingRow = row;
          break;
        }
      }
  
      if (matchingRow) {
        // Map matching row to headers
        var formData = {};
        headers.forEach((header, index) => {
          formData[header] = matchingRow[index];
        });
  
        return formData;
      }
      else {
        return null;
      }
    }
    catch (e) {
      Logger.log('Error in fetchFormData: ' + e.toString());
      return null;
    }
  }
  
  function saveFormData(formData) {
    try{
      var sheet = SpreadsheetApp.openById('1hFpcDc7CtFx4MHh4jalZkbE041xQvyULyS_Tz47vBUo').getSheetByName('Form1Data');
      if (!sheet) {
        throw new Error('Sheet "Form1Data" not found.');
      }
  
      var headers = ['done_by_1', 'split_id', 'lot_id', 'lot_id_suffix', 'tool_number', 'kevin_probe_test', 'for_4wire_input', 'date_code', 'date_code_input', 'job_type', 'machine', 'xcut_allowed', 'ipc_class', 'mass_lam', 'resistive', 'continuity', 'test_voltage', 'isolation', 'adjacency_used', 'mfg_qty', 'tested_qty', 'first_passed_qty', 'open', 'short', 'sdp'];
      var row = [];
  
      // Save the form data
      headers.forEach(function(header) {
        // row.push(formData[header] || '');
        let value = formData[header] || '';
        // Convert value to string and handle parentheses
        if (typeof value === 'string' && value.match(/^\(.*\)$/)) {
          value = `'${value}`; // Add single quote to preserve parentheses
        }
        row.push(value.toString());
      });
  
      // Concatenate Split ID, Lot ID and Lot ID Suffix 
      var splitId = formData['split_id'] || '';
      var lotId = formData['lot_id'];
      var lotIdSuffix = formData['lot_id_suffix'] || '';
      var concatenatedId = splitId + lotId + lotIdSuffix; // Concatenate values
      row.push(concatenatedId);
  
      sheet.appendRow(row)
    }
    catch (e) {
      Logger.log('Error in saving form data: ' + e.toString());
      return 'Error in saving data: ' + e.toString();
    }
  }
  
  function autoFillForm2(splitId, lotId, lotIdSuffix) {
    try {
      var sheet = SpreadsheetApp.openById('1hFpcDc7CtFx4MHh4jalZkbE041xQvyULyS_Tz47vBUo').getSheetByName('Form1Data');
      if (!sheet) {
        throw new Error('Sheet "Form1Data" not found.');
      }
  
      var data = sheet.getDataRange().getValues();
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
      var splitIdIndex = 1;
      var lotIdIndex = 2;
      var lotIdSuffixIndex = 3;
  
      // Iterate through the rows to find a match
      for (var i = 1; i < data.length; i++) {
        const matchesSplitId= splitId === "" || data[i][splitIdIndex] === splitId;
        const matchesLotId = data[i][lotIdIndex] == lotId;
        const matchesLotIdSuffix = lotIdSuffix === "" || data[i][lotIdSuffixIndex] === lotIdSuffix 
  
        if (matchesSplitId && matchesLotId && matchesLotIdSuffix) {
          const formData = {};
          headers.forEach((header, index) => {
            formData[header] = data[i][index];
          });
  
          return formData; 
        } 
      }
  
      return null
    }
    catch (e) {
      Logger.log('Error in auto-filling form: ' + e.toString());
      return null;
    }
  }
  
  function deleteRowByFullLotId(fullLotId) {
    try {
      var sheet = SpreadsheetApp.openById('1hFpcDc7CtFx4MHh4jalZkbE041xQvyULyS_Tz47vBUo').getSheetByName('Form1Data');
      if (!sheet) {
        throw new Error('Sheet "Form1Data" not found.');
      }
  
      var data = sheet.getDataRange().getValues();
      for (var i =1; i < data.length; i++) {
        var cellValue = data[i][25].toString().trim().toUpperCase();
        var inputValue = fullLotId.toString().trim().toUpperCase();
        if (cellValue === inputValue) {
          // Delete the matching row
          sheet.deleteRow(i + 1); // Adjust for 1-based row index
          return `Row ${i + 1} with Full Lot ID "${fullLotId}" deleted successfully.`;
        }
      }
  
      return `No matching row found for Full Lot ID "${fullLotId}".`;
    }
    catch (error) {
      throw error;
    }
  }
  
  function submitFormData(formObject) {
    try {
      var primarySheet = SpreadsheetApp.openById('14thocGSc9x-eaVVVMITJ3DWSj3SNUR1uhQQrNzDL4wE').getSheetByName('MDv2');
      if (!primarySheet) {
        throw new Error('Sheet "MDv2" not found');
      }
  
      var backupSheet = SpreadsheetApp.openById('1GoVB-vJGFxZlZUDZgRlpMOISSRYI8604xsM2roYpLOk').getSheetByName('MDv2');
      if (!backupSheet) {
        throw new Error('Backup sheet not found');
      }
  
      // Handle timestamp
      formObject['timestamp'] = new Date();
  
      var headers = [
        'timestamp', 'done_by_1', 'done_by_2', 'split_id', 'lot_id', 'lot_id_suffix', 'tool_number', 'kevin_probe_test', 'for_4wire_input', 'date_code', 'date_code_input', 'job_type', 'machine', 'xcut_allowed', 'ipc_class', 'mass_lam', 'resistive', 'continuity', 'test_voltage', 'isolation', 'adjacency_used', 'mfg_qty', 'tested_qty', 'first_passed_qty', 'open', 'short', 'sdp', 'open_circuit_selection', 'il_strip_080', 'chema_180_ol_cm_dm_o', 'chema_477_ol_cm_dm_o', 'ol_lpsm_cure_473', 'ol_cm_osp_473', 'ol_cm_strp_473', 'ol_drl_133', 'outerlayer_open_repair', 'photoprint_080_ol_pp_dev', 'mask_on_pad_finger_repair', 'ol_sm_cure_205', 'plating_nodule_repair', 'ol_cm_strp_178', 'legend_on_pad_hole_repair', 'ol_lpsm_cure_233', 'others_open_repair', 'others_open_reject', 'false_oc_repair', 'short_circuit_selection', 'short_repair', 'photoprint_083_ol_pp_dev', 'scratches_repair', 'photoprint_483_ol_pp_dev', 'chemb_483_ol_cm_strp', 'chemc_483_ol_cm_osp', 'cu_residue_repair', 'chemb_172_ol_cm_strp', 'under_etch_repair', 'chemb_076_ol_cm_strp', 'il_strip_083', 'mis_registration_repair', 'ol_lam_471', 'ol_lam_469', 'ol_drl_132', 'micro_short_repair', 'chemb_083_ol_cm_strp', 'lpsm_083_ol_sm_cure', 'incomplete_solder_strip_repair', 'chemc_256_ol_cm_osp', 'feathering_repair', 'chemc_250_ol_cm_osp', 'incomplete_resist_strip_repair', 'chemb_51_ol_cm_strp', 'npth_short_repair', 'hello_142', 'short_others_repair', 'short_others_reject', 'ol_lam_479', 'ol_pp_dev_479', 'ol_lam_32', 'impedance_fail', 'impedance_fail_high', 'impedance_fail_low', 'final_passed_qty'
      ];
  
      var expectedTypes = {
        // 'timestamp': 'date', 
        'done_by_1': 'string', 
        'done_by_2': 'string',
        'split_id': 'string', 
        'lot_id': 'string', 
        'lot_id_suffix': 'string',
        'tool_number': 'string', 
        'kevin_probe_test': 'string', 
        'for_4wire_input': 'string', 
        'date_code': 'string', 
        'date_code_input': 'string', 
        'job_type': 'string', 
        'machine': 'string', 
        'xcut_allowed': 'string', 
        'ipc_class': 'string', 
        'mass_lam': 'string', 
        'resistive': 'string', 
        'continuity': 'number', 
        'test_voltage': 'number', 
        'isolation': 'number', 
        'adjacency_used': 'string',  
        'mfg_qty': 'number', 
        'tested_qty': 'number', 
        'first_passed_qty': 'number', 
        'open': 'number', 
        'short': 'number', 
        'sdp': 'number', 
        'open_circuit_selection': 'string', 
        'il_strip_080': 'number', 
        'chema_180_ol_cm_dm_o': 'number', 
        'chema_477_ol_cm_dm_o': 'number', 
        'ol_lpsm_cure_473': 'number', 
        'ol_cm_osp_473': 'number', 
        'ol_cm_strp_473': 'number', 
        'ol_drl_133': 'number', 
        'outerlayer_open_repair': 'number', 
        'photoprint_080_ol_pp_dev': 'number', 
        'mask_on_pad_finger_repair': 'number', 
        'ol_sm_cure_205': 'number', 
        'plating_nodule_repair': 'number', 
        'ol_cm_strp_178': 'number', 
        'legend_on_pad_hole_repair': 'number', 
        'ol_lpsm_cure_233': 'number', 
        'others_open_repair': 'number', 
        'others_open_reject': 'number', 
        'false_oc_repair': 'number', 
        'short_circuit_selection': 'string', 
        'short_repair': 'number', 
        'photoprint_083_ol_pp_dev': 'number', 
        'scratches_repair': 'number', 
        'photoprint_483_ol_pp_dev': 'number', 
        'chemb_483_ol_cm_strp': 'number', 
        'chemc_483_ol_cm_osp': 'number', 
        'cu_residue_repair': 'number', 
        'chemb_172_ol_cm_strp': 'number', 
        'under_etch_repair': 'number', 
        'chemb_076_ol_cm_strp': 'number', 
        'il_strip_083': 'number', 
        'mis_registration_repair': 'number', 
        'ol_lam_471': 'number', 
        'ol_lam_469': 'number', 
        'ol_drl_132': 'number', 
        'micro_short_repair': 'number', 
        'chemb_083_ol_cm_strp': 'number', 
        'lpsm_083_ol_sm_cure': 'number', 
        'incomplete_solder_strip_repair': 'number', 
        'chemc_256_ol_cm_osp': 'number', 
        'feathering_repair': 'number', 
        'chemc_250_ol_cm_osp': 'number', 
        'incomplete_resist_strip_repair': 'number', 
        'chemb_51_ol_cm_strp': 'number', 
        'npth_short_repair': 'number', 
        'hello_142': 'number', 
        'short_others_repair': 'number', 
        'short_others_reject': 'number', 
        'ol_lam_479': 'number', 
        'ol_pp_dev_479': 'number', 
        'ol_lam_32': 'number', 
        'impedance_fail': 'string', 
        'impedance_fail_high': 'number',
        'impedance_fail_low': 'number',
        'final_passed_qty': 'number'
      };
  
      var row = [];
      headers.forEach(function(header) {
        var value = formObject[header] || '';
        var type = expectedTypes[header] || 'string';
  
        // Convert based on expected type
        if (type === 'number') {
          value = Number(value);
          if (isNaN(value)) {
            value = '';  // Handle NaN cases if needed
          }
        }
  
        // Special handling for fields like 'lot_id_suffix' to keep parentheses
        else if (header === 'lot_id_suffix' && value.startsWith("(") && value.endsWith(")")) {
          // Adding a single quote before parentheses values
          value = "'" + value; 
        } 
  
        row.push(value);
      });
  
      // Concatenate Split ID, Lot ID and Lot ID Suffix 
      var splitId = formObject['split_id'] || '';
      var lotId = formObject['lot_id'];
      var lotIdSuffix = formObject['lot_id_suffix'] || '';
      var concatenatedId = splitId + lotId + lotIdSuffix; // Concatenate values
      row.push(concatenatedId);
  
      primarySheet.appendRow(row);
  
      backupSheet.appendRow(row);
  
      return 'Form submission successful!';
    } 
    catch (e) {
      Logger.log('Error in submitFormData: ' + e.toString());
      return 'Form submission failed: ' + e.toString();
    }
  }
  