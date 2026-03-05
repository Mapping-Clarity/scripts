/**
 * Mapping Clarity - Optimized Google Apps Script Version
 */

// ===== CONFIGURATION =====
const CONFIG = {
    API_URL: "https://api.mappingclarity.com/v1/process",
    JOB_ID_HEADER: "Job ID (Hidden)",
    CONFIDENCE_HEADER: "Confidence",
    REASONING_HEADER: "Reasoning",
    POLL_INTERVAL_MS: 10000, 
    MAX_POLL_ATTEMPTS: 30    
  };
  
  function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Mapping Clarity')
      .addItem('Run Process', 'runMappingClarity')
      .addSeparator()
      .addItem('Clear Saved API Key', 'clearCredentials')
      .addToUi();
  }
  
  /**
   * Normalizes keys: "first_name" -> "First Name"
   */
  function normalizeKey(key) {
    if (!key) return key;
    return key.split('_')
              .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
              .join(' ');
  }
  
  function runMappingClarity() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const ui = SpreadsheetApp.getUi();
    
    let fileName = ss.getName();
    if (!fileName.toLowerCase().endsWith(".csv")) {
      fileName += ".csv";
    }
  
    const userProps = PropertiesService.getUserProperties();
    let apiKey = userProps.getProperty('MC_API_KEY');
    let pipelineId = userProps.getProperty('MC_PIPELINE_ID');
  
    if (!apiKey) {
      const res = ui.prompt('Setup', 'Enter your API Key:', ui.ButtonSet.OK_CANCEL);
      if (res.getSelectedButton() !== ui.Button.OK) return;
      apiKey = res.getResponseText().trim();
      userProps.setProperty('MC_API_KEY', apiKey);
    }
  
    if (!pipelineId) {
      const res = ui.prompt('Setup', 'Enter your Pipeline ID:', ui.ButtonSet.OK_CANCEL);
      if (res.getSelectedButton() !== ui.Button.OK) return;
      pipelineId = res.getResponseText().trim();
      userProps.setProperty('MC_PIPELINE_ID', pipelineId);
    }
  
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      ui.alert("No data found.");
      return;
    }
  
    const headers = data[0];
    // Case-insensitive search for Job ID index
    const jobIdColIdx = headers.findIndex(h => h && h.toString().toLowerCase() === CONFIG.JOB_ID_HEADER.toLowerCase());
    
    const rowsToProcess = [];
    const payloadRows = [];
  
    for (let i = 1; i < data.length; i++) {
      const currentRow = data[i];
      if (jobIdColIdx === -1 || !currentRow[jobIdColIdx]) {
        const rowObject = {};
        headers.forEach((header, index) => {
          if (header) {
            const hLower = header.toString().toLowerCase();
            // Don't send metadata columns to the API
            if (![CONFIG.JOB_ID_HEADER.toLowerCase(), CONFIG.CONFIDENCE_HEADER.toLowerCase(), CONFIG.REASONING_HEADER.toLowerCase()].includes(hLower)) {
              rowObject[header] = String(currentRow[index] || "").replace(/"/g, "'");
            }
          }
        });
        payloadRows.push(rowObject);
        rowsToProcess.push(i + 1); 
      }
    }
  
    if (payloadRows.length === 0) {
      ui.alert("No new rows found to process.");
      return;
    }
  
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey },
      payload: JSON.stringify({
        pipelineId: pipelineId,
        originalFilename: fileName, 
        rows: payloadRows
      }),
      muteHttpExceptions: true
    };
  
    try {
      const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
      const statusCode = response.getResponseCode();
      const responseData = JSON.parse(response.getContentText());
  
      if (statusCode === 200 || statusCode === 202) {
        const jobId = responseData.jobId;
        if (responseData.async === true) {
          pollAndProcessBatch(sheet, apiKey, jobId, rowsToProcess);
        } else {
          writeBatchResults(sheet, responseData, rowsToProcess, jobId);
        }
      } else {
        ui.alert(`API Error ${statusCode}: ${response.getContentText()}`);
      }
    } catch (e) {
      ui.alert("Network Error: " + e.toString());
    }
  }
  
  function pollAndProcessBatch(sheet, apiKey, jobId, rowsToProcess) {
    const statusUrl = `https://api.mappingclarity.com/v1/jobs/${jobId}`;
    const resultsUrl = `${statusUrl}/results`;
    
    for (let attempt = 0; attempt < CONFIG.MAX_POLL_ATTEMPTS; attempt++) {
      Utilities.sleep(CONFIG.POLL_INTERVAL_MS);
      const response = UrlFetchApp.fetch(statusUrl, { headers: { 'x-api-key': apiKey }, muteHttpExceptions: true });
      const statusData = JSON.parse(response.getContentText());
      
      if (statusData.status === "completed") {
        const resResponse = UrlFetchApp.fetch(resultsUrl, { headers: { 'x-api-key': apiKey }, muteHttpExceptions: true });
        const resultsData = JSON.parse(resResponse.getContentText());
        writeBatchResults(sheet, resultsData, rowsToProcess, jobId);
        return;
      } else if (statusData.status === "failed") {
        SpreadsheetApp.getUi().alert("API Job failed.");
        return;
      }
    }
    SpreadsheetApp.getUi().alert("Job timed out.");
  }
  
  function writeBatchResults(sheet, responseData, rowsToProcess, jobId) {
    const results = responseData.results;
    if (!results || results.length === 0) return;
  
    let currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 1. Discover Mapped Data and Normalize Keys
    results.forEach(res => {
      if (res.mappedData) {
        const normalizedData = {};
        Object.keys(res.mappedData).forEach(key => {
          const normKey = normalizeKey(key);
          normalizedData[normKey] = res.mappedData[key];
          getOrAddColumn(sheet, currentHeaders, normKey);
        });
        res.normalizedMappedData = normalizedData;
      }
    });
  
    // 2. Add Fixed Columns (Job ID first, then Confidence/Reasoning)
    const jobIdIdxForHiding = getOrAddColumn(sheet, currentHeaders, CONFIG.JOB_ID_HEADER);
    getOrAddColumn(sheet, currentHeaders, CONFIG.CONFIDENCE_HEADER);
    getOrAddColumn(sheet, currentHeaders, CONFIG.REASONING_HEADER);
  
    // 3. Setup Write Block
    const finalHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const startRow = Math.min(...rowsToProcess);
    const endRow = Math.max(...rowsToProcess);
    const numRows = endRow - startRow + 1;
    const numCols = finalHeaders.length;
    
    const range = sheet.getRange(startRow, 1, numRows, numCols);
    const valueBlock = range.getValues();
  
    const colMap = {};
    finalHeaders.forEach((h, i) => {
      if (h) colMap[h.toString().toLowerCase()] = i;
    });
  
    // 4. Fill Write Block
    results.forEach((result, index) => {
      const targetRowNumber = rowsToProcess[index];
      const relativeRowIdx = targetRowNumber - startRow;
      const rowArray = valueBlock[relativeRowIdx];
  
      if (result.normalizedMappedData) {
        Object.keys(result.normalizedMappedData).forEach(normKey => {
          const idx = colMap[normKey.toLowerCase()];
          if (idx !== undefined) rowArray[idx] = result.normalizedMappedData[normKey];
        });
      }
  
      const jobIdIdx = colMap[CONFIG.JOB_ID_HEADER.toLowerCase()];
      const confIdx = colMap[CONFIG.CONFIDENCE_HEADER.toLowerCase()];
      const reasonIdx = colMap[CONFIG.REASONING_HEADER.toLowerCase()];
  
      if (jobIdIdx !== undefined) rowArray[jobIdIdx] = jobId;
      if (confIdx !== undefined && result.confidence !== undefined) rowArray[confIdx] = result.confidence;
      if (reasonIdx !== undefined && result.reasoning !== undefined) rowArray[reasonIdx] = result.reasoning;
    });
  
    // 5. Bulk Write to Sheet
    range.setValues(valueBlock);
  
    // 6. HIDE THE JOB ID COLUMN
    sheet.hideColumns(jobIdIdxForHiding + 1);
  
    SpreadsheetApp.getUi().alert("Successfully processed " + results.length + " rows.");
  }
  
  /**
   * CASE-INSENSITIVE Column Helper
   */
  function getOrAddColumn(sheet, headers, hName) {
    const searchName = hName.toString().toLowerCase();
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && headers[i].toString().toLowerCase() === searchName) {
        return i; 
      }
    }
    const newIdx = headers.length;
    sheet.getRange(1, newIdx + 1).setValue(hName);
    headers.push(hName); 
    return newIdx;
  }
  
  function clearCredentials() {
    PropertiesService.getUserProperties().deleteProperty('MC_API_KEY');
    PropertiesService.getUserProperties().deleteProperty('MC_PIPELINE_ID');
    SpreadsheetApp.getUi().alert("API Key and Pipeline ID cleared.");
  }