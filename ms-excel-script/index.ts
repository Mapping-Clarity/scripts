// --- Type Definitions for Strict TypeScript ---
interface ProcessRow { 
    [key: string]: string | number | boolean; 
  }
  
  interface MappedResult {
    mappedData: { [key: string]: string };
    normalizedData?: { [key: string]: string };
    confidence?: string | number;
    reasoning?: string;
  }
  
  interface ApiResponse {
    jobId: string;
    async: boolean;
    results?: MappedResult[];
    status?: string;
    message?: string;
  }
  
  // --- HARDCODED ENDPOINTS ---
  const CONFIG = {
    PROCESS_URL: "https://api.mappingclarity.com/v1/process",
    STATUS_BASE_URL: "https://api.mappingclarity.com/v1/jobs",
    SETTINGS_SHEET: "MC_Settings",
    JOB_ID_HEADER: "Job ID (Hidden)",
    CONFIDENCE_HEADER: "Confidence",
    REASONING_HEADER: "Reasoning",
    POLL_INTERVAL_MS: 10000,
    MAX_POLL_ATTEMPTS: 30
  };
  
  async function main(workbook: ExcelScript.Workbook) {
    const activeSheet = workbook.getActiveWorksheet();
    
    // 1. CONFIGURATION LAYER: Get or Create Shadow Sheet
    let configSheet = workbook.getWorksheet(CONFIG.SETTINGS_SHEET);
    
    if (!configSheet) {
      createShadowConfigSheet(workbook);
      console.log(`ACTION REQUIRED: Created [${CONFIG.SETTINGS_SHEET}] tab. Enter credentials and run again.`);
      return;
    }
  
    if (activeSheet.getName() === CONFIG.SETTINGS_SHEET) {
      console.log("Error: Switch to your data sheet before running.");
      return;
    }
  
    // 2. CREDENTIAL GATHERING
    const apiKey: string = configSheet.getRange("B1").getValue()?.toString().trim() || "";
    const pipelineId: string = configSheet.getRange("B2").getValue()?.toString().trim() || "";
  
    if (!apiKey || apiKey === "" || apiKey.includes("Enter")) {
      console.log(`Error: API Key missing in [${CONFIG.SETTINGS_SHEET}] B1.`);
      configSheet.activate();
      return;
    }
  
    // 3. DATA EXTRACTION
    const usedRange = activeSheet.getUsedRange();
    if (!usedRange) {
      console.log("No data found.");
      return;
    }
    
    const data: (string | number | boolean)[][] = usedRange.getValues();
    const headers: string[] = data[0].map(h => String(h));
    
    const jobIdColIdx = headers.findIndex(h => h.toLowerCase() === CONFIG.JOB_ID_HEADER.toLowerCase());
  
    const rowsToProcess: number[] = [];
    const payloadRows: ProcessRow[] = [];
  
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (jobIdColIdx === -1 || !row[jobIdColIdx]) {
        const rowObj: ProcessRow = {};
        headers.forEach((h, idx) => {
          const hLow = h.toLowerCase();
          if (h && ![CONFIG.JOB_ID_HEADER.toLowerCase(), CONFIG.CONFIDENCE_HEADER.toLowerCase(), CONFIG.REASONING_HEADER.toLowerCase()].includes(hLow)) {
            rowObj[h] = String(row[idx] || "").replace(/"/g, "'");
          }
        });
        if (Object.keys(rowObj).length > 0) {
          payloadRows.push(rowObj);
          rowsToProcess.push(i + 1); 
        }
      }
    }
  
    if (payloadRows.length === 0) {
      console.log("No new data to process.");
      return;
    }
  
    // 4. API EXECUTION
    console.log(`Sending ${payloadRows.length} rows to Mapping Clarity...`);
    
    try {
      const response = await fetch(CONFIG.PROCESS_URL, {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json', 
          'x-api-key': apiKey 
        },
        body: JSON.stringify({
          pipelineId: pipelineId,
          originalFilename: workbook.getName() + ".csv",
          rows: payloadRows
        })
      });
  
      const resData: ApiResponse = await response.json();
  
      if (response.status === 200 || response.status === 202) {
        if (resData.async) {
          console.log(`Job submitted (ID: ${resData.jobId}). Waiting for results...`);
          await pollAndProcess(activeSheet, apiKey, resData.jobId, rowsToProcess);
        } else {
          writeResults(activeSheet, resData, rowsToProcess, resData.jobId);
        }
      } else {
        console.log(`API Error ${response.status}: ${resData.message || "Unknown error"}`);
      }
    } catch (error) {
      console.log("Connection Error: " + error.toString());
    }
  }
  
  function createShadowConfigSheet(workbook: ExcelScript.Workbook) {
    const settingsSheet = workbook.addWorksheet(CONFIG.SETTINGS_SHEET);
    settingsSheet.getRange("A1:A2").setValues([["API Key"], ["Pipeline ID"]]);
    settingsSheet.getRange("A1:A2").getFormat().getFont().setBold(true);
    
    const inputRange = settingsSheet.getRange("B1:B2");
    inputRange.setValues([["Enter API Key Here"], ["Enter Pipeline ID Here"]]);
    inputRange.getFormat().getFill().setColor("FFFF00"); 
    
    settingsSheet.getRange("A1:B2").getFormat().autofitColumns();
    settingsSheet.activate();
  }
  
  function normalize(key: string): string {
    if (!key) return key;
    return key.split('_').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
  }
  
  /**
   * Polling logic using HARDCODED status and results URLs
   */
  async function pollAndProcess(sheet: ExcelScript.Worksheet, apiKey: string, jobId: string, rows: number[]) {
    const statusUrl = `${CONFIG.STATUS_BASE_URL}/${jobId}`;
    const resultsUrl = `${CONFIG.STATUS_BASE_URL}/${jobId}/results`;
  
    for (let i = 0; i < CONFIG.MAX_POLL_ATTEMPTS; i++) {
      await new Promise(r => setTimeout(r, CONFIG.POLL_INTERVAL_MS));
      
      try {
        const res = await fetch(statusUrl, { headers: { 'x-api-key': apiKey } });
        const data: ApiResponse = await res.json();
        
        if (data.status === "completed") {
          console.log("Job complete. Fetching final results...");
          const resultsRes = await fetch(resultsUrl, { headers: { 'x-api-key': apiKey } });
          const resultsData: ApiResponse = await resultsRes.json();
          writeResults(sheet, resultsData, rows, jobId);
          return;
        } else if (data.status === "failed") {
          console.log("API reported a failed status.");
          return;
        }
      } catch (e) {
        console.log("Polling error: " + e.toString());
      }
    }
    console.log("Polling timed out.");
  }
  
  function writeResults(sheet: ExcelScript.Worksheet, resData: ApiResponse, rows: number[], jobId: string) {
    const results = resData.results;
    if (!results) return;
  
    const usedRange = sheet.getUsedRange();
    if (!usedRange) return;
  
    let headers: string[] = sheet.getRangeByIndexes(0, 0, 1, usedRange.getColumnCount()).getValues()[0].map(h => String(h));
  
    // 1. Discover Mapped Data and Normalize Keys
    results.forEach(res => {
      if (res.mappedData) {
        res.normalizedData = {};
        Object.keys(res.mappedData).forEach(key => {
          const normKey = normalize(key);
          res.normalizedData![normKey] = res.mappedData[key];
          getOrAddCol(sheet, headers, normKey);
        });
      }
    });
  
    // 2. Add Metadata headers in sequence
    const jobIdIdx = getOrAddCol(sheet, headers, CONFIG.JOB_ID_HEADER);
    getOrAddCol(sheet, headers, CONFIG.CONFIDENCE_HEADER);
    getOrAddCol(sheet, headers, CONFIG.REASONING_HEADER);
  
    // 3. Final Header Map for Case-Insensitive write
    const finalUsedRange = sheet.getUsedRange();
    const finalHeaders = sheet.getRangeByIndexes(0, 0, 1, finalUsedRange.getColumnCount()).getValues()[0].map(h => String(h));
    const colMap: { [key: string]: number } = {};
    finalHeaders.forEach((h, i) => colMap[h.toLowerCase()] = i);
  
    // 4. Batch Process Results
    results.forEach((res, i) => {
      const rowNum = rows[i];
      const range = sheet.getRangeByIndexes(rowNum - 1, 0, 1, finalHeaders.length);
      const vals = range.getValues()[0];
  
      if (res.normalizedData) {
        Object.keys(res.normalizedData).forEach(k => {
          const idx = colMap[k.toLowerCase()];
          if (idx !== undefined) vals[idx] = res.normalizedData![k];
        });
      }
  
      vals[colMap[CONFIG.JOB_ID_HEADER.toLowerCase()]] = jobId;
      vals[colMap[CONFIG.CONFIDENCE_HEADER.toLowerCase()]] = res.confidence || "";
      vals[colMap[CONFIG.REASONING_HEADER.toLowerCase()]] = res.reasoning || "";
      
      range.setValues([vals]);
    });
  
    // 5. Hide Job ID Column
    if (jobIdIdx !== -1) {
      sheet.getRangeByIndexes(0, jobIdIdx, 1, 1).setColumnHidden(true);
    }
    
    console.log(`Success: Processed ${results.length} rows.`);
  }
  
  function getOrAddCol(sheet: ExcelScript.Worksheet, headers: string[], name: string): number {
    const searchName = name.toLowerCase();
    const idx = headers.findIndex(h => h.toLowerCase() === searchName);
    
    if (idx !== -1) return idx;
  
    const newIdx = headers.length;
    sheet.getRangeByIndexes(0, newIdx, 1, 1).setValue(name);
    headers.push(name);
    return newIdx;
  }