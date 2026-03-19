// --- Type Definitions ---
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

  // 1. CONFIGURATION LAYER
  let configSheet = workbook.getWorksheet(CONFIG.SETTINGS_SHEET);
  if (!configSheet) {
    createShadowConfigSheet(workbook);
    console.log("Config sheet created. Please add credentials and run again.");
    return;
  }

  if (activeSheet.getName() === CONFIG.SETTINGS_SHEET) {
    console.log("Error: Please switch to your data sheet before running.");
    return;
  }

  const apiKey: string = configSheet.getRange("B1").getValue()?.toString().trim() || "";
  const pipelineId: string = configSheet.getRange("B2").getValue()?.toString().trim() || "";

  if (!apiKey || apiKey === "" || apiKey.includes("Enter")) {
    configSheet.activate();
    return;
  }

  // 2. DATA EXTRACTION
  const usedRange = activeSheet.getUsedRange();
  if (!usedRange) return;

  const data = usedRange.getValues();
  const headers = data[0].map(h => String(h));
  const jobIdColIdx = headers.findIndex(h => h.toLowerCase() === CONFIG.JOB_ID_HEADER.toLowerCase());

  const rowsToProcess: number[] = [];
  const payloadRows: ProcessRow[] = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (jobIdColIdx === -1 || !row[jobIdColIdx]) {
      const rowObj: ProcessRow = {};
      headers.forEach((h, idx) => {
        const hLow = h.toLowerCase();
        // Skip metadata columns from the input payload
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
    console.log("All rows processed.");
    return;
  }

  // 3. API EXECUTION
  console.log(`Sending ${payloadRows.length} rows to Mapping Clarity...`);

  try {
    const response = await fetch(CONFIG.PROCESS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey },
      body: JSON.stringify({
        pipelineId: pipelineId,
        rows: payloadRows
      })
    });

    const resData: ApiResponse = await response.json();

    if (response.status === 200 || response.status === 202) {
      if (resData.async) {
        await pollAndProcess(activeSheet, apiKey, resData.jobId, rowsToProcess);
      } else if (resData.results) {
        writeResults(activeSheet, resData, rowsToProcess, resData.jobId);
      }
    } else {
      console.log(`API Error (${response.status}): ${resData.message}`);
    }
  } catch (error) {
    console.log(`Connection Error: ${String(error)}`);
  }
}

/**
 * Polling Helper Function
 */
async function fetchStatusHelper(url: string, key: string): Promise<ApiResponse | null> {
  try {
    const res = await fetch(url, { headers: { 'x-api-key': key } });
    const data: ApiResponse = await res.json();
    return data;
  } catch (e) {
    return null;
  }
}

async function pollAndProcess(sheet: ExcelScript.Worksheet, apiKey: string, jobId: string, rows: number[]) {
  const statusUrl = `${CONFIG.STATUS_BASE_URL}/${jobId}`;
  const resultsUrl = `${CONFIG.STATUS_BASE_URL}/${jobId}/results`;

  for (let i = 0; i < CONFIG.MAX_POLL_ATTEMPTS; i++) {
    await new Promise(resolve => setTimeout(resolve, CONFIG.POLL_INTERVAL_MS));

    const data = await fetchStatusHelper(statusUrl, apiKey);

    if (data && data.status === "completed") {
      const resultsRes = await fetch(resultsUrl, { headers: { 'x-api-key': apiKey } });
      const resultsData: ApiResponse = await resultsRes.json();
      writeResults(sheet, resultsData, rows, jobId);
      return;
    } else if (data && data.status === "failed") {
      console.log("Job failed.");
      return;
    }
  }
  console.log("Polling timed out.");
}

function normalize(key: string): string {
  if (!key) return key;
  return key.split('_').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
}

/**
 * Results Writing:
 * Implements Case-Insensitive Header Mapping to prevent duplication.
 */
function writeResults(sheet: ExcelScript.Worksheet, resData: ApiResponse, rows: number[], jobId: string) {
  const results = resData.results;
  if (!results) return;

  const usedRange = sheet.getUsedRange();
  if (!usedRange) return;

  // 1. Map Existing Headers to prevent casing duplicates (e.g. Account Name vs account_name)
  let currentHeaders: string[] = sheet.getRangeByIndexes(0, 0, 1, usedRange.getColumnCount()).getValues()[0].map(h => String(h));
  let headerMap: { [key: string]: string } = {};
  currentHeaders.forEach(h => {
    headerMap[h.toLowerCase()] = h;
  });

  // 2. Discover and Setup Columns
  results.forEach(res => {
    if (res.mappedData) {
      res.normalizedData = {};
      Object.keys(res.mappedData).forEach(key => {
        const exactKeyLow = key.toLowerCase();
        const normName = normalize(key);
        const lowName = normName.toLowerCase();

        // If a version of this header exists already (exact match or normalized), use it.
        let finalHeaderName: string;
        if (headerMap[exactKeyLow]) {
          finalHeaderName = headerMap[exactKeyLow];
        } else if (headerMap[lowName]) {
          finalHeaderName = headerMap[lowName];
        } else {
          finalHeaderName = normName;
        }

        res.normalizedData![finalHeaderName] = res.mappedData[key];
        getOrAddCol(sheet, currentHeaders, finalHeaderName);

        // Update local map if we created a new column
        if (!headerMap[lowName]) { headerMap[lowName] = finalHeaderName; }
        if (!headerMap[exactKeyLow]) { headerMap[exactKeyLow] = finalHeaderName; }
      });
    }
  });

  // 3. Add Fixed Metadata Headers
  const jobIdIdx = getOrAddCol(sheet, currentHeaders, CONFIG.JOB_ID_HEADER);
  getOrAddCol(sheet, currentHeaders, CONFIG.CONFIDENCE_HEADER);
  getOrAddCol(sheet, currentHeaders, CONFIG.REASONING_HEADER);

  // 4. Get Final Column Index Mapping
  const finalHeaders = sheet.getRangeByIndexes(0, 0, 1, sheet.getUsedRange().getColumnCount()).getValues()[0].map(h => String(h));
  const colMap: { [key: string]: number } = {};
  finalHeaders.forEach((h, i) => colMap[h.toLowerCase()] = i);

  // 5. Bulk Write values to rows
  results.forEach((res, i) => {
    const rowNum = rows[i];
    const range = sheet.getRangeByIndexes(rowNum - 1, 0, 1, finalHeaders.length);
    const vals = range.getValues()[0];

    // Write Mapped Data
    if (res.normalizedData) {
      Object.keys(res.normalizedData).forEach(headerKey => {
        const idx = colMap[headerKey.toLowerCase()];
        if (idx !== undefined) vals[idx] = res.normalizedData![headerKey];
      });
    }

    // Write Metadata
    vals[colMap[CONFIG.JOB_ID_HEADER.toLowerCase()]] = jobId;
    vals[colMap[CONFIG.CONFIDENCE_HEADER.toLowerCase()]] = res.confidence || "";
    vals[colMap[CONFIG.REASONING_HEADER.toLowerCase()]] = res.reasoning || "";

    range.setValues([vals]);
  });

  // 6. Final Task: Hide the Job ID column
  if (jobIdIdx !== -1) {
    sheet.getRangeByIndexes(0, jobIdIdx, 1, 1).setColumnHidden(true);
  }

  console.log("Processing Complete.");
}

/**
 * Case-Insensitive Column Index Finder
 */
function getOrAddCol(sheet: ExcelScript.Worksheet, headers: string[], name: string): number {
  const searchName = name.toLowerCase();
  const idx = headers.findIndex(h => h.toLowerCase() === searchName);

  if (idx !== -1) return idx;

  // Not found: Create new column
  const newIdx = headers.length;
  sheet.getRangeByIndexes(0, newIdx, 1, 1).setValue(name);
  headers.push(name);
  return newIdx;
}

/**
 * Setup Utility: Create Configuration Tab
 */
function createShadowConfigSheet(workbook: ExcelScript.Workbook) {
  const settingsSheet = workbook.addWorksheet(CONFIG.SETTINGS_SHEET);
  settingsSheet.getRange("A1:A2").setValues([["API Key"], ["Pipeline ID"]]);
  settingsSheet.getRange("A1:A2").getFormat().getFont().setBold(true);
  settingsSheet.getRange("B1:B2").setValues([["Enter API Key Here"], ["Enter Pipeline ID Here"]]);
  settingsSheet.getRange("B1:B2").getFormat().getFill().setColor("FFFF00");
  settingsSheet.getRange("A1:B2").getFormat().autofitColumns();
  settingsSheet.activate();
}
