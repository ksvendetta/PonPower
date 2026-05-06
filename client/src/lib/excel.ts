import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

function colNumberToLetters(n: number): string {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export interface Terminal {
  rowIndex: number; // 0-based
  waldoId: string;
  terminalName: string;
  cableId: string;
  powerTestStrand: number;
  otdrTestStrand: string;
  totalStrands: number;
  testpQty: number | string;
  testpaQty: number | string;
  staggeredPort?: number;
  staggeredStrand?: number;
}

export interface ParsedWorkbook {
  fileBuffer: ArrayBuffer;
  sheetName: string;
  headerRowIndex: number; // 0-based
  powerStrandColIndex: number; // 0-based
  totalStrandsColIndex: number; // 0-based
  terminals: Terminal[];
}

export interface JsonTemplate {
  Label: {
    tester: string;
    wireCenterClli: string;
    cfas: string;
    aLoc: string;
    zLoc: string;
    dateTested: string;
    model: string;
    serialNumber: string;
    version: string;
    calibrationDate: string;
    strandPowerThreshold: number;
    cableId: string;
    Tested: TestedItem[];
  };
}

export interface TestedItem {
  strand: number;
  passFail: string;
  linkMeasurements: LinkMeasurement[];
}

export interface LinkMeasurement {
  wavelength: number;
  totalLoss: number;
}

const DEFAULT_TEMPLATE: JsonTemplate = {
  Label: {
    tester: "kl131s",
    wireCenterClli: "",
    cfas: "",
    aLoc: "",
    zLoc: "",
    dateTested: new Date().toISOString().split('T')[0],
    model: "PM-1v2-2X-VFL",
    serialNumber: "1406971",
    version: "1.0",
    calibrationDate: "0001-01-01",
    strandPowerThreshold: -24,
    cableId: "",
    Tested: []
  }
};

export async function parseExcelFile(file: File): Promise<{ cableId: string | null; strands: number[], cfas: string | null }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

        if (jsonData.length === 0) {
          resolve({ cableId: null, strands: [], cfas: null });
          return;
        }
        
        // ---------------------------------------------------------
        // Look for CFAS in range W2:AM4 (Indices: Rows 1-3, Cols 22-40)
        // User stated X2, but data suggests it might be in W (index 22).
        // We scan a wider range (Cols 20-40) in Rows 1-3 to be robust.
        // ---------------------------------------------------------
        let foundCfas: string | null = null;
        console.log("Searching for CFAS in expanded range Cols 20-40, Rows 1-3");

        // Row loop: 1 to 3 (Rows 2-4)
        for (let r = 1; r <= 3; r++) {
           if (jsonData.length <= r) break;
           const row = jsonData[r];
           if (!row) continue;
           
           // Col loop: 20 to 40 (Cols U to AN) - wider net to catch W or X
           for (let c = 20; c <= 40; c++) {
              if (row.length <= c) continue;
              const cell = row[c];
              
              // Check for string or number
              if (cell !== null && cell !== undefined && cell !== '') {
                 const strVal = String(cell).trim();
                 
                 // Filter out common labels if they appear in this area (unlikely but safe)
                 // CFAS is typically a code like "A058G7R"
                 if (strVal.length > 0 && strVal !== "-" && !strVal.includes("TEST SHEET")) {
                    foundCfas = strVal;
                    console.log(`Found CFAS candidate in scan at [${r}, ${c}]: ${foundCfas}`);
                    break;
                 }
              }
           }
           if (foundCfas) break;
        }

        // ---------------------------------------------------------
        // PRIORITY STRATEGY: Look for "Power Test Strand" explicitly
        // ---------------------------------------------------------
        let powerStrandColIndex = -1;
        let powerStrandHeaderRow = -1;

        // Search for the specific header "Power Test Strand"
        for (let i = 0; i < Math.min(50, jsonData.length); i++) {
          const row = jsonData[i];
          if (!row) continue;
          
          for (let j = 0; j < row.length; j++) {
            const cell = row[j];
            if (typeof cell === 'string') {
              const lower = cell.toLowerCase().trim();
              if (lower.includes("power test strand")) {
                powerStrandColIndex = j;
                powerStrandHeaderRow = i;
                console.log(`Found PRIORITY header 'Power Test Strand' at Row ${i}, Col ${j}`);
                break;
              }
            }
          }
          if (powerStrandColIndex !== -1) break;
        }

        // If we found the specific column, iterate down from there
        if (powerStrandColIndex !== -1) {
           const strands: number[] = [];
           
           // Extract strands from the Power Test Strand column
           for (let i = powerStrandHeaderRow + 1; i < jsonData.length; i++) {
             const row = jsonData[i];
             if (!row) continue;
             
             const cell = row[powerStrandColIndex];
             if (cell !== undefined && cell !== null && cell !== '') {
               const val = parseInt(cell);
               // Check if it's a valid strand number (and not another header or label)
               if (!isNaN(val) && val > 0) {
                 strands.push(val);
               }
             }
           }

           // ---------------------------------------------------------
           // Look for Cable ID specifically around the header row
           // ---------------------------------------------------------
           let foundCableId: string | null = null;
           
           // Search window: 5 rows before and 5 rows after the Power Strand Header
           const startRow = Math.max(0, powerStrandHeaderRow - 5);
           const endRow = Math.min(jsonData.length, powerStrandHeaderRow + 5);

           console.log(`Searching for Cable ID between rows ${startRow} and ${endRow}`);

           for (let r = startRow; r < endRow; r++) {
             const row = jsonData[r];
             if (!row) continue;
             
             for (let c = 0; c < row.length; c++) {
               const cell = row[c];
               if (typeof cell === 'string') {
                 const val = cell.toLowerCase().trim();
                 // Check for "Cable ID" label
                 if ((val.includes('cable') && val.includes('id')) || val === 'cable' || val === 'cableid') {
                    console.log(`Found Cable ID candidate label at [${r}, ${c}]: ${cell}`);
                    
                    // Strategy 1: Value is in the same cell (e.g. "Cable ID: PON915B")
                    const parts = cell.split(/[:=]/); // Split by : or =
                    if (parts.length > 1) {
                      const potentialValue = parts[1].trim();
                      if (potentialValue.length > 0) {
                        foundCableId = potentialValue;
                        console.log(`-> Extracted from same cell: ${foundCableId}`);
                      }
                    } 
                    
                    // Strategy 2: Check cell to the right (Col + 1)
                    if (!foundCableId && row[c+1] && String(row[c+1]).trim()) {
                      foundCableId = String(row[c+1]).trim();
                      console.log(`-> Extracted from right cell [${r}, ${c+1}]: ${foundCableId}`);
                    }

                    // Strategy 3: Check cell below (Row + 1)
                    if (!foundCableId && jsonData[r+1] && jsonData[r+1][c] && String(jsonData[r+1][c]).trim()) {
                       foundCableId = String(jsonData[r+1][c]).trim();
                       console.log(`-> Extracted from cell below [${r+1}, ${c}]: ${foundCableId}`);
                    }
                 }
               }
             }
             if (foundCableId) break;
           }
           
           // Fallback: If user explicitly mentioned N23 (index 13, row 22 if 0-indexed), let's peek there if we failed
           if (!foundCableId && jsonData[22] && jsonData[22][13]) {
              console.log("Checking specific coordinate N23 (approx [22,13])");
              // This is a blind guess based on user hint, strictly as fallback
              const val = jsonData[22][13];
              if (typeof val === 'string' || typeof val === 'number') {
                 // Only use if it looks like an ID (alphanumeric)
                 foundCableId = String(val).trim();
                 console.log(`-> Extracted from N23 fallback: ${foundCableId}`);
              }
           }

           const uniqueStrands = Array.from(new Set(strands));
           console.log(`Priority Extraction complete. CableID: ${foundCableId}, Strands: ${uniqueStrands.length}, CFAS: ${foundCfas}`);
           resolve({ cableId: foundCableId, strands: uniqueStrands, cfas: foundCfas });
           return;
        }

        // ---------------------------------------------------------
        // FALLBACK STRATEGY: Standard Header Search
        // ---------------------------------------------------------
        let cableIdIndex = -1;
        let strandIndex = -1;
        let headerRowIndex = -1;

        for (let i = 0; i < Math.min(20, jsonData.length); i++) {
          const row = jsonData[i];
          if (!row) continue;

          row.forEach((cell: any, index: number) => {
            if (typeof cell === 'string') {
              const lower = cell.toLowerCase().trim();
              // Check for Cable ID synonyms
              if (
                (lower.includes('cable') && lower.includes('id')) || 
                lower === 'cable' || 
                lower === 'cableid' ||
                lower === 'cable_id'
              ) {
                cableIdIndex = index;
              }
              
              // Check for Strand synonyms
              if (
                lower.includes('strand') || 
                lower.includes('fiber') || 
                lower === 'core' || 
                lower === 'no' || 
                lower === '#' ||
                lower === 'position'
              ) {
                strandIndex = index;
              }
            }
          });
          
          // If we found at least one relevant column, assume this is the header
          if (cableIdIndex !== -1 || strandIndex !== -1) {
            headerRowIndex = i;
            console.log(`Found headers at row ${i}: CableIndex=${cableIdIndex}, StrandIndex=${strandIndex}`);
            break;
          }
        }

        const strands: number[] = [];
        let foundCableId: string | null = null;

        if (headerRowIndex !== -1) {
          // Header based extraction
          for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            if (cableIdIndex !== -1 && row[cableIdIndex]) {
              if (!foundCableId) foundCableId = String(row[cableIdIndex]).trim();
            }

            if (strandIndex !== -1 && row[strandIndex] !== undefined) {
              const val = parseInt(row[strandIndex]);
              if (!isNaN(val)) strands.push(val);
            }
          }
        } 
        
        // ---------------------------------------------------------
        // FINAL FALLBACK: Numeric Column Search
        // ---------------------------------------------------------
        if (strands.length === 0) {
          console.log("No strands found via headers, attempting numeric fallback...");
          
          // Find a column that has mostly numbers
          const columnScores = new Map<number, number>();
          
          for (let i = 0; i < Math.min(100, jsonData.length); i++) {
            const row = jsonData[i];
            if (!row) continue;
            row.forEach((cell: any, index: number) => {
              if (typeof cell === 'number' || (typeof cell === 'string' && !isNaN(parseInt(cell)))) {
                const val = parseInt(cell as string);
                // Simple heuristic: strands are usually 1-1000
                if (val > 0 && val < 1000) {
                   columnScores.set(index, (columnScores.get(index) || 0) + 1);
                }
              }
            });
          }

          // Find column with most numeric matches
          let bestCol = -1;
          let maxScore = 0;
          columnScores.forEach((score, col) => {
            if (score > maxScore) {
              maxScore = score;
              bestCol = col;
            }
          });

          if (bestCol !== -1 && maxScore > 5) { // Threshold to avoid noise
            console.log(`Fallback: Using column ${bestCol} as Strand column`);
            for (let i = 0; i < jsonData.length; i++) {
               const row = jsonData[i];
               if (row && row[bestCol] !== undefined) {
                 const val = parseInt(row[bestCol]);
                 if (!isNaN(val) && val > 0) strands.push(val);
               }
            }
          }
        }

        // Deduplicate strands and sort
        const uniqueStrands = Array.from(new Set(strands));
        
        console.log(`Extraction complete. CableID: ${foundCableId}, Strands: ${uniqueStrands.length}, CFAS: ${foundCfas}`);
        resolve({ cableId: foundCableId, strands: uniqueStrands, cfas: foundCfas });

      } catch (err) {
        console.error("Excel parsing error:", err);
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

export function generateReport(
  cableId: string,
  strands: number[],
  averageLoss: number,
  wireCenterClli: string,
  cfas: string,
  baseTemplate: JsonTemplate = DEFAULT_TEMPLATE
): JsonTemplate {
  const newReport = JSON.parse(JSON.stringify(baseTemplate)); // Deep copy
  newReport.Label.cableId = cableId;
  newReport.Label.dateTested = new Date().toISOString().split('T')[0];
  newReport.Label.wireCenterClli = wireCenterClli;
  newReport.Label.cfas = cfas;
  newReport.Label.aLoc = wireCenterClli ? `${wireCenterClli}PFP` : "";
  newReport.Label.zLoc = wireCenterClli ? `${wireCenterClli}PFP` : "";

  newReport.Label.Tested = strands.map(strand => {
    // Randomize loss: average +/- 0.5
    // Ensure we handle negative numbers correctly (loss is usually negative dB)
    const randomOffset = (Math.random() * 1) - 0.5;
    const randomizedLoss = Number((averageLoss + randomOffset).toFixed(2));

    return {
      strand: strand,
      passFail: "pass",
      linkMeasurements: [
        {
          wavelength: 1310,
          totalLoss: 0
        },
        {
          wavelength: 1550,
          totalLoss: randomizedLoss
        }
      ]
    };
  });

  return newReport;
}

export async function parseTerminals(file: File): Promise<ParsedWorkbook> {
  const arrayBuffer = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer.slice(0));
  const ws = wb.worksheets[0];

  const cellText = (r: number, c: number): string => {
    const v = ws.getCell(r, c).value;
    if (v === null || v === undefined) return '';
    if (typeof v === 'object' && 'richText' in (v as any)) {
      return ((v as any).richText as Array<{ text: string }>).map((p) => p.text).join('');
    }
    if (typeof v === 'object' && 'text' in (v as any)) {
      return String((v as any).text);
    }
    return String(v);
  };

  const cellNumber = (r: number, c: number): number | null => {
    const v = ws.getCell(r, c).value;
    if (typeof v === 'number') return v;
    if (typeof v === 'string') {
      const n = parseInt(v);
      return isNaN(n) ? null : n;
    }
    return null;
  };

  let headerRowIndex = -1;
  let powerStrandColIndex = -1;
  let waldoColIndex = -1;
  let terminalColIndex = -1;
  let cableIdColIndex = -1;
  let otdrColIndex = -1;
  let totalStrandsColIndex = -1;
  let testpColIndex = -1;
  let testpaColIndex = -1;

  const maxRow = Math.min(50, ws.actualRowCount || ws.rowCount);
  const maxCol = Math.max(40, ws.actualColumnCount || ws.columnCount);

  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getCell(r, c);
      if (cell.isMerged && cell.master !== cell) continue;
      const text = cellText(r, c);
      if (!text) continue;
      const lower = text.toLowerCase().trim();
      if (lower.includes('power test strand')) powerStrandColIndex = c - 1;
      if (lower === 'waldo id' || lower.includes('waldo')) waldoColIndex = c - 1;
      if (lower === 'terminal') terminalColIndex = c - 1;
      if (lower === 'cable id' || (lower.includes('cable') && lower.includes('id'))) cableIdColIndex = c - 1;
      if (lower.includes('otdr') && lower.includes('strand')) otdrColIndex = c - 1;
      if (lower.includes('total') && lower.includes('strand')) totalStrandsColIndex = c - 1;
      if (lower.includes('testp') && !lower.includes('testpa') && lower.includes('qty')) testpColIndex = c - 1;
      if (lower.includes('testpa') && lower.includes('qty')) testpaColIndex = c - 1;
    }
    if (powerStrandColIndex !== -1 && totalStrandsColIndex !== -1) {
      headerRowIndex = r - 1;
      break;
    }
  }

  if (headerRowIndex === -1 || powerStrandColIndex === -1 || totalStrandsColIndex === -1) {
    throw new Error('Could not locate Power Test Strand / Total Strands header columns.');
  }

  const terminals: Terminal[] = [];
  const lastRow = Math.max(ws.rowCount, ws.actualRowCount, 0);
  for (let r = headerRowIndex + 2; r <= lastRow; r++) {
    const powerCell = ws.getCell(r, powerStrandColIndex + 1);
    if (powerCell.isMerged && powerCell.master !== powerCell) continue;
    const power = cellNumber(r, powerStrandColIndex + 1);
    const total = cellNumber(r, totalStrandsColIndex + 1);
    if (power === null || total === null || power <= 0 || total <= 0) continue;
    terminals.push({
      rowIndex: r - 1,
      waldoId: waldoColIndex !== -1 ? cellText(r, waldoColIndex + 1) : '',
      terminalName: terminalColIndex !== -1 ? cellText(r, terminalColIndex + 1) : '',
      cableId: cableIdColIndex !== -1 ? cellText(r, cableIdColIndex + 1) : '',
      powerTestStrand: power,
      otdrTestStrand: otdrColIndex !== -1 ? cellText(r, otdrColIndex + 1) : '',
      totalStrands: total,
      testpQty: testpColIndex !== -1 ? (cellNumber(r, testpColIndex + 1) ?? cellText(r, testpColIndex + 1)) : '',
      testpaQty: testpaColIndex !== -1 ? (cellNumber(r, testpaColIndex + 1) ?? cellText(r, testpaColIndex + 1)) : '',
    });
  }

  return {
    fileBuffer: arrayBuffer,
    sheetName: ws.name,
    headerRowIndex,
    powerStrandColIndex,
    totalStrandsColIndex,
    terminals,
  };
}

export function computeStaggeredColumns(terminals: Terminal[]): Terminal[] {
  let port = 1;
  return terminals.map((t, idx) => {
    const startPort = t.totalStrands === 2 ? 2 : 1;
    const offset = Math.max(0, port - startPort);
    const staggeredStrand = t.powerTestStrand + offset;
    const result = { ...t, staggeredPort: port, staggeredStrand };

    const next = terminals[idx + 1];
    const currentIsTwo = t.totalStrands === 2;
    const nextIsTwo = next && next.totalStrands === 2;
    if (currentIsTwo && nextIsTwo && port === 2) {
      port = 1;
    } else {
      port = (port % 4) + 1;
    }
    return result;
  });
}

export async function generateConvertedXlsx(
  parsed: ParsedWorkbook,
  terminals: Terminal[],
  portColIndex = 38,
  strandColIndex = 39
): Promise<Uint8Array> {
  const portColNum = portColIndex + 1;
  const strandColNum = strandColIndex + 1;
  const portColLetter = colNumberToLetters(portColNum);
  const strandColLetter = colNumberToLetters(strandColNum);
  const headerRowNum = parsed.headerRowIndex + 1;
  const powerColLetter = colNumberToLetters(parsed.powerStrandColIndex + 1);

  const zip = unzipSync(new Uint8Array(parsed.fileBuffer));
  const sheetPath = Object.keys(zip).find((p) => /^xl\/worksheets\/sheet1\.xml$/i.test(p));
  if (!sheetPath) throw new Error('Could not find sheet1.xml in workbook.');
  let xml = strFromU8(zip[sheetPath]);

  const findCellStyle = (cellRef: string): string => {
    const m = xml.match(new RegExp(`<c[^>]*\\sr="${cellRef}"[^>]*>`));
    if (!m) return '';
    const s = m[0].match(/\ss="(\d+)"/);
    return s ? ` s="${s[1]}"` : '';
  };

  const headerStyleAttr = findCellStyle(`${powerColLetter}${headerRowNum}`);
  const firstTerminal = terminals.find((t) => t.staggeredPort !== undefined);
  const dataStyleAttr = firstTerminal
    ? findCellStyle(`${powerColLetter}${firstTerminal.rowIndex + 1}`)
    : '';

  const replaceColDef = (colNum: number, width: number) => {
    const colLetter = colNumberToLetters(colNum);
    const newDef = `<col min="${colNum}" max="${colNum}" customWidth="1" width="${width}"></col>`;
    const re = new RegExp(`<col\\s[^>]*\\bmin="${colNum}"[^>]*\\bmax="${colNum}"[^>]*(?:\\/>|>\\s*<\\/col>)`);
    if (re.test(xml)) {
      xml = xml.replace(re, newDef);
    } else {
      xml = xml.replace(/<\/cols>/, `${newDef}</cols>`);
      if (!/<\/cols>/.test(xml)) {
        xml = xml.replace(/<sheetData>/, `<cols>${newDef}</cols><sheetData>`);
      }
    }
    void colLetter;
  };
  replaceColDef(portColNum, 14);
  replaceColDef(strandColNum, 16);

  const insertCellsIntoRow = (rowNum: number, cellsXml: string) => {
    const re = new RegExp(`(<row\\s[^>]*\\br="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`);
    if (re.test(xml)) {
      xml = xml.replace(re, `$1$2${cellsXml}$3`);
    } else {
      const sheetDataClose = '</sheetData>';
      const newRow = `<row r="${rowNum}">${cellsXml}</row>`;
      xml = xml.replace(sheetDataClose, `${newRow}${sheetDataClose}`);
    }
  };

  const headerCells =
    `<c${headerStyleAttr} t="inlineStr" r="${portColLetter}${headerRowNum}"><is><t xml:space="preserve">${escapeXml('Staggered Port')}</t></is></c>` +
    `<c${headerStyleAttr} t="inlineStr" r="${strandColLetter}${headerRowNum}"><is><t xml:space="preserve">${escapeXml('Staggered Strand')}</t></is></c>`;
  insertCellsIntoRow(headerRowNum, headerCells);

  for (const t of terminals) {
    if (t.staggeredPort === undefined || t.staggeredStrand === undefined) continue;
    const rowNum = t.rowIndex + 1;
    const cells =
      `<c${dataStyleAttr} r="${portColLetter}${rowNum}"><v>${t.staggeredPort}</v></c>` +
      `<c${dataStyleAttr} r="${strandColLetter}${rowNum}"><v>${t.staggeredStrand}</v></c>`;
    insertCellsIntoRow(rowNum, cells);
  }

  zip[sheetPath] = strToU8(xml);
  const out = zipSync(zip, { level: 6 });
  return out;
}
