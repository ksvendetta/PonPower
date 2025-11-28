import * as XLSX from 'xlsx';

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
        // Look for CFAS in range X2:AM4 (Indices: Rows 1-3, Cols 23-38)
        // The user states the value is ALWAYS in this range (likely merged).
        // In sheet_to_json, the value of a merged cell is in the top-left index.
        // X2 is Row 2 (index 1), Col X (index 23).
        // ---------------------------------------------------------
        let foundCfas: string | null = null;
        console.log("Searching for CFAS in range X2:AM4");
        
        // Direct check at X2 (Row 2, Col 23) first as it's the top-left of the range
        if (jsonData[1] && jsonData[1][23] !== undefined) {
           const val = jsonData[1][23];
           if (typeof val === 'string' && val.trim().length > 0) {
              foundCfas = val.trim();
              console.log(`Found CFAS at exact X2 [1, 23]: ${foundCfas}`);
           } else if (typeof val === 'number') {
              foundCfas = String(val);
              console.log(`Found numeric CFAS at exact X2 [1, 23]: ${foundCfas}`);
           }
        }

        // Fallback: Scan the range if X2 was empty
        if (!foundCfas) {
          // Row loop: 1 to 3 (Rows 2-4)
          for (let r = 1; r <= 3; r++) {
             if (jsonData.length <= r) break;
             const row = jsonData[r];
             if (!row) continue;
             
             // Col loop: 23 to 38 (Cols X to AM)
             for (let c = 23; c <= 38; c++) {
                if (row.length <= c) break;
                const cell = row[c];
                
                // Check for string or number
                if (cell !== null && cell !== undefined && cell !== '') {
                   const strVal = String(cell).trim();
                   if (strVal.length > 0) {
                      foundCfas = strVal;
                      console.log(`Found CFAS candidate in scan at [${r}, ${c}]: ${foundCfas}`);
                      break;
                   }
                }
             }
             if (foundCfas) break;
          }
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

           const uniqueStrands = Array.from(new Set(strands)).sort((a, b) => a - b);
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
        const uniqueStrands = Array.from(new Set(strands)).sort((a, b) => a - b);
        
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
