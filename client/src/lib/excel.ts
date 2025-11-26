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
    wireCenterClli: "LKGNWI01",
    cfas: "A02VJPG",
    aLoc: "LKGNWI01",
    zLoc: "LKGNWI01",
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

export async function parseExcelFile(file: File): Promise<{ cableId: string | null; strands: number[] }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

        console.log("Parsed Excel Data (first 5 rows):", jsonData.slice(0, 5));

        if (jsonData.length === 0) {
          resolve({ cableId: null, strands: [] });
          return;
        }

        let cableIdIndex = -1;
        let strandIndex = -1;
        let headerRowIndex = -1;

        // Enhanced Header Detection
        for (let i = 0; i < Math.min(20, jsonData.length); i++) {
          const row = jsonData[i];
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
        
        // Fallback: If no strands found via headers, look for numeric columns
        if (strands.length === 0) {
          console.log("No strands found via headers, attempting fallback...");
          
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
        
        console.log(`Extraction complete. CableID: ${foundCableId}, Strands: ${uniqueStrands.length}`);
        resolve({ cableId: foundCableId, strands: uniqueStrands });

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
  baseTemplate: JsonTemplate = DEFAULT_TEMPLATE
): JsonTemplate {
  const newReport = JSON.parse(JSON.stringify(baseTemplate)); // Deep copy
  newReport.Label.cableId = cableId;
  newReport.Label.dateTested = new Date().toISOString().split('T')[0];

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
