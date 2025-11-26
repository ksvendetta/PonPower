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

        if (jsonData.length === 0) {
          resolve({ cableId: null, strands: [] });
          return;
        }

        // Simple heuristic to find columns
        // We look for "Cable ID" (or similar) and "Strand" (or similar) in the first few rows
        let cableIdIndex = -1;
        let strandIndex = -1;
        let headerRowIndex = -1;

        for (let i = 0; i < Math.min(10, jsonData.length); i++) {
          const row = jsonData[i];
          row.forEach((cell: any, index: number) => {
            if (typeof cell === 'string') {
              const lower = cell.toLowerCase();
              if (lower.includes('cable') && lower.includes('id')) cableIdIndex = index;
              if (lower.includes('strand')) strandIndex = index;
            }
          });
          if (cableIdIndex !== -1 || strandIndex !== -1) {
            headerRowIndex = i;
            break;
          }
        }

        // If we found a header row, extract data
        const strands: number[] = [];
        let foundCableId: string | null = null;

        if (headerRowIndex !== -1) {
          for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row) continue;

            if (cableIdIndex !== -1 && row[cableIdIndex]) {
              // Capture the first non-empty cable ID we find
              if (!foundCableId) foundCableId = String(row[cableIdIndex]);
            }

            if (strandIndex !== -1 && row[strandIndex]) {
              const val = parseInt(row[strandIndex]);
              if (!isNaN(val)) strands.push(val);
            }
          }
        } else {
           // Fallback: try to find any column that looks like strand numbers (integers)
           // and maybe cable id is in the filename or we just ask user
        }

        resolve({ cableId: foundCableId, strands });
      } catch (err) {
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
