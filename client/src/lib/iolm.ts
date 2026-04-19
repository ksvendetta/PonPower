export interface IolmLinkMeasurement {
  wavelength: number;
  totalLoss: number;
  totalORL?: number;
}

export interface IolmTestedEntry {
  strand: number;
  spanLength?: number;
  passFail: string;
  linkMeasurements: IolmLinkMeasurement[];
  // OTDR event data is stripped during cleanup:
  events?: unknown;
  [key: string]: unknown;
}

export interface IolmReport {
  Label: {
    tester?: string;
    wireCenterClli?: string;
    cfas?: string;
    aLoc?: string;
    zLoc?: string;
    dateTested?: string;
    model?: string;
    serialNumber?: string;
    version?: string;
    calibrationDate?: string;
    spliceLossThreshold?: number;
    connectorLossThreshold?: number;
    eventReflectanceThreshold?: number;
    strandOrlThreshold?: number;
    cableId?: string;
    Tested: IolmTestedEntry[];
    [key: string]: unknown;
  };
}

export type PfpSize = 288 | 432 | 576 | 864;
export const PFP_SIZES: PfpSize[] = [288, 432, 576, 864];

/**
 * Clean an IOLM-style JSON report:
 *   - Force every Tested[].passFail to "pass"
 *   - Strip the per-strand "events" array (link-level data only)
 * Mutates a deep copy of the input; original object is preserved.
 */
export function cleanReport(input: IolmReport): IolmReport {
  const clone: IolmReport = JSON.parse(JSON.stringify(input));
  const tested = clone?.Label?.Tested;
  if (Array.isArray(tested)) {
    for (const entry of tested) {
      entry.passFail = "pass";
      if ("events" in entry) {
        delete entry.events;
      }
    }
  }
  return clone;
}

/**
 * Return the strand numbers in [1..pfpSize] that are NOT present in the
 * Tested array. Sorted ascending.
 */
export function findMissingStrands(
  report: IolmReport | null,
  pfpSize: PfpSize | null
): number[] {
  if (!report || !pfpSize) return [];
  const tested = report.Label?.Tested ?? [];
  const present = new Set<number>();
  for (const entry of tested) {
    if (typeof entry.strand === "number") present.add(entry.strand);
  }
  const missing: number[] = [];
  for (let i = 1; i <= pfpSize; i++) {
    if (!present.has(i)) missing.push(i);
  }
  return missing;
}

/**
 * Stringify in the 1-space-indent style matching the "after" example
 * (Waldo.json).
 */
export function formatReport(report: IolmReport): string {
  return JSON.stringify(report, null, 1);
}
