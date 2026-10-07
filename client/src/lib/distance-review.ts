import { manhattanFeet } from './exfo';

export interface MapPoint { lat: number; lng: number }
export interface DistanceOutlier {
  row: number;
  nearestFeet: number;
  averageFeet: number;
  thresholdFeet: number;
  reason: 'terminal' | 'pfp' | 'cluster';
}

/** Average nearest-neighbor spacing, trimming the largest 20% so bad pins cannot set the baseline. */
export function findDistanceOutliers(points: Map<number, MapPoint>, pfp: MapPoint | null): DistanceOutlier[] {
  const entries = Array.from(points).filter(([, p]) => Number.isFinite(p.lat) && Number.isFinite(p.lng));
  if (entries.length < 4) return [];
  const nearest = entries.map(([row, point]) => ({ row, nearestFeet: Math.min(...entries.filter(([r]) => r !== row).map(([, p]) => manhattanFeet(point, p))) }));
  const sorted = nearest.map(p => p.nearestFeet).sort((a, b) => a - b);
  const baseline = sorted.slice(0, Math.max(1, Math.floor(sorted.length * 0.8)));
  const averageFeet = baseline.reduce((sum, ft) => sum + ft, 0) / baseline.length;
  const thresholdFeet = Math.max(1500, averageFeet * 6);
  // Also catch small remote groups: two wrong pins near each other would have
  // short nearest-neighbor distances despite being far from the main cluster.
  const components: number[][] = [];
  const remaining = new Set(entries.map(([row]) => row));
  while (remaining.size) {
    const queue = [Array.from(remaining)[0]];
    remaining.delete(queue[0]);
    for (let i = 0; i < queue.length; i++) {
      const point = points.get(queue[i])!;
      for (const row of Array.from(remaining)) if (manhattanFeet(point, points.get(row)!) <= thresholdFeet) {
        queue.push(row);
        remaining.delete(row);
      }
    }
    components.push(queue);
  }
  components.sort((a, b) => b.length - a.length);
  const remote = new Set(components[0].length >= entries.length * 0.6
    ? components.slice(1).filter(group => group.length <= Math.max(1, entries.length * 0.2)).flat() : []);
  return nearest.filter(p => p.nearestFeet > thresholdFeet || remote.has(p.row)).map(p => ({ ...p, averageFeet, thresholdFeet, reason: remote.has(p.row) && p.nearestFeet <= thresholdFeet ? 'cluster' : 'terminal' }));
}

export function reviewedDistances(raw: Map<number, number>, outliers: DistanceOutlier[], decisions: Map<number, 'keep' | 'exclude'>, activeRows: Set<number>): Map<number, number> {
  const flagged = new Set(outliers.map(o => o.row));
  return new Map(Array.from(raw).filter(([row, ft]) => activeRows.has(row) && Number.isFinite(ft) && ft >= 0 && decisions.get(row) !== 'exclude' && (!flagged.has(row) || decisions.get(row) === 'keep')));
}
