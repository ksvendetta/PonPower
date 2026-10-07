import { useEffect, useMemo, useState } from 'react';
import { findDistanceOutliers, reviewedDistances, type MapPoint } from '@/lib/distance-review';

export function useDistanceReview(raw: Map<number, number>, points: Map<number, MapPoint>, pfp: MapPoint | null, rows: number[]) {
  const [decisions, setDecisions] = useState<Map<number, 'keep' | 'exclude'>>(new Map());
  useEffect(() => { setDecisions(new Map()); }, [raw, points]);
  const activeKey = rows.join(',');
  const active = useMemo(() => new Set(rows), [activeKey]);
  const outliers = useMemo(() => findDistanceOutliers(new Map(Array.from(points).filter(([row]) => active.has(row))), pfp), [points, pfp, active]);
  const distances = useMemo(() => reviewedDistances(raw, outliers, decisions, active), [raw, outliers, decisions, active]);
  const decide = (row: number, choice: 'keep' | 'exclude') => setDecisions(current => new Map(current).set(row, choice));
  return { distances, outliers, decisions, decide, suppressedRows: new Set(outliers.filter(o => decisions.get(o.row) !== 'keep').map(o => o.row)) };
}
