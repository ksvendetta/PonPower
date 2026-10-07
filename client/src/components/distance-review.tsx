import { Button } from '@/components/ui/button';
import { formatDistance, type ExfoTerminal } from '@/lib/exfo';
import type { DistanceOutlier, MapPoint } from '@/lib/distance-review';

export function DistanceReview({ outliers, decisions, raw, terminals, points, pfp, onDecision }: {
  outliers: DistanceOutlier[]; decisions: Map<number, 'keep' | 'exclude'>;
  raw: Map<number, number>; terminals: ExfoTerminal[]; points: Map<number, MapPoint>;
  pfp: MapPoint | null;
  onDecision: (row: number, choice: 'keep' | 'exclude') => void;
}) {
  if (!outliers.length) return null;
  return <div className="rounded border border-amber-500/50 p-3 space-y-3">
    <p className="font-semibold text-sm">Verify unusual map distances</p>
    <p className="text-xs text-muted-foreground">Flagged measurements are excluded from displays, exports, shared maps, and test profile calculations until you confirm them. The filter uses six times the average nearest-terminal spacing, with a 1,500 ft minimum.</p>
    {outliers.map(o => {
      const decision = decisions.get(o.row);
      const point = points.get(o.row);
      const terminal = terminals.find(t => t.row === o.row);
      return <div key={o.row} className="border-t pt-2 text-xs space-y-2">
        <p>{terminal?.terminal || `Row ${o.row}`} · {decision === 'exclude' ? 'Measurement excluded' : decision === 'keep' ? 'Measurement confirmed' : 'Needs review'}</p>
        {decision !== 'exclude' && <p>{raw.has(o.row) ? `${formatDistance(raw.get(o.row)!)} from PFP. ` : ''}{o.reason === 'pfp' ? 'PFP is far from every terminal; verify the PFP pin.' : o.reason === 'cluster' ? 'This small group of pins is isolated from the main terminal cluster.' : `Nearest terminal: ${formatDistance(o.nearestFeet)}; average spacing: ${formatDistance(o.averageFeet)}.`}</p>}
        <div className="flex flex-wrap gap-2 items-center">
          {point && <a className="text-primary underline" href={`https://www.google.com/maps/search/?api=1&query=${point.lat},${point.lng}`} target="_blank" rel="noopener noreferrer">Verify pin in Google Maps</a>}
          {o.reason === 'pfp' && pfp && <a className="text-primary underline" href={`https://www.google.com/maps/search/?api=1&query=${pfp.lat},${pfp.lng}`} target="_blank" rel="noopener noreferrer">Verify PFP pin</a>}
          <Button size="sm" variant={decision === 'keep' ? 'default' : 'outline'} onClick={() => onDecision(o.row, 'keep')}>Keep measurement</Button>
          <Button size="sm" variant={decision === 'exclude' ? 'default' : 'outline'} onClick={() => onDecision(o.row, 'exclude')}>Exclude measurement</Button>
        </div>
      </div>;
    })}
  </div>;
}
