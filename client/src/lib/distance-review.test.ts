import assert from 'node:assert/strict';
import { test } from 'node:test';
import { findDistanceOutliers, reviewedDistances } from './distance-review';
import { buildShareData, type GeocodeHit } from './exfo-maps';
import { IOLM_SHORT_LINK_FT } from './exfo';

const hit = (lat: number, lng: number): GeocodeHit => ({ lat, lng, locality: null, country: 'US', stateCode: 'WI' });
const points = new Map([[1, hit(44, -88)], [2, hit(44.001, -88)], [3, hit(44.002, -88)], [4, hit(45, -89)]]);
const pfp = hit(44, -88.001);

test('nearest-neighbor average detects isolated pins without letting them inflate the baseline', () => {
  const flags = findDistanceOutliers(points, pfp);
  assert.deepEqual(flags.map(f => f.row), [4]);
  assert.ok(flags[0].nearestFeet > flags[0].thresholdFeet);
  assert.ok(flags[0].averageFeet < 1000);
  assert.deepEqual(findDistanceOutliers(new Map(Array.from(points).slice(0, 3)), pfp), []);
  assert.deepEqual(findDistanceOutliers(new Map([[1, hit(44, -88)], [2, hit(44.001, -88)], [3, hit(44.002, -88)], [4, hit(44.003, -88)]]), pfp), []);
});

test('a distant PFP does not classify normal terminals as outliers', () => {
  const cluster = new Map([[1, hit(44, -88)], [2, hit(44.001, -88)], [3, hit(44.002, -88)], [4, hit(44.003, -88)]]);
  const flags = findDistanceOutliers(cluster, hit(46, -90));
  assert.deepEqual(flags, []);
});

test('a small group of nearby bad pins far from the main cluster is still flagged', () => {
  const cluster = new Map(Array.from({ length: 8 }, (_, i) => [i + 1, hit(44 + i * 0.001, -88)] as const));
  cluster.set(9, hit(45, -89));
  cluster.set(10, hit(45.001, -89));
  assert.deepEqual(findDistanceOutliers(cluster, pfp).map(o => o.row), [9, 10]);
});

test('pending and rejected measurements are omitted from shared maps and IOLM span calculations; confirming restores them', () => {
  const raw = new Map([[1, 1000], [2, 2000], [3, 3000], [4, 500000]]);
  const flags = findDistanceOutliers(points, pfp);
  const active = new Set([1, 2, 3, 4]);
  for (const decisions of [new Map(), new Map([[4, 'exclude' as const]])]) {
    const distances = reviewedDistances(raw, flags, decisions, active);
    assert.equal(distances.has(4), false);
    assert.ok(Math.max(...distances.values()) < IOLM_SHORT_LINK_FT);
    const cache = new Map([['Address', points.get(4)!]]);
    const data = buildShareData('test', 'PFP', pfp, '', [{ row: 4, terminal: 'Address', waldo: '123', cable: 'PON', powerStrand: 1, total: 2, otdrRaw: '', otdrStrands: [1, 2] }], cache, false, undefined, distances);
    assert.equal(data.terms[0].dist, null);
    assert.ok(!JSON.stringify(data).includes('500,000'));
  }
  const kept = reviewedDistances(raw, flags, new Map([[4, 'keep']]), active);
  assert.equal(kept.get(4), 500000);
  const excludedTerminal = reviewedDistances(raw, flags, new Map([[4, 'keep']]), new Set([1, 2, 3]));
  assert.equal(excludedTerminal.has(4), false);
  // Changing duplicate selections can change the baseline; a rejected measurement stays rejected.
  assert.equal(reviewedDistances(raw, [], new Map([[4, 'exclude']]), active).has(4), false);
});
