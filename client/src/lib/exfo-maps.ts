import { compressToEncodedURIComponent } from 'lz-string';
import { manhattanFeet, parseAddress, terminalStrandNumbers, terminalStrandRange, ExfoTerminal, resolveCityFromCLLI } from './exfo';

declare global {
  interface Window {
    google?: any;
    gm_authFailure?: () => void;
  }
}

export interface GeocodeHit {
  lat: number;
  lng: number;
  locality: string | null;
  stateCode: string | null;
  country: string | null;
  partial?: boolean;
}

const GEOCODE_CACHE_KEY = 'f2job.geocodeCache.v3';

export function loadGeocodeCache(): Map<string, GeocodeHit> {
  try {
    const raw = localStorage.getItem(GEOCODE_CACHE_KEY);
    if (raw) return new Map(Object.entries(JSON.parse(raw)));
  } catch (_) {}
  return new Map();
}

export function saveGeocodeCache(cache: Map<string, GeocodeHit>) {
  try {
    const obj = Object.fromEntries(cache);
    localStorage.setItem(GEOCODE_CACHE_KEY, JSON.stringify(obj));
  } catch (_) {}
}

let mapsLoadPromise: Promise<void> | null = null;
export async function loadGoogleMapsScript(apiKey: string): Promise<void> {
  if (window.google && window.google.maps) return;
  if (mapsLoadPromise) return mapsLoadPromise;
  mapsLoadPromise = new Promise((resolve, reject) => {
    const cb = '__gmapsLoaded_' + Math.random().toString(36).slice(2);
    (window as any)[cb] = () => {
      delete (window as any)[cb];
      resolve();
    };
    const s = document.createElement('script');
    s.src = `https://maps.googleapis.com/maps/api/js?key=${encodeURIComponent(apiKey)}&callback=${cb}`;
    s.async = true;
    s.defer = true;
    s.onerror = () => reject(new Error('Failed to load Google Maps JavaScript API (check your API key and referrer restrictions).'));
    document.head.appendChild(s);
  });
  return mapsLoadPromise;
}

export function geocodeOne(
  geocoder: any,
  query: string,
  cache: Map<string, GeocodeHit>
): Promise<GeocodeHit | null> {
  const cached = cache.get(query);
  if (cached) return Promise.resolve(cached);
  const attempt = (retriesLeft: number): Promise<GeocodeHit | null> => new Promise((resolve) => {
    geocoder.geocode({ address: query, componentRestrictions: { country: 'US' } }, (results: any, status: any) => {
      if (status === 'OK' && results[0]) {
        const r = results[0];
        const loc = r.geometry.location;
        let locality = null, stateCode = null, country = null;
        for (const c of r.address_components || []) {
          if (c.types.includes('locality')) locality = c.short_name;
          if (c.types.includes('administrative_area_level_1')) stateCode = c.short_name;
          if (c.types.includes('country')) country = c.short_name;
        }
        const hit: GeocodeHit = { lat: loc.lat(), lng: loc.lng(), locality, stateCode, country };
        cache.set(query, hit);
        resolve(hit);
      } else if (status === 'OVER_QUERY_LIMIT' && retriesLeft > 0) {
        setTimeout(() => attempt(retriesLeft - 1).then(resolve), 1000);
      } else {
        resolve(null);
      }
    });
  });
  return attempt(3);
}

export async function pool<T, R>(
  items: T[],
  fn: (item: T, index: number) => Promise<R>,
  concurrency = 10,
  onProgress?: (done: number, total: number) => void
): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let cursor = 0, done = 0;
  async function worker() {
    while (cursor < items.length) {
      const i = cursor++;
      results[i] = await fn(items[i], i);
      done++;
      if (onProgress) onProgress(done, items.length);
    }
  }
  await Promise.all(Array(Math.min(concurrency, items.length || 1)).fill(0).map(worker));
  return results;
}

const PON_BUCKET_COLORS = [
  '#2962ff', '#ff6f00', '#2e7d32', '#6d4c41', '#607d8b',
  '#ffffff', '#d32f2f', '#212121', '#fbc02d', '#8e24aa',
];
const PON_LIGHT_BUCKETS = new Set([5, 8]);

export function ponBucketWeights(label: string): number[] {
  const weights = new Array(PON_BUCKET_COLORS.length).fill(0);
  if (!label) return weights;
  for (const seg of String(label).split(',')) {
    const m = seg.trim().match(/^(\d+)(?:-(\d+))?$/);
    if (!m) continue;
    const a = parseInt(m[1], 10);
    const b = m[2] != null ? parseInt(m[2], 10) : a;
    if (Number.isNaN(a) || Number.isNaN(b)) continue;
    const lo = Math.min(a, b), hi = Math.max(a, b);
    for (let n = lo; n <= hi; n++) {
      let idx = Math.floor(n / 100);
      if (idx < 0) idx = 0;
      else if (idx >= weights.length) idx = weights.length - 1;
      weights[idx] += 1;
    }
  }
  return weights;
}

export function buildPfpIconUrl(): string {
  const label = 'PFP';
  const w = 40, h = 28, pad = 2, rx = 7;
  const rectW = w - pad * 2, rectH = h - pad * 2;
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">` +
    `<defs><filter id="pfpShadow" x="-30%" y="-30%" width="160%" height="160%">` +
    `<feDropShadow dx="0" dy="1.5" stdDeviation="2" flood-color="#000" flood-opacity="0.55"/>` +
    `</filter></defs>` +
    `<g filter="url(#pfpShadow)">` +
    `<rect x="0" y="0" width="${w}" height="${h}" rx="${rx + 1}" fill="#ffffff"/>` +
    `<rect x="${pad}" y="${pad}" width="${rectW}" height="${rectH}" rx="${rx}" fill="#d32f2f" stroke="#7a0000" stroke-width="1"/>` +
    `</g>` +
    `<text x="${w / 2}" y="${h - 8}" text-anchor="middle" ` +
    `font-family="Arial,sans-serif" font-size="15" font-weight="900" fill="white" ` +
    `stroke="#7a0000" stroke-width="0.4" paint-order="stroke fill">${label}</text>` +
    `</svg>`;
  return 'data:image/svg+xml;charset=UTF-8,' + encodeURIComponent(svg);
}

const TERMINAL_ICON_CACHE = new Map<string, string>();
export function buildTerminalIconUrl(label: string, port: number | null = null): string {
  const cacheKey = `${label}|${port ?? ''}`;
  if (TERMINAL_ICON_CACHE.has(cacheKey)) return TERMINAL_ICON_CACHE.get(cacheKey)!;
  const display = port != null ? `P${port}·${label}` : label;
  const w = Math.max(28, 10 + display.length * 7);
  const h = 22;
  const innerX = 0.75, innerY = 0.75, innerW = w - 1.5, innerH = h - 1.5, rx = 6;
  // Color buckets are driven by PON numbers only; the port prefix is text-only.
  const weights = ponBucketWeights(label);
  const total = weights.reduce((s, n) => s + n, 0);
  let body: string;
  if (total === 0) {
    body = `<rect x="${innerX}" y="${innerY}" width="${innerW}" height="${innerH}" rx="${rx}" fill="${PON_BUCKET_COLORS[0]}"/>`;
  } else {
    const parts: string[] = [];
    let x = innerX;
    for (let i = 0; i < weights.length; i++) {
      if (!weights[i]) continue;
      const segW = (weights[i] / total) * innerW;
      parts.push(`<rect x="${x}" y="${innerY}" width="${segW}" height="${innerH}" fill="${PON_BUCKET_COLORS[i]}"/>`);
      x += segW;
    }
    body = `<defs><clipPath id="ponClip"><rect x="${innerX}" y="${innerY}" width="${innerW}" height="${innerH}" rx="${rx}"/></clipPath></defs>` +
      `<g clip-path="url(#ponClip)">${parts.join('')}</g>`;
  }
  body += `<rect x="${innerX}" y="${innerY}" width="${innerW}" height="${innerH}" rx="${rx}" fill="none" stroke="white" stroke-width="1.5"/>`;
  let domIdx = 0, domVal = -1;
  for (let i = 0; i < weights.length; i++) if (weights[i] > domVal) { domVal = weights[i]; domIdx = i; }
  const textFill = PON_LIGHT_BUCKETS.has(domIdx) ? '#000' : '#fff';
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">` +
    body +
    `<text x="${w / 2}" y="${h - 6}" text-anchor="middle" font-family="Arial,sans-serif" font-size="12" font-weight="bold" fill="${textFill}">${display}</text>` +
    `</svg>`;
  const url = 'data:image/svg+xml;charset=UTF-8,' + encodeURIComponent(svg);
  TERMINAL_ICON_CACHE.set(cacheKey, url);
  return url;
}

export interface RenderMapOptions {
  apiKey: string;
  city: string;
  pfpName: string | null;
  terminals: ExfoTerminal[];
  cache: Map<string, GeocodeHit>;
  portByRow?: Map<number, number>;
  onStatus?: (msg: string, kind?: 'err' | 'ok' | '') => void;
}

export interface RenderMapResult {
  pfpLocation: GeocodeHit | null;
  distances: Map<number, number>;
  resolved: number;
  failed: number;
  elapsedSec: number;
}

export async function renderEmbeddedMap(
  canvas: HTMLDivElement,
  state: { gmap: any | null; gmarkers: any[]; connections: any[]; showConnections: boolean },
  opts: RenderMapOptions
): Promise<RenderMapResult> {
  const { apiKey, city, pfpName, terminals, cache, portByRow, onStatus } = opts;
  const buildQuery = (addr: string) => city ? `${addr}, ${city}` : addr;

  onStatus?.('Loading Google Maps…');
  await loadGoogleMapsScript(apiKey);

  if (!state.gmap) {
    state.gmap = new window.google.maps.Map(canvas, { zoom: 12, center: { lat: 44.2619, lng: -88.4154 } });
  }
  canvas.style.display = 'block';

  for (const m of state.gmarkers) m.setMap(null);
  state.gmarkers = [];
  for (const l of state.connections) l.setMap(null);
  state.connections = [];

  const geocoder = new window.google.maps.Geocoder();
  const bounds = new window.google.maps.LatLngBounds();
  const distances = new Map<number, number>();

  interface Target { kind: 'pfp' | 'term'; name?: string; t?: ExfoTerminal; addr: string; q: string; }
  const targets: Target[] = [];
  if (pfpName) {
    const a = parseAddress(pfpName);
    targets.push({ kind: 'pfp', name: pfpName, addr: a, q: buildQuery(a) });
  }
  for (const t of terminals) {
    const a = parseAddress(t.terminal);
    targets.push({ kind: 'term', t, addr: a, q: buildQuery(a) });
  }

  const t0 = performance.now();
  onStatus?.(`Geocoding ${targets.length} addresses…`);
  const locations = await pool(targets, (tgt) => geocodeOne(geocoder, tgt.q, cache), 10, (done, total) => {
    if (done % 5 === 0 || done === total) onStatus?.(`Geocoded ${done} / ${total}…`);
  });
  saveGeocodeCache(cache);

  onStatus?.('Checking for geocode outliers…');
  const fixed = await reconcileOutlierLocations(geocoder, targets as any, locations, cache);
  if (fixed) onStatus?.(`Nudged ${fixed} outlier pin${fixed === 1 ? '' : 's'} back to the cluster.`);
  saveGeocodeCache(cache);

  let pfpLocation: GeocodeHit | null = null;
  let resolved = 0, failed = 0;
  for (let i = 0; i < targets.length; i++) {
    const tgt = targets[i];
    const place = locations[i];
    if (!place) {
      failed++;
      continue;
    }
    resolved++;
    if (tgt.kind === 'pfp') {
      pfpLocation = place;
      const pos = new window.google.maps.LatLng(place.lat, place.lng);
      bounds.extend(pos);
      const marker = new window.google.maps.Marker({
        map: state.gmap, position: pos,
        title: `${tgt.name}\n${tgt.addr}`,
        icon: {
          url: buildPfpIconUrl(),
          anchor: new window.google.maps.Point(20, 14),
          scaledSize: new window.google.maps.Size(40, 28),
        },
        zIndex: 999999,
      });
      state.gmarkers.push(marker);
    } else if (tgt.t) {
      const pos = new window.google.maps.LatLng(place.lat, place.lng);
      bounds.extend(pos);
      if (pfpLocation) {
        distances.set(tgt.t.row, manhattanFeet(pfpLocation, place));
      }
      const label = (() => {
        const nums = (tgt.t.powerStrand && tgt.t.powerStrand > 0 ? [tgt.t.powerStrand] : []).concat(tgt.t.otdrStrands).sort((a, b) => a - b);
        if (!nums.length) return '';
        const segs: string[] = [];
        let lo = nums[0], hi = nums[0];
        for (let i = 1; i < nums.length; i++) {
          if (nums[i] === hi + 1) { hi = nums[i]; continue; }
          segs.push(lo === hi ? `${lo}` : `${lo}-${hi}`);
          lo = hi = nums[i];
        }
        segs.push(lo === hi ? `${lo}` : `${lo}-${hi}`);
        return segs.join(',');
      })();
      const port = portByRow?.get(tgt.t.row) ?? null;
      const display = port != null ? `P${port}·${label}` : label;
      const w = Math.max(28, 10 + display.length * 7);
      const marker = new window.google.maps.Marker({
        map: state.gmap, position: pos, title: tgt.t.terminal,
        icon: {
          url: buildTerminalIconUrl(label, port),
          anchor: new window.google.maps.Point(w / 2, 11),
          scaledSize: new window.google.maps.Size(w, 22),
        },
      });
      state.gmarkers.push(marker);
    }
  }

  if (state.gmarkers.length) state.gmap.fitBounds(bounds, 40);
  const elapsedSec = (performance.now() - t0) / 1000;
  return { pfpLocation, distances, resolved, failed, elapsedSec };
}

export function geocodeAllBiased(geocoder: any, query: string, bounds: any): Promise<GeocodeHit[]> {
  return new Promise((resolve) => {
    geocoder.geocode({
      address: query,
      bounds,
      componentRestrictions: { country: 'US' },
    }, (results: any, status: any) => {
      if (status !== 'OK' || !results) { resolve([]); return; }
      const parsed = results.map((r: any) => {
        const loc = r.geometry.location;
        let locality = null, stateCode = null, country = null;
        for (const c of r.address_components || []) {
          if (c.types.includes('locality')) locality = c.short_name;
          if (c.types.includes('administrative_area_level_1')) stateCode = c.short_name;
          if (c.types.includes('country')) country = c.short_name;
        }
        return { lat: loc.lat(), lng: loc.lng(), locality, stateCode, country, partial: !!r.partial_match };
      });
      resolve(parsed);
    });
  });
}

interface OutlierTarget { kind: 'pfp' | 'term'; addr: string; q: string; }

export async function reconcileOutlierLocations(
  geocoder: any,
  targets: OutlierTarget[],
  locations: (GeocodeHit | null)[],
  cache: Map<string, GeocodeHit>
): Promise<number> {
  const pfpIdx = targets.findIndex(t => t.kind === 'pfp');
  const pfpLoc = pfpIdx >= 0 ? locations[pfpIdx] : null;
  if (!pfpLoc) return 0;

  const termEntries: Array<{ i: number; dist?: number }> = [];
  for (let i = 0; i < targets.length; i++) {
    if (targets[i].kind !== 'term' || !locations[i]) continue;
    termEntries.push({ i });
  }
  if (termEntries.length < 3) return 0;

  const allLats = [pfpLoc.lat, ...termEntries.map(e => locations[e.i]!.lat)].slice().sort((a, b) => a - b);
  const allLngs = [pfpLoc.lng, ...termEntries.map(e => locations[e.i]!.lng)].slice().sort((a, b) => a - b);
  const centroid = {
    lat: allLats[Math.floor(allLats.length / 2)],
    lng: allLngs[Math.floor(allLngs.length / 2)],
  };

  for (const e of termEntries) e.dist = manhattanFeet(centroid, locations[e.i]!);
  const sortedDist = termEntries.map(e => e.dist!).slice().sort((a, b) => a - b);
  const pct = (p: number) => sortedDist[Math.min(sortedDist.length - 1, Math.floor(sortedDist.length * p))];
  const p75 = pct(0.75);
  const p90 = pct(0.9);
  const threshold = Math.max(5000, p90 * 2.5, p75 * 3);

  const outliers = termEntries.filter(e => e.dist! > threshold);
  if (!outliers.length) return 0;

  const inliers = termEntries.filter(e => e.dist! <= threshold);
  const bounds = new window.google.maps.LatLngBounds();
  bounds.extend(new window.google.maps.LatLng(pfpLoc.lat, pfpLoc.lng));
  for (const e of inliers) {
    const p = locations[e.i]!;
    bounds.extend(new window.google.maps.LatLng(p.lat, p.lng));
  }

  const cityTally = new Map<string, number>();
  if (pfpLoc.locality && pfpLoc.stateCode) {
    cityTally.set(`${pfpLoc.locality}, ${pfpLoc.stateCode}`, 1);
  }
  for (const e of inliers) {
    const p = locations[e.i]!;
    if (p.locality && p.stateCode) {
      const k = `${p.locality}, ${p.stateCode}`;
      cityTally.set(k, (cityTally.get(k) || 0) + 1);
    }
  }
  const topCities = Array.from(cityTally.entries()).sort((a, b) => b[1] - a[1]).map(x => x[0]);

  let fixed = 0;
  for (const o of outliers) {
    const tgt = targets[o.i];
    const queries = [tgt.q];
    for (const c of topCities) {
      const q = `${tgt.addr}, ${c}`;
      if (!queries.includes(q)) queries.push(q);
    }
    let best: GeocodeHit | null = null;
    let bestScore = o.dist!;
    for (const q of queries) {
      const candidates = await geocodeAllBiased(geocoder, q, bounds);
      for (const cand of candidates) {
        const d = manhattanFeet(centroid, cand);
        if (d + 500 < bestScore || (d < bestScore && best && best.partial && !cand.partial)) {
          best = cand;
          bestScore = d;
        }
      }
    }
    if (best) {
      locations[o.i] = best;
      cache.set(tgt.q, best);
      fixed++;
    }
  }
  return fixed;
}

export async function detectCityCandidates(
  apiKey: string,
  pfpName: string | null,
  terminals: ExfoTerminal[]
): Promise<Array<{ city: string; count: number; total: number }>> {
  if (!apiKey) return [];
  if (!pfpName && !terminals.length) return [];
  await loadGoogleMapsScript(apiKey);
  const geocoder = new window.google.maps.Geocoder();

  const samples: string[] = [];
  const streetsSeen = new Set<string>();
  const streetOf = (a: string) => String(a || '').replace(/^\d+\s+/, '').toUpperCase();
  if (pfpName) {
    const a = parseAddress(pfpName);
    samples.push(a);
    streetsSeen.add(streetOf(a));
  }
  for (const t of terminals) {
    const a = parseAddress(t.terminal);
    const key = streetOf(a);
    if (streetsSeen.has(key)) continue;
    samples.push(a);
    streetsSeen.add(key);
    if (samples.length >= 6) break;
  }
  if (!samples.length) return [];

  const results = await pool(samples, (addr) => new Promise<any>((resolve) => {
    geocoder.geocode({ address: addr, componentRestrictions: { country: 'US' } }, (res: any, status: any) => {
      resolve(status === 'OK' && res && res[0] ? res[0] : null);
    });
  }), 4);

  const tally = new Map<string, number>();
  for (const r of results) {
    if (!r) continue;
    let city: string | null = null, st: string | null = null, country: string | null = null;
    for (const c of r.address_components || []) {
      if (c.types.includes('locality')) city = c.short_name;
      if (c.types.includes('administrative_area_level_1')) st = c.short_name;
      if (c.types.includes('country')) country = c.short_name;
    }
    if (country !== 'US') continue;
    if (city && st) {
      const key = `${city}, ${st}`;
      tally.set(key, (tally.get(key) || 0) + 1);
    }
  }
  const total = samples.length;
  return Array.from(tally.entries())
    .map(([city, count]) => ({ city, count, total }))
    .sort((a, b) => b.count - a.count);
}

export async function dryRunGeocode(
  apiKey: string,
  city: string,
  pfpName: string | null,
  terminals: ExfoTerminal[],
  cache: Map<string, GeocodeHit>
): Promise<{ resolved: number; total: number }> {
  if (!apiKey) return { resolved: 0, total: 0 };
  await loadGoogleMapsScript(apiKey);
  const geocoder = new window.google.maps.Geocoder();
  const buildQuery = (addr: string) => city ? `${addr}, ${city}` : addr;
  const pfpQuery = pfpName ? buildQuery(parseAddress(pfpName)) : null;
  const termQueries = terminals.map(t => buildQuery(parseAddress(t.terminal)));
  if (!termQueries.length) return { resolved: 0, total: 0 };
  const allQueries = pfpQuery ? [pfpQuery, ...termQueries] : termQueries;
  const allResults = await pool(allQueries, (q) => geocodeOne(geocoder, q, cache), 10);
  saveGeocodeCache(cache);
  const pfpLoc = pfpQuery ? allResults[0] : null;
  const termResults = pfpQuery ? allResults.slice(1) : allResults;
  if (!pfpLoc) return { resolved: 0, total: termResults.length };
  let resolved = 0;
  for (const r of termResults) {
    if (!r) continue;
    const dist = manhattanFeet(pfpLoc, r);
    if (dist > 0) resolved++;
  }
  return { resolved, total: termResults.length };
}

const MIN_GEOCODE_COVERAGE = 0.80;

export async function autoResolveCity(
  apiKey: string,
  cllis: string[],
  pfpName: string | null,
  terminals: ExfoTerminal[],
  cache: Map<string, GeocodeHit>,
  manualCity: string,
  onStatus?: (msg: string, kind?: 'err' | 'ok' | '') => void
): Promise<string> {
  if (!apiKey) return manualCity;
  const candidates: Array<{ source: string; city: string }> = [];
  for (const clli of cllis) {
    const hit = resolveCityFromCLLI(clli);
    if (hit && !candidates.some(c => c.city === hit.city)) {
      candidates.push({ source: `CLLI ${hit.code}`, city: hit.city });
    }
  }
  const voteList = await detectCityCandidates(apiKey, pfpName, terminals);
  for (const v of voteList) {
    if (!candidates.some(c => c.city === v.city)) {
      candidates.push({ source: `vote ${v.count}/${v.total}`, city: v.city });
    }
  }
  if (!candidates.length && manualCity) {
    candidates.push({ source: 'manual', city: manualCity });
  }

  for (const cand of candidates) {
    onStatus?.(`Validating "${cand.city}" (${cand.source})…`);
    const { resolved, total } = await dryRunGeocode(apiKey, cand.city, pfpName, terminals, cache);
    const pct = total ? Math.round((resolved / total) * 100) : 0;
    if (total && resolved / total >= MIN_GEOCODE_COVERAGE) {
      onStatus?.(`City: ${cand.city} (${cand.source}, ${resolved}/${total} = ${pct}%)`, 'ok');
      return cand.city;
    }
    onStatus?.(`"${cand.city}" geocoded ${resolved}/${total} (${pct}%) — below 80%, trying next…`, 'err');
  }
  if (candidates.length) {
    onStatus?.(`No candidate reached 80% coverage — falling back to "${candidates[0].city}".`, 'err');
    return candidates[0].city;
  }
  return manualCity;
}

export function drawConnections(
  state: { gmap: any; gmarkers: any[]; connections: any[] },
  city: string,
  terminals: ExfoTerminal[],
  cache: Map<string, GeocodeHit>
) {
  for (const l of state.connections) l.setMap(null);
  state.connections = [];
  if (!state.gmap) return;
  const buildQuery = (addr: string) => city ? `${addr}, ${city}` : addr;
  const ordered = terminals
    .map(t => ({ t, nums: terminalStrandNumbers(t) }))
    .filter(x => x.nums.length > 0)
    .sort((a, b) => a.nums[0] - b.nums[0])
    .map(x => x.t);
  let prev: any = null;
  for (const t of ordered) {
    const addr = parseAddress(t.terminal);
    const loc = cache.get(buildQuery(addr));
    if (!loc) continue;
    const pos = new window.google.maps.LatLng(loc.lat, loc.lng);
    if (prev) {
      const line = new window.google.maps.Polyline({
        path: [prev, pos],
        geodesic: true,
        strokeColor: '#ff6f00',
        strokeOpacity: 0.9,
        strokeWeight: 3,
        icons: [{
          icon: {
            path: window.google.maps.SymbolPath.FORWARD_CLOSED_ARROW,
            scale: 4,
            strokeColor: '#ff6f00',
            fillColor: '#ff6f00',
            fillOpacity: 1,
          },
          offset: '92%',
        }],
        map: state.gmap,
      });
      state.connections.push(line);
    }
    prev = pos;
  }
}

export function clearConnections(state: { connections: any[] }) {
  for (const l of state.connections) l.setMap(null);
  state.connections = [];
}

export function flashPonOnMap(
  state: { gmap: any; gmarkers: any[] },
  ponNum: number,
  terminals: ExfoTerminal[]
): { matches: number; ok: boolean } {
  if (!state.gmap || !state.gmarkers.length) return { matches: 0, ok: false };
  const matchTerms = terminals.filter(t => terminalStrandNumbers(t).includes(ponNum));
  const matchTitles = new Set(matchTerms.map(t => t.terminal));
  const matches = state.gmarkers.filter(m => matchTitles.has(m.getTitle()));
  if (!matches.length) return { matches: 0, ok: true };
  const first = matches[0].getPosition();
  if (first) state.gmap.panTo(first);
  for (const m of matches) m.setAnimation(window.google.maps.Animation.BOUNCE);
  setTimeout(() => { for (const m of matches) m.setAnimation(null); }, 3000);
  return { matches: matches.length, ok: true };
}

export function buildOpenInNewTabHtml(
  apiKey: string,
  cfas: string,
  pfpName: string | null,
  pfpLocation: GeocodeHit | null,
  city: string,
  terminals: ExfoTerminal[],
  cache: Map<string, GeocodeHit>,
  distances: Map<number, number>,
  showConnections: boolean,
  portByRow?: Map<number, number>,
): string {
  const buildQuery = (addr: string) => city ? `${addr}, ${city}` : addr;
  const escHtml = (s: any) => String(s ?? '').replace(/[&<>"']/g, (c: string) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] || c));
  const terms = terminals.map(t => {
    const addr = parseAddress(t.terminal);
    const loc = cache.get(buildQuery(addr));
    if (!loc) return null;
    const distFt = pfpLocation ? manhattanFeet(pfpLocation, loc) : null;
    const nums = terminalStrandNumbers(t);
    return {
      name: t.terminal, addr, waldo: t.waldo || '',
      port: portByRow?.get(t.row) ?? null,
      lat: loc.lat, lng: loc.lng,
      strands: terminalStrandRange(t),
      strandNums: nums,
      minStrand: nums.length ? nums[0] : Number.POSITIVE_INFINITY,
      dist: distFt != null ? `${Math.round(distFt).toLocaleString()} ft` : null,
    };
  }).filter(Boolean) as any[];

  const pfp = pfpName && pfpLocation ? {
    name: pfpName,
    addr: parseAddress(pfpName),
    lat: pfpLocation.lat,
    lng: pfpLocation.lng,
  } : null;

  return `<!doctype html>
<html lang="en"><head>
<meta charset="utf-8" />
<title>${escHtml(cfas)} — Terminal map</title>
<style>
  html, body, #map { height: 100%; margin: 0; padding: 0; }
  .legend { position: absolute; top: 10px; left: 10px; z-index: 5; background: rgba(255,255,255,0.95); padding: 8px 12px; font: 12px -apple-system, "Segoe UI", Roboto, sans-serif; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,0.2); color: #222; }
  .legend .sw { display: inline-block; width: 10px; height: 10px; margin-right: 4px; border-radius: 50%; vertical-align: middle; }
  .legend label { display: flex; align-items: center; gap: 6px; margin-top: 8px; cursor: pointer; user-select: none; }
  .legend label input { margin: 0; }
</style>
</head><body>
<div class="legend">
  <div><span class="sw" style="background:#d32f2f"></span> PFP (starting point)</div>
  <div style="margin-top:6px; font-weight:600;">Terminal (Pon count)</div>
  <div style="display:grid; grid-template-columns:auto auto; gap:2px 10px; margin-top:2px;">
    <div><span class="sw" style="background:#2962ff"></span>1-99</div>
    <div><span class="sw" style="background:#ff6f00"></span>100-199</div>
    <div><span class="sw" style="background:#2e7d32"></span>200-299</div>
    <div><span class="sw" style="background:#6d4c41"></span>300-399</div>
    <div><span class="sw" style="background:#607d8b"></span>400-499</div>
    <div><span class="sw" style="background:#ffffff; border:1px solid #bbb;"></span>500-599</div>
    <div><span class="sw" style="background:#d32f2f"></span>600-699</div>
    <div><span class="sw" style="background:#212121"></span>700-799</div>
    <div><span class="sw" style="background:#fbc02d"></span>800-899</div>
    <div><span class="sw" style="background:#8e24aa"></span>900-999</div>
  </div>
  <div style="margin-top:4px; color:#666;">${terms.length} terminals${pfp ? ' — from ' + escHtml(pfp.addr) : ''}</div>
  <label><input type="checkbox" id="connToggle" ${showConnections ? 'checked' : ''}/> Connect pins by strand order</label>
  <div style="margin-top:6px; display:flex; gap:6px; align-items:center;">
    <input type="number" id="ponSearch" placeholder="Find Pon count" min="0" style="flex:1; width:110px; padding:3px 6px; font:12px inherit; border:1px solid #bbb; border-radius:4px;" />
    <button id="ponSearchBtn" style="padding:3px 8px; font:12px inherit; border:1px solid #bbb; border-radius:4px; background:#eee; cursor:pointer;">Flash</button>
  </div>
  <div id="ponSearchStatus" style="margin-top:4px; color:#666; font-size:11px; min-height:14px;"></div>
</div>
<div id="map"></div>
<script>
  const PFP = ${JSON.stringify(pfp)};
  const TERMS = ${JSON.stringify(terms)};
  const TERM_ICONS = {};
  function pfpIcon() {
    const w=40,h=28,pad=2,rx=7,rectW=w-pad*2,rectH=h-pad*2;
    const svg='<svg xmlns="http://www.w3.org/2000/svg" width="'+w+'" height="'+h+'" viewBox="0 0 '+w+' '+h+'"><defs><filter id="pfpShadow" x="-30%" y="-30%" width="160%" height="160%"><feDropShadow dx="0" dy="1.5" stdDeviation="2" flood-color="#000" flood-opacity="0.55"/></filter></defs><g filter="url(#pfpShadow)"><rect x="0" y="0" width="'+w+'" height="'+h+'" rx="'+(rx+1)+'" fill="#ffffff"/><rect x="'+pad+'" y="'+pad+'" width="'+rectW+'" height="'+rectH+'" rx="'+rx+'" fill="#d32f2f" stroke="#7a0000" stroke-width="1"/></g><text x="'+(w/2)+'" y="'+(h-8)+'" text-anchor="middle" font-family="Arial,sans-serif" font-size="15" font-weight="900" fill="white" stroke="#7a0000" stroke-width="0.4" paint-order="stroke fill">PFP</text></svg>';
    return { url:'data:image/svg+xml;charset=UTF-8,'+encodeURIComponent(svg), anchor:new google.maps.Point(w/2,h/2), scaledSize:new google.maps.Size(w,h) };
  }
  var PON_BUCKET_COLORS=['#2962ff','#ff6f00','#2e7d32','#6d4c41','#607d8b','#ffffff','#d32f2f','#212121','#fbc02d','#8e24aa'];
  var PON_LIGHT_BUCKETS={5:1,8:1};
  function ponBucketWeights(label){var w=[0,0,0,0,0,0,0,0,0,0];if(!label)return w;var s=String(label).split(',');for(var si=0;si<s.length;si++){var m=s[si].trim().match(/^(\\d+)(?:-(\\d+))?$/);if(!m)continue;var a=parseInt(m[1],10),b=m[2]!=null?parseInt(m[2],10):a;if(isNaN(a)||isNaN(b))continue;var lo=Math.min(a,b),hi=Math.max(a,b);for(var n=lo;n<=hi;n++){var idx=Math.floor(n/100);if(idx<0)idx=0;else if(idx>=w.length)idx=w.length-1;w[idx]+=1;}}return w;}
  function terminalIcon(label,port){var key=label+'|'+(port==null?'':port);if(TERM_ICONS[key])return TERM_ICONS[key];var display=(port!=null)?('P'+port+'·'+label):label;var w=Math.max(28,10+String(display).length*7),h=22,iX=0.75,iY=0.75,iW=w-1.5,iH=h-1.5,rx=6;var weights=ponBucketWeights(label),total=0;for(var i=0;i<weights.length;i++)total+=weights[i];var body;if(total===0){body='<rect x="'+iX+'" y="'+iY+'" width="'+iW+'" height="'+iH+'" rx="'+rx+'" fill="'+PON_BUCKET_COLORS[0]+'"/>';}else{var parts='',x=iX;for(var j=0;j<weights.length;j++){if(!weights[j])continue;var sw=(weights[j]/total)*iW;parts+='<rect x="'+x+'" y="'+iY+'" width="'+sw+'" height="'+iH+'" fill="'+PON_BUCKET_COLORS[j]+'"/>';x+=sw;}body='<defs><clipPath id="ponClip"><rect x="'+iX+'" y="'+iY+'" width="'+iW+'" height="'+iH+'" rx="'+rx+'"/></clipPath></defs><g clip-path="url(#ponClip)">'+parts+'</g>';}body+='<rect x="'+iX+'" y="'+iY+'" width="'+iW+'" height="'+iH+'" rx="'+rx+'" fill="none" stroke="white" stroke-width="1.5"/>';var dI=0,dV=-1;for(var k=0;k<weights.length;k++)if(weights[k]>dV){dV=weights[k];dI=k;}var tF=PON_LIGHT_BUCKETS[dI]?'#000':'#fff';var svg='<svg xmlns="http://www.w3.org/2000/svg" width="'+w+'" height="'+h+'" viewBox="0 0 '+w+' '+h+'">'+body+'<text x="'+(w/2)+'" y="'+(h-6)+'" text-anchor="middle" font-family="Arial,sans-serif" font-size="12" font-weight="bold" fill="'+tF+'">'+display+'</text></svg>';TERM_ICONS[key]={url:'data:image/svg+xml;charset=UTF-8,'+encodeURIComponent(svg),anchor:new google.maps.Point(w/2,h/2),scaledSize:new google.maps.Size(w,h)};return TERM_ICONS[key];}
  let MAP=null;const LINES=[];function clearLines(){for(const l of LINES)l.setMap(null);LINES.length=0;}
  function drawLines(){clearLines();if(!MAP)return;const ord=TERMS.slice().sort((a,b)=>a.minStrand-b.minStrand);let prev=null;for(const t of ord){const pos=new google.maps.LatLng(t.lat,t.lng);if(prev){LINES.push(new google.maps.Polyline({path:[prev,pos],geodesic:true,strokeColor:'#ff6f00',strokeOpacity:0.9,strokeWeight:3,icons:[{icon:{path:google.maps.SymbolPath.FORWARD_CLOSED_ARROW,scale:4,strokeColor:'#ff6f00',fillColor:'#ff6f00',fillOpacity:1},offset:'92%'}],map:MAP}));}prev=pos;}}
  function initMap(){const center=PFP?{lat:PFP.lat,lng:PFP.lng}:{lat:TERMS[0].lat,lng:TERMS[0].lng};MAP=new google.maps.Map(document.getElementById('map'),{zoom:12,center});const bounds=new google.maps.LatLngBounds();if(PFP){const m=new google.maps.Marker({map:MAP,position:{lat:PFP.lat,lng:PFP.lng},icon:pfpIcon(),title:PFP.name+'\\n'+PFP.addr,zIndex:999999});bounds.extend(m.getPosition());const info=new google.maps.InfoWindow({content:'<b>PFP (starting point)</b><br>'+PFP.name+'<br>'+PFP.addr});m.addListener('click',()=>info.open({anchor:m,map:MAP}));}var TERM_MARKERS=[];for(const t of TERMS){const m=new google.maps.Marker({map:MAP,position:{lat:t.lat,lng:t.lng},icon:terminalIcon(t.strands,t.port),title:t.name});m._strandNums=Array.isArray(t.strandNums)?t.strandNums:[];TERM_MARKERS.push(m);bounds.extend(m.getPosition());const info=new google.maps.InfoWindow({content:'<b>'+t.name+'</b><br>'+t.addr+'<br>Strands: '+t.strands+(t.port!=null?'<br>Test Port: '+t.port:'')+(t.waldo?'<br>Waldo: '+t.waldo:'')+(t.dist?'<br>Distance from PFP: '+t.dist:'')});m.addListener('click',()=>info.open({anchor:m,map:MAP}));}MAP.fitBounds(bounds,40);const tg=document.getElementById('connToggle');tg.addEventListener('change',()=>{if(tg.checked)drawLines();else clearLines();});if(tg.checked)drawLines();var sI=document.getElementById('ponSearch'),sB=document.getElementById('ponSearchBtn'),sS=document.getElementById('ponSearchStatus');function flashPon(){var raw=(sI.value||'').trim();if(!raw)return;var n=parseInt(raw,10);if(isNaN(n)){sS.textContent='Enter a numeric Pon count.';sS.style.color='#b71c1c';return;}var matches=TERM_MARKERS.filter(function(mk){return mk._strandNums&&mk._strandNums.indexOf(n)!==-1;});if(!matches.length){sS.textContent='No terminal contains Pon count '+n+'.';sS.style.color='#b71c1c';return;}MAP.panTo(matches[0].getPosition());matches.forEach(function(mk){mk.setAnimation(google.maps.Animation.BOUNCE);});setTimeout(function(){matches.forEach(function(mk){mk.setAnimation(null);});},3000);sS.textContent='Flashing '+matches.length+' terminal'+(matches.length>1?'s':'')+' with Pon count '+n+'.';sS.style.color='#1b5e20';}sB.addEventListener('click',flashPon);sI.addEventListener('keydown',function(e){if(e.key==='Enter')flashPon();});}
  window.gm_authFailure=function(){document.body.insertAdjacentHTML('afterbegin','<div style="padding:12px;background:#ffebee;color:#b71c1c;font:13px sans-serif;">Google Maps rejected the API key.</div>');};
<\/script>
<script src="https://maps.googleapis.com/maps/api/js?key=${encodeURIComponent(apiKey)}&callback=initMap" async defer><\/script>
</body></html>`;
}

interface ShareTerm {
  name: string; addr: string; waldo: string;
  port: number | null;
  lat: number; lng: number;
  strands: string; strandNums: number[]; minStrand: number;
  dist: string | null;
}

export function buildShareData(
  title: string,
  pfpName: string | null,
  pfpLocation: GeocodeHit | null,
  city: string,
  terminals: ExfoTerminal[],
  cache: Map<string, GeocodeHit>,
  showConnections: boolean,
  portByRow?: Map<number, number>,
): { title: string; pfp: any; terms: ShareTerm[]; connect: boolean } {
  const buildQuery = (addr: string) => city ? `${addr}, ${city}` : addr;
  const terms = terminals.map(t => {
    const addr = parseAddress(t.terminal);
    const loc = cache.get(buildQuery(addr));
    if (!loc) return null;
    const distFt = pfpLocation ? manhattanFeet(pfpLocation, loc) : null;
    const nums = terminalStrandNumbers(t);
    return {
      name: t.terminal, addr, waldo: t.waldo || '',
      port: portByRow?.get(t.row) ?? null,
      lat: loc.lat, lng: loc.lng,
      strands: terminalStrandRange(t),
      strandNums: nums,
      minStrand: nums.length ? nums[0] : Number.POSITIVE_INFINITY,
      dist: distFt != null ? `${Math.round(distFt).toLocaleString()} ft` : null,
    } as ShareTerm;
  }).filter(Boolean) as ShareTerm[];
  const pfp = pfpName && pfpLocation ? {
    name: pfpName, addr: parseAddress(pfpName),
    lat: pfpLocation.lat, lng: pfpLocation.lng,
  } : null;
  return { title, pfp, terms, connect: showConnections };
}

// Origin for share links. In dev this is localhost (recipient can't open, but the flow works);
// in prod it's whatever domain serves the app (e.g. https://www.agenticpropertyos.com).
function viewerOrigin(): string {
  return window.location.origin;
}

function longFormUrl(apiKey: string, data: ReturnType<typeof buildShareData>): string {
  const compressed = compressToEncodedURIComponent(JSON.stringify(data));
  return `${viewerOrigin()}/map.html#k=${encodeURIComponent(apiKey)}&d=${compressed}`;
}

// Tries the /api/share endpoint for a short ID-style URL.
// Falls back to the long-form fragment URL if the server is unavailable.
export async function buildMapViewerUrl(
  apiKey: string,
  data: ReturnType<typeof buildShareData>,
): Promise<string> {
  const compressed = compressToEncodedURIComponent(JSON.stringify(data));
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), 8000);
  try {
    const r = await fetch('/api/share', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ k: apiKey, d: compressed }),
      signal: ctl.signal,
    });
    if (r.ok) {
      const j = await r.json();
      if (j && typeof j.id === 'string' && /^[A-Za-z0-9]{4,32}$/.test(j.id)) {
        return `${viewerOrigin()}/map.html?id=${j.id}`;
      }
    }
  } catch (_) {
    // fall through to long-form
  } finally {
    clearTimeout(timer);
  }
  return longFormUrl(apiKey, data);
}

