import { manhattanFeet, parseAddress, ExfoTerminal } from './exfo';

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
export function buildTerminalIconUrl(label: string): string {
  if (TERMINAL_ICON_CACHE.has(label)) return TERMINAL_ICON_CACHE.get(label)!;
  const w = Math.max(28, 10 + label.length * 7);
  const h = 22;
  const innerX = 0.75, innerY = 0.75, innerW = w - 1.5, innerH = h - 1.5, rx = 6;
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
    `<text x="${w / 2}" y="${h - 6}" text-anchor="middle" font-family="Arial,sans-serif" font-size="12" font-weight="bold" fill="${textFill}">${label}</text>` +
    `</svg>`;
  const url = 'data:image/svg+xml;charset=UTF-8,' + encodeURIComponent(svg);
  TERMINAL_ICON_CACHE.set(label, url);
  return url;
}

export interface RenderMapOptions {
  apiKey: string;
  city: string;
  pfpName: string | null;
  terminals: ExfoTerminal[];
  cache: Map<string, GeocodeHit>;
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
  const { apiKey, city, pfpName, terminals, cache, onStatus } = opts;
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
      const w = Math.max(28, 10 + label.length * 7);
      const marker = new window.google.maps.Marker({
        map: state.gmap, position: pos, title: tgt.t.terminal,
        icon: {
          url: buildTerminalIconUrl(label),
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
