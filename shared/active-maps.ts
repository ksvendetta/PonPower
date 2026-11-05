import LZString from "lz-string";

export interface ActiveMap {
  id: string;
  project: string;
  createdAt: string;
}

const lz = (LZString as { decompressFromEncodedURIComponent?: (s: string) => string | null }).decompressFromEncodedURIComponent
  ? (LZString as { decompressFromEncodedURIComponent: (s: string) => string | null })
  : ((LZString as { default?: { decompressFromEncodedURIComponent: (s: string) => string | null } }).default ?? LZString);

/** Project number stored as the share payload title. */
export function projectFromSharePayload(compressed: string): string {
  try {
    const raw = lz.decompressFromEncodedURIComponent(compressed);
    if (!raw) return "Terminal map";
    const data = JSON.parse(raw) as { title?: unknown };
    const title = typeof data?.title === "string" ? data.title.trim() : "";
    return title || "Terminal map";
  } catch {
    return "Terminal map";
  }
}

export function toCreatedAtIso(value: unknown): string {
  const date = value instanceof Date ? value : new Date(typeof value === "number" || typeof value === "string" ? value : "");
  if (Number.isNaN(date.getTime())) return new Date(0).toISOString();
  return date.toISOString();
}
