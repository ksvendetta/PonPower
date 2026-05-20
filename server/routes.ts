import type { Express, Request, Response } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";

// In-memory mirror of the Vercel /api/share function for local dev.
// Vercel serverless functions in /api don't run under the Express dev server,
// so without this mock the client falls back to long-form fragment URLs that
// exceed QR-code capacity.
interface ShareRecord { k: string; d: string; createdAt: number }
const SHARE_STORE = new Map<string, ShareRecord>();
const SHARE_CAP = 10;
const ALPHABET = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
function genShareId(n = 8): string {
  let out = '';
  for (let i = 0; i < n; i++) out += ALPHABET[Math.floor(Math.random() * ALPHABET.length)];
  return out;
}

export async function registerRoutes(app: Express): Promise<Server> {
  app.post('/api/share', (req: Request, res: Response) => {
    const body = (req.body && typeof req.body === 'object') ? req.body : {};
    const k = typeof body.k === 'string' ? body.k : '';
    const d = typeof body.d === 'string' ? body.d : '';
    if (!k || !d) {
      res.status(400).json({ error: 'k and d required' });
      return;
    }
    if (d.length > 200_000) {
      res.status(413).json({ error: 'payload too large' });
      return;
    }
    let id = genShareId();
    while (SHARE_STORE.has(id)) id = genShareId();
    SHARE_STORE.set(id, { k, d, createdAt: Date.now() });
    if (SHARE_STORE.size > SHARE_CAP) {
      const entries: Array<[string, ShareRecord]> = [];
      SHARE_STORE.forEach((rec, key) => { entries.push([key, rec]); });
      entries.sort((a, b) => a[1].createdAt - b[1].createdAt);
      while (SHARE_STORE.size > SHARE_CAP) {
        const ev = entries.shift();
        if (!ev) break;
        SHARE_STORE.delete(ev[0]);
      }
    }
    res.status(200).json({ id });
  });

  app.get('/api/share', (req: Request, res: Response) => {
    const idRaw = req.query.id;
    const id = Array.isArray(idRaw) ? idRaw[0] : idRaw;
    if (typeof id !== 'string' || !/^[A-Za-z0-9]{4,32}$/.test(id)) {
      res.status(400).json({ error: 'bad id' });
      return;
    }
    const rec = SHARE_STORE.get(id);
    if (!rec) {
      res.status(404).json({ error: 'not found' });
      return;
    }
    res.status(200).json({ k: rec.k, d: rec.d });
  });

  const httpServer = createServer(app);
  return httpServer;
}
