import { neon } from '@neondatabase/serverless';

const sql = neon(process.env.DATABASE_URL!);

const ALPHABET = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
function genId(n = 8): string {
  const bytes = new Uint8Array(n);
  crypto.getRandomValues(bytes);
  let out = '';
  for (let i = 0; i < n; i++) out += ALPHABET[bytes[i] % ALPHABET.length];
  return out;
}

interface Req {
  method?: string;
  query: Record<string, string | string[] | undefined>;
  body: any;
}
interface Res {
  status(code: number): Res;
  json(body: any): void;
  end(): void;
  setHeader(name: string, value: string): void;
}

export default async function handler(req: Req, res: Res) {
  try {
    if (req.method === 'GET') {
      const idRaw = req.query.id;
      const id = Array.isArray(idRaw) ? idRaw[0] : idRaw;
      if (!id || !/^[A-Za-z0-9]{4,32}$/.test(id)) {
        res.status(400).json({ error: 'bad id' });
        return;
      }
      const rows = (await sql`SELECT api_key AS k, data AS d FROM ponshares WHERE id = ${id}`) as any[];
      if (!rows[0]) {
        res.status(404).json({ error: 'not found' });
        return;
      }
      res.status(200).json({ k: rows[0].k, d: rows[0].d });
      return;
    }

    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {});
      const k = typeof body.k === 'string' ? body.k : '';
      const d = typeof body.d === 'string' ? body.d : '';
      if (!k || !d) {
        res.status(400).json({ error: 'k and d required' });
        return;
      }
      // 200 KB cap on compressed payload — generous; PRs over this hint at API misuse.
      if (d.length > 200_000) {
        res.status(413).json({ error: 'payload too large' });
        return;
      }
      // Retry on the astronomically rare id collision.
      let lastErr: any = null;
      for (let attempt = 0; attempt < 3; attempt++) {
        const id = genId();
        try {
          await sql`INSERT INTO ponshares (id, api_key, data) VALUES (${id}, ${k}, ${d})`;
          // Keep only the 20 most recent shares; evict the rest.
          await sql`DELETE FROM ponshares WHERE id NOT IN (SELECT id FROM ponshares ORDER BY created_at DESC LIMIT 20)`;
          res.status(200).json({ id });
          return;
        } catch (e: any) {
          lastErr = e;
          if (!String(e?.message || '').toLowerCase().includes('duplicate')) break;
        }
      }
      throw lastErr;
    }

    res.status(405).json({ error: 'method not allowed' });
  } catch (e: any) {
    console.error('share handler error', e);
    res.status(500).json({ error: 'internal error' });
  }
}
