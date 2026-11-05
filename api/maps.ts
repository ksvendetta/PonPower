import { neon } from "@neondatabase/serverless";
import { projectFromSharePayload, toCreatedAtIso, type ActiveMap } from "../shared/active-maps";

const sql = neon(process.env.DATABASE_URL!);

interface Req {
  method?: string;
}
interface Res {
  status(code: number): Res;
  json(body: unknown): void;
}

export default async function handler(req: Req, res: Res) {
  if (req.method !== "GET") {
    res.status(405).json({ error: "method not allowed" });
    return;
  }
  try {
    const rows = (await sql`
      SELECT id, data AS d, created_at
      FROM ponshares
      ORDER BY created_at DESC
      LIMIT 20
    `) as Array<{ id: string; d: string; created_at: unknown }>;
    const maps: ActiveMap[] = rows.map((row) => ({
      id: row.id,
      project: projectFromSharePayload(row.d),
      createdAt: toCreatedAtIso(row.created_at),
    }));
    res.status(200).json({ maps });
  } catch (e) {
    console.error("maps handler error", e);
    res.status(500).json({ error: "internal error" });
  }
}
