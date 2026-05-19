-- Run this once in the new Neon database (Neon console → SQL Editor).
CREATE TABLE IF NOT EXISTS ponshares (
  id          TEXT PRIMARY KEY,
  api_key     TEXT NOT NULL,
  data        TEXT NOT NULL,
  created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
