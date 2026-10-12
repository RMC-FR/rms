# Sales Ledger

Static web app (GitHub Pages) + Supabase (Auth + Postgres with row level security).

- `index.html` – the whole app, no build step
- `config.js` – Supabase Project URL + anon public key
- `db/schema.sql` – tables, policies, helper functions (no secrets, no data)

Accounts and business data are loaded from separate SQL files that are **not** stored in this repository.
