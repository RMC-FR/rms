# Sales Ledger

Static web app (GitHub Pages) + Google Sheet as the database via Google Apps Script.

- `index.html` – the whole app, no build step
- `config.js` – address of the Apps Script web app
- `gas/Code.gs` – backend (login, permissions, read/write the sheet). No data, no passwords.

The data/accounts seed file (`Seed.gs`) is **not** stored in this repository.
