# Sales Ledger

Static web app (GitHub Pages) + Google Sheet as the database via Google Apps Script.

- `index.html` – the whole app, no build step
- `config.js` – address of the Apps Script web app
- `gas/Code.gs` – backend (login, permissions, read/write the sheet). No data, no passwords.

The data/accounts seed file (`Seed.gs`) is **not** stored in this repository.

## Invoice photos (Gemini free tier)
Upload a photo or PDF in the entry form: it is shrunk in the browser, saved to `Sales Ledger Invoices/<branch>` in the owner's Drive, and read by the Gemini API (key in Script Property `GEMINI_KEY`, never in this repo). The managers review the lines before saving; the items go to the `items` tab. Run `authorize()` once in the Apps Script editor after adding Drive/UrlFetch code.
