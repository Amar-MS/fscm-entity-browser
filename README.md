# F&SCM Entity Browser

A Microsoft Edge browser extension for browsing, inspecting, and querying data entities in **Dynamics 365 Finance & Supply Chain Management** — directly from your browser, no developer tools required.

---

## What It Does

| Feature | Details |
|---|---|
| **Entity list** | Browse and search all OData data entities in your D365 F&SCM environment |
| **Schema inspection** | View field names, types, key fields, and mandatory vs optional columns |
| **Live data query** | Query real data with pagination (100 records at a time) |
| **Visual filter builder** | Build OData filters with column / operator / value rows — no OData syntax knowledge needed |
| **Advanced filter** | Raw `$filter` input for power users |
| **Column-level search** | Filter by specific column values directly from the grid header |
| **Legal entity switcher** | Switch between companies — data is always scoped correctly |
| **Excel template download** | Download an import-ready `.xls` template with mandatory fields first and a field reference sheet |
| **CSV export** | Export current query results to CSV |
| **Metadata cache** | Schema cached for 7 days, entity list for 24 hours — fast after first load |

---

## Installation

> No build step. No npm. Works straight from the folder.

1. **Download** this repository — click **Code → Download ZIP** and extract it
2. Open **Microsoft Edge** and go to `edge://extensions`
3. Enable **Developer mode** (toggle in the top-right corner)
4. Click **Load unpacked** → select the extracted `fscm-table-browser` folder
5. The extension icon will appear in your toolbar

---

## How to Use

1. **Log in** to your D365 F&SCM environment in another Edge tab (must be signed in)
2. Click the **F&SCM Entity Browser** extension icon — a new tab opens
3. Click the environment badge (top right) and enter your environment URL, e.g.:
   ```
   https://yourenv.operations.dynamics.com
   ```
4. The entity list loads on the left — search or scroll to find an entity
5. Click an entity → use the **Schema** tab to inspect fields, **Data** tab to query records

---

## Requirements

- Microsoft Edge (Chromium)
- Access to a D365 F&SCM environment (cloud-hosted or production)
- Must be **logged in** to D365 in the same browser session

---

## How It Works

The extension injects a lightweight content script into your already-authenticated D365 tab. When you query data, requests are routed through that tab using your existing session — no credentials are stored, no external servers are involved. Everything stays within your browser.

---

## Tips

- Use **M / T badges** on the entity list to quickly identify Master data vs Transaction entities
- **Mandatory fields are highlighted** in the Excel template — hand it straight to a consultant for data migration
- If the entity list looks stale, click the **refresh icon** (↻) in the top bar to reload from the API
- To clear all filters at once, click the red **✕ Clear filters** button in the data status bar

---

## Disclaimer

This tool was born out of a personal experiment — built entirely through conversation with **GitHub Copilot** in VS Code, by someone with a business process background and no formal coding experience.

It is **not an official Microsoft product**. It is shared as-is, with no guarantees. Use it in non-production environments first and validate any data it returns against your source system.

You are encouraged to fork it, extend it, break it, and build your own version. That's the whole point.

---

## License

MIT — free to use, modify, and share.
