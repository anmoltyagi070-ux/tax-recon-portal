# GSTR Tax Reconciliation SPA - Project Context

This document outlines the architecture, file mapping, and execution instructions for the **Indirect Tax Reconciliation and Auditing Portal**.

---

## 1. Project Overview
A client-side Single Page Application (SPA) designed to parse 12 monthly GSTR-3B PDF files, consolidate records chronologically, perform multi-period aggregations/pivots, and map them to a formatted **Indirect Tax Reconciliation Matrix (Para D6)**.

---

## 2. Technology Stack
1. **Core**: HTML5, Vanilla JS (ES6).
2. **Styling**: Vanilla CSS (modern dark mode, custom grid styling, slide-down banners, responsive cards).
3. **Libraries**:
   - **PDF.js**: For parsing raw text and table coordinate data directly from PDF files in the browser.
   - **SheetJS (xlsx.js)**: For generating multi-sheet formatted Excel files with live calculations and formulas.
   - **Chart.js**: For rendering stacked bar visualizations of tax components.

---

## 3. Directory & File Mapping

```
tax_recon_web_app/
├── dist/
│   ├── index.html       # Single-page interface structure and tab panels
│   ├── style.css        # Executive dark mode styling and animations
│   ├── app.js           # PDF parsing engine, Comma Trap sanitization, pivots, & export formulas
│   ├── sample_data.js   # Simulated 12-month return files for demo mode
│   └── lib/             # Third-party runtime dependencies (PDF.js, SheetJS, Chart.js)
└── netlify.toml         # Hosting context (publish target, SPA routing redirects, security headers)
```

### Key Components:
- **PDF Extraction Engine (`app.js` -> `parseGstr3bPdf`)**:
  - Scans all pages of GSTR files.
  - Dynamically merges kern-split tokens (resolving comma splits).
  - Pads missing columns (handling coordinates missing blank Cess data).
  - Sanitizes numeric strings by stripping commas/spaces and replacing empty, hyphens (`-`), `Nil`, or `N/A` values with `0.00` before casting to float.
- **D6 Matrix Calculations (`app.js` -> `populateD6FromConsolidated`)**:
  - Automatically maps regular forward charges to Row 8.
  - Automatically maps RCM and Section 9(5) items to Row 10.
  - Aggregates ITC Utilised and Cash Ledger payments from Table 6.1 into Rows 15–22.
  - Performs live mathematical evaluations for totals and difference (variance) metrics.

---

## 4. Run Locally
Since the application uses PDF.js workers, it must be served via HTTP (not opened as a raw file path).

Start a simple web server from the project directory:
```bash
# Using python (standard built-in)
python3 -m http.server 8085
```
Then navigate to `http://localhost:8085/dist/`.

---

## 5. Deployment Context (Netlify)
A `netlify.toml` file is configured in the root directory. Because deployment requires interactive login and authentication, you can deploy the app to Netlify from your local terminal.

### Steps to Deploy:
1. Ensure the Netlify CLI is installed:
   ```bash
   npm install -g netlify-cli
   ```
2. Log in to your Netlify account:
   ```bash
   netlify login
   ```
3. Run the deployment command from the project root directory:
   ```bash
   netlify deploy --prod
   ```
   *When prompted, choose to create a new site or link an existing one. The CLI will automatically detect the settings in `netlify.toml` and publish the `dist` folder.*
