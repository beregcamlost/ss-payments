<p align="center">
  <img src="https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white" alt="Apps Script"/>
  <img src="https://img.shields.io/badge/Gemini%20AI-8E75B2?style=for-the-badge&logo=google-gemini&logoColor=white" alt="Gemini"/>
  <img src="https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white" alt="Sheets"/>
  <img src="https://img.shields.io/badge/Google%20Drive-4285F4?style=for-the-badge&logo=google-drive&logoColor=white" alt="Drive"/>
</p>

<h1 align="center">🧾 SS Payments</h1>

<p align="center">
  <strong>Automated expense tracker powered by Gemini AI</strong>
</p>

<p align="center">
  <em>Receipts, bank statements, PDFs → Structured expense data → Keto tracking included</em>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/currency-CLP%20🇨🇱-blue?style=flat-square" alt="CLP"/>
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat-square" alt="MIT"/>
  <img src="https://img.shields.io/badge/cost-$0%20(free%20tier)-brightgreen?style=flat-square" alt="Free"/>
</p>

---

## How It Works

```
📱 WhatsApp Group        📁 Google Drive         🤖 Gemini AI          📊 Google Sheet
 (receipt photos)  ──▶  (synced folder)    ──▶  (vision + JSON)  ──▶  (structured data)
🏦 Bank Statements       (CSV/Excel/PDF)         (auto-detect)         (reconciled)
```

1. **Upload** receipt images, bank statement CSVs/Excel, or PDFs to a Google Drive folder
2. **Run** the script from the Sheets menu — Gemini processes each file automatically
3. **Done** — bank transactions are matched to receipts, items are keto-classified, and summaries update

## Supported File Types

| Type | Format | Processing |
|------|--------|-----------|
| 🖼️ **Images** | JPEG, PNG, WEBP | Gemini vision extracts receipt/transfer data |
| 📄 **PDFs** | PDF | Sent directly to Gemini (native PDF support) |
| 📊 **CSV** | CSV | Auto-detect columns via Gemini, parse transactions |
| 📊 **Excel** | XLSX, XLS | Convert to Google Sheet via Drive API, then parse |

## Supported Receipt Types

| Type | Source | Fields Extracted |
|------|--------|-----------------|
| 🛒 **Boletas Electrónicas** | Store receipts (retail, food, pharmacy) | Store, date, items (name/qty/price), total, category |
| 🏦 **BCI Transfers** | "Operación exitosa" screenshots | Amount, recipient, bank, date |
| 🏦 **Scotiabank Transfers** | "Comprobante de Transferencia" | Amount, recipient, bank, date |
| 🏦 **Bank Statements** | CSV/Excel from any Chilean bank | Date, description, amount, auto-categorized |

## Three-Phase Processing

1. **Bank statements first** — CSVs and Excel files create transaction rows (Origen=`extracto`)
2. **Receipts second** — Images and PDFs create receipt rows with keto classification (Origen=`recibo`/`transferencia`)
3. **Reconciliation** — Matches receipts to bank transactions using fuzzy store matching + Gemini fallback

## Sheet Structure

The script creates and manages 7 tabs (6 visible + 1 hidden):

### Gastos (main expenses)

| Column | Description |
|--------|-------------|
| **Fecha** | Transaction date (YYYY-MM-DD) |
| **Tipo** | `boleta`, `transferencia`, or `extracto` |
| **Comercio / Destinatario** | Store name or transfer recipient |
| **Categoria** | Auto-categorized (see below) |
| **Descripcion** | Item summary or transfer concept |
| **Total** | Total in CLP ($#,##0) |
| **Archivo** | Original filename |
| **File ID** | Drive file ID (dedup key) |
| **Procesado** | Processing timestamp |
| **Origen** | `recibo`, `transferencia`, or `extracto` |
| **Estado** | `Conciliado`, `Sin Comprobante`, or `Sin Extracto` |

### Incluidos (keto items)

Line items from boletas classified as keto-friendly (`SI` in dictionary).

| Column | Description |
|--------|-------------|
| **Fecha** | Receipt date |
| **Comercio** | Store name |
| **Categoria** | Expense category |
| **Item** | Product/service name |
| **Cantidad** | Quantity |
| **Precio Unitario** | Unit price (CLP) |
| **Total** | Qty × unit price |

### Excluidos (non-keto food items)

Same structure as Incluidos — food/drink items classified as non-keto (`NO` in dictionary).

### No Comestible (non-food items)

Same structure as Incluidos — items that aren't food or drink (cleaning supplies, bags, hygiene products, etc.). Classified as `NO_FOOD` in the dictionary.

### Diccionario Keto

User-editable keyword reference with ~90 pre-populated entries. Gemini auto-learns new keywords when it encounters unknown items.

| Column | Description |
|--------|-------------|
| **Palabra Clave** | Keyword to match against item names (case-insensitive, longest-match-first) |
| **Keto** | `SI` (keto food), `NO` (non-keto food), or `NO_FOOD` (non-food item) |

Add, edit, or remove rows anytime. Changes take effect on the next processing run. Longer keywords take priority (e.g., "chocolate bitter" → SI matches before "chocolate" → NO).

### Resumen (summary)

Auto-generated summary with four sections:

| Section | Description |
|---------|-------------|
| **KETO (Incluidos)** | Spending on keto-friendly food by category |
| **NO KETO (Excluidos)** | Spending on non-keto food by category |
| **NO COMESTIBLE** | Spending on non-food items by category |
| **SIN DETALLE** | Spending from receipts without itemized detail |
| **TOTAL GENERAL** | Grand total across all sections |

All monetary values are formatted as CLP ($#,##0). Refreshed automatically after processing or manually via **Recibos → Actualizar Resumen**.

### Cache (hidden)

Hidden sheet that caches all Gemini API responses to avoid redundant calls. Keyed by `fileId + lastModified` so re-uploads are detected automatically.

### Auto-Categories

`Alimentacion` · `Farmacia` · `Vestuario` · `Hogar` · `Transferencia Personal` · `Arriendo` · `Servicios` · `Transporte` · `Otro`

## Bank Reconciliation

After processing both bank statements and receipts, the system matches them using two tiers:

1. **Deterministic match** — Fuzzy store name (normalized, accents stripped) + exact date + exact amount. No API calls.
2. **Gemini fallback** — For unmatched receipts, sends receipt data + candidate bank rows (similar amount ±10%) to Gemini to identify the match.

| Receipt | Bank Row | Estado |
|---------|----------|--------|
| Matched | Matched | `Conciliado` |
| None | Exists | `Sin Comprobante` |
| Exists | None | `Sin Extracto` |

## Menu Options

| Menu Item | Description |
|-----------|-------------|
| **Procesar Recibos** | Process all new files (images, PDFs, CSVs, Excel) |
| **Reprocesar Todo** | Clear all data and reprocess everything from scratch |
| **Actualizar Resumen** | Refresh only the Resumen tab |
| **Limpiar Datos** | Clear all data from all tabs (keeps headers, preserves cache) |
| **Invalidar Cache** | Clear all cached Gemini responses |
| **Configurar API Key** | Set your Gemini API key |

## Setup

### 1. Get a Gemini API Key (free)

1. Go to [Google AI Studio](https://aistudio.google.com/apikey)
2. **Important**: Make sure your Google account has billing/tax info configured in [Google Cloud Console](https://console.cloud.google.com/billing). The free tier won't activate without it — you'll get `429` errors with `limit: 0`
3. Create an API key. If prompted, use the **default project** (not a manually created one)

> **Free tier**: 15 requests/min, 1M tokens/day — more than enough for personal use. You won't be charged.

### 2. Create the Google Sheet

1. Create a new Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete the default `Code.gs` content, then create 8 script files and paste the contents from this repo:

| File | Purpose |
|------|---------|
| `Config.gs` | Constants, sheet names, keto dictionary defaults, API key helpers |
| `DriveService.gs` | Drive folder scanning, file loading, Excel conversion |
| `GeminiService.gs` | Gemini API integration, receipt extraction, CSV column detection, bank categorization, receipt matching |
| `SheetService.gs` | Sheet operations, keto classification, Gemini fallback, summary generation |
| `CacheService.gs` | Gemini response caching (hidden Cache sheet) |
| `CsvService.gs` | CSV/Excel bank statement parsing with auto column detection |
| `ReconciliationService.gs` | Two-tier receipt-to-bank-transaction matching |
| `Code.gs` | Menu, orchestrator, three-phase processing loop |

4. Update `FOLDER_ID` in `Config.gs` with your Google Drive folder ID
5. **Enable Drive API**: In the Apps Script editor, go to **Services** (+ icon) → Add **Drive API** (v2). Required for Excel file conversion.

### 3. Configure & Run

1. **Save** all files in the Apps Script editor (Ctrl+S)
2. **Refresh** the Google Sheet — a **"Recibos"** menu will appear
3. Click **Recibos → Configurar API Key** → paste your Gemini key
4. Click **Recibos → Procesar Recibos** → grant permissions when prompted
5. Watch the progress toasts as each file is processed

> **Deduplication is built-in** — running the script multiple times will only process new files. Bank transactions are deduped by date+description+amount hash.

### 4. (Optional) Auto-Processing

Set up a time-based trigger in Apps Script:

1. In the Apps Script editor, go to **Triggers** (clock icon)
2. Add trigger → `processReceipts` → Time-driven → choose interval (e.g., every hour)

## Smart Classification

Items from boletas are classified in three tiers:

1. **Dictionary match** (instant) — pre-compiled regex patterns match item names against Diccionario Keto keywords, longest match first
2. **Gemini fallback** (API call) — unknown items are batched and sent to Gemini via a shared `callGemini()` helper for classification as keto food, non-keto food, or non-food
3. **Auto-learn** — Gemini's answers are saved back to the dictionary and patterns are rebuilt in-memory, so subsequent receipts in the same batch (and future runs) skip Gemini for those items

## Gemini Response Cache

All Gemini API calls are cached in a hidden **Cache** sheet to minimize token usage:

- **Cache key**: `fileId + lastModified` (avoids loading large files into memory)
- **Cache types**: `receipt`, `csv_columns`, `csv_categories`, `matching`
- **Limpiar Datos** does NOT clear cache — cached responses remain valid
- **Invalidar Cache** clears only the Cache sheet — next run will re-call Gemini
- **Reprocesar Todo** clears data but preserves cache — fast reprocess from cached results

## Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `429` with `limit: 0` | Free tier not activated | Complete billing/tax info in [GCP Console](https://console.cloud.google.com/billing) |
| `404` "model no longer available" | Deprecated model ID | Update `GEMINI_MODEL` in `Config.gs` |
| `403` Permission denied | OAuth scopes not granted | Re-run from menu, click "Allow" on all prompts |
| No "Recibos" menu | `onOpen()` not triggered | Save all files, refresh the sheet |
| No items in Incluidos/Excluidos | Transfers have no line items | Normal — only boletas produce item rows |
| Items in wrong tab | Missing/wrong dictionary keyword | Edit Diccionario Keto, then Reprocesar Todo |
| Excel conversion fails | Drive API not enabled | Add Drive API v2 in Apps Script Services |
| CSV columns not detected | Unusual format | Gemini couldn't parse — check the CSV structure |

## Architecture

```
CSV/Excel ──▶ CsvService ──▶ Gemini (detect columns) ──▶ Gastos (Origen=extracto)
                                                                │
Images ──────▶ GeminiService ──▶ Gemini (extract receipt) ──────┤
                                                                │
PDF ─────────▶ GeminiService ──▶ Gemini (extract PDF) ──────────┤
                                                                ▼
                                                  ReconciliationService
                                                    matchAndMerge()
                                                    │           │
                                               Conciliado  Sin Comprobante
                                                    │
                                            Incluidos/Excluidos/No Comestible
                                                    │
                                                 Resumen

All Gemini calls ──▶ cachedCallGemini() ──▶ Cache hit? return : call API + store
```

## Key Features

- **Zero infrastructure** — runs entirely in Google Apps Script
- **Multi-format ingestion** — images, PDFs, CSVs, and Excel files from any Chilean bank
- **Smart extraction** — Gemini understands receipt context, extracts individual line items
- **Bank reconciliation** — matches receipts to bank transactions (fuzzy + AI)
- **AI-powered classification** — dictionary for known items, Gemini fallback for unknowns
- **Auto-learning dictionary** — Gemini results are saved for future runs
- **Response caching** — all Gemini calls cached to minimize token usage
- **Three-way item split** — keto food / non-keto food / non-food in separate tabs
- **Summary dashboard** — four-section Resumen with subtotals and grand total
- **CLP formatting** — all monetary values displayed as Chilean pesos ($#,##0)
- **One-click reprocess** — Reprocesar Todo clears and rebuilds everything
- **Structured output** — JSON schema enforcement ensures consistent data
- **Deduplication** — receipts by Drive ID, bank transactions by content hash
- **Rate limiting** — 4.5s delay between API calls, batch size of 35
- **Error isolation** — one bad file won't stop the entire batch
- **Secure** — API key stored in Script Properties, never in code

## Security

> ⚠️ **Never hardcode your API key.** The key is stored in Google Apps Script's [Script Properties](https://developers.google.com/apps-script/guides/properties), which are encrypted at rest and not visible in the code editor.

## License

MIT — do whatever you want with it.

---

<p align="center">
  <sub>Built with 🤖 Claude Code + Gemini AI</sub>
</p>
