<p align="center">
  <img src="https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white" alt="Apps Script"/>
  <img src="https://img.shields.io/badge/Gemini%20AI-8E75B2?style=for-the-badge&logo=google-gemini&logoColor=white" alt="Gemini"/>
  <img src="https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white" alt="Sheets"/>
  <img src="https://img.shields.io/badge/Google%20Drive-4285F4?style=for-the-badge&logo=google-drive&logoColor=white" alt="Drive"/>
</p>

<h1 align="center">🧾 SS Payments</h1>

<p align="center">
  <strong>Automated WhatsApp receipt scanner powered by Gemini AI</strong>
</p>

<p align="center">
  <em>Screenshots in → Structured expense data out → Keto tracking included</em>
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
```

1. **Share** payment receipts in your WhatsApp group
2. **Sync** the images to a Google Drive folder (via WhatsApp backup or manual upload)
3. **Run** the script from the Sheets menu — Gemini reads each image and extracts the data
4. **Done** — expenses are categorized, items are extracted, and keto classification is automatic

## Supported Receipt Types

| Type | Source | Fields Extracted |
|------|--------|-----------------|
| 🛒 **Boletas Electrónicas** | Store receipts (retail, food, pharmacy) | Store, date, items (name/qty/price), total, category |
| 🏦 **BCI Transfers** | "Operación exitosa" screenshots | Amount, recipient, bank, date |
| 🏦 **Scotiabank Transfers** | "Comprobante de Transferencia" | Amount, recipient, bank, date |

## Sheet Structure

The script creates and manages 5 tabs:

### Gastos (main expenses)

| Column | Description |
|--------|-------------|
| **Fecha** | Transaction date (YYYY-MM-DD) |
| **Tipo** | `boleta` or `transferencia` |
| **Comercio / Destinatario** | Store name or transfer recipient |
| **Categoria** | Auto-categorized (see below) |
| **Descripcion** | Item summary or transfer concept |
| **Total** | Total in CLP |
| **Banco Destino** | Recipient bank (transfers) |
| **Archivo** | Original filename |
| **File ID** | Drive file ID (dedup key) |
| **Procesado** | Processing timestamp |

### Incluidos (keto items)

Line items from boletas that match keto keywords in the dictionary.

| Column | Description |
|--------|-------------|
| **Fecha** | Receipt date |
| **Comercio** | Store name |
| **Categoria** | Expense category |
| **Item** | Product/service name |
| **Cantidad** | Quantity |
| **Precio Unitario** | Unit price (CLP) |
| **Subtotal** | Qty × unit price |
| **File ID** | Link back to receipt |

### Excluidos (non-keto items)

Same structure as Incluidos — items that don't match keto keywords or match as "NO".

### Diccionario Keto

User-editable keyword reference with ~90 pre-populated entries.

| Column | Description |
|--------|-------------|
| **Palabra Clave** | Keyword to match against item names (case-insensitive, substring match) |
| **Keto** | `SI` or `NO` |

Add, edit, or remove rows anytime. Changes take effect on the next processing run.

### Resumen (summary)

Auto-generated summary aggregating spending by category, split into keto vs non-keto totals. Refreshed automatically after processing or manually via **Recibos → Actualizar Resumen**.

### Auto-Categories

`Alimentacion` · `Farmacia` · `Vestuario` · `Hogar` · `Transferencia Personal` · `Arriendo` · `Servicios` · `Transporte` · `Otro`

## Setup

### 1. Get a Gemini API Key (free)

1. Go to [Google AI Studio](https://aistudio.google.com/apikey)
2. **Important**: Make sure your Google account has billing/tax info configured in [Google Cloud Console](https://console.cloud.google.com/billing). The free tier won't activate without it — you'll get `429` errors with `limit: 0`
3. Create an API key. If prompted, use the **default project** (not a manually created one)

> **Free tier**: 15 requests/min, 1M tokens/day — more than enough for personal use. You won't be charged.

### 2. Create the Google Sheet

1. Create a new Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete the default `Code.gs` content, then create 5 script files and paste the contents from this repo:

| File | Purpose |
|------|---------|
| `Config.gs` | Constants, sheet names, keto dictionary defaults, API key helpers |
| `DriveService.gs` | Drive folder scanning, image loading |
| `GeminiService.gs` | Gemini API integration, receipt + item extraction |
| `SheetService.gs` | Sheet operations, keto classification, summary generation |
| `Code.gs` | Menu, orchestrator, main processing loop |

4. Update `FOLDER_ID` in `Config.gs` with your Google Drive folder ID

### 3. Configure & Run

1. **Save** all files in the Apps Script editor (Ctrl+S)
2. **Refresh** the Google Sheet — a **"Recibos"** menu will appear
3. Click **Recibos → Configurar API Key** → paste your Gemini key
4. Click **Recibos → Procesar Recibos** → grant permissions when prompted
5. Watch the progress toasts as each receipt is processed

> **Deduplication is built-in** — running the script multiple times will only process new images.

### 4. (Optional) Auto-Processing

Set up a time-based trigger in Apps Script:

1. In the Apps Script editor, go to **Triggers** (clock icon)
2. Add trigger → `processReceipts` → Time-driven → choose interval (e.g., every hour)

## Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `429` with `limit: 0` | Free tier not activated | Complete billing/tax info in [GCP Console](https://console.cloud.google.com/billing) |
| `404` "model no longer available" | Deprecated model ID | Update `GEMINI_MODEL` in `Config.gs` |
| `403` Permission denied | OAuth scopes not granted | Re-run from menu, click "Allow" on all prompts |
| No "Recibos" menu | `onOpen()` not triggered | Save all files, refresh the sheet |
| No items in Incluidos/Excluidos | Transfers have no line items | Normal — only boletas produce item rows |

## Architecture

```
┌─────────────────────────────────────────────────────┐
│                     Code.gs                          │
│              (orchestrator + menu)                    │
│                                                       │
│  onOpen() ── processReceipts() ── updateResumen()    │
└────┬──────────────┬──────────────────┬───────────────┘
     │              │                  │
     ▼              ▼                  ▼
┌──────────┐ ┌─────────────┐   ┌──────────────┐
│DriveService│ │GeminiService│   │SheetService   │
│           │ │             │   │               │
│getImage   │ │extractReceipt│  │Gastos         │
│ Files()   │ │  Data()     │   │Incluidos      │
│getImage   │ │  + items[]  │   │Excluidos      │
│ Base64()  │ └──────┬──────┘   │Diccionario    │
└─────┬─────┘        │          │Resumen        │
      ▼              ▼          └───────┬───────┘
 Google Drive   Gemini API              ▼
                                  Google Sheet
                                   (5 tabs)
```

## Key Features

- **Zero infrastructure** — runs entirely in Google Apps Script
- **Smart extraction** — Gemini understands receipt context, extracts individual line items
- **Keto tracking** — auto-classifies items via editable keyword dictionary
- **Included/Excluded split** — separate tabs for keto vs non-keto items
- **Summary dashboard** — aggregated spending by category with keto breakdown
- **Structured output** — JSON schema enforcement ensures consistent data
- **Deduplication** — tracks processed files by Drive ID
- **Rate limiting** — 4.5s delay between API calls, batch size of 70
- **Error isolation** — one bad image won't stop the entire batch
- **Secure** — API key stored in Script Properties, never in code

## Security

> ⚠️ **Never hardcode your API key.** The key is stored in Google Apps Script's [Script Properties](https://developers.google.com/apps-script/guides/properties), which are encrypted at rest and not visible in the code editor.

## License

MIT — do whatever you want with it.

---

<p align="center">
  <sub>Built with 🤖 Claude Code + Gemini AI</sub>
</p>
