# Multi-Format Ingestion, Bank Reconciliation & Gemini Cache

**Date:** 2026-03-14
**Status:** Approved

## Goal

Extend ss-payments to process any file type (images, PDFs, CSVs, Excel) from any Chilean bank. Correlate bank statement transactions with receipt detail. Cache all Gemini responses to minimize token usage.

## Current State

- Processes only JPEG/PNG/WEBP images from a Drive folder
- Extracts receipt data via Gemini vision
- Classifies items as keto/non-keto/non-food
- No bank statement support, no caching

## Design

### 1. File Ingestion

`DriveService.gs` accepts all file types from the Drive folder, tagging each with a processing strategy:

| File Type | Strategy | Processing |
|---|---|---|
| JPEG/PNG/WEBP | `image` | Gemini vision (unchanged) |
| PDF | `pdf` | Send PDF blob directly to Gemini (supports `application/pdf` natively) |
| CSV | `csv` | Auto-detect columns with Gemini, parse via `Utilities.parseCsv()` |
| XLSX/XLS | `excel` | Convert to Google Sheet via Drive API, read as structured data |

`getImageFiles()` becomes `getFiles()`, returning all files with a `strategy` field.

**PDF handling:** Gemini accepts `application/pdf` as inline data, same as images. No conversion needed — send the base64-encoded PDF blob directly.

**Excel handling:** Upload the Excel blob to Drive as a Google Sheet (`MimeType.GOOGLE_SHEETS`), read the data via SpreadsheetApp, then convert to CSV-equivalent rows. This leverages Apps Script's native Sheet reading.

### 2. CSV Processing

1. **Column detection** — first 5 rows sent to Gemini: identify date, description, amount, bank name columns + their formats. If Gemini fails to detect or the file has fewer than 2 rows, skip the file and log a warning.
2. **Row parsing** — extract all transactions using detected mapping
3. **Category assignment** — batch all descriptions, send to Gemini for categorization (same pattern as keto classification)
4. **Each row becomes a Gastos entry** with `Origen='extracto'`

New file `CsvService.gs` handles CSV parsing, column detection, and row extraction. Excel files are converted to row arrays by `DriveService.gs` before reaching `CsvService.gs`.

**Data flow:** `DriveService.getFiles()` → `CsvService.parseCsvFile()` (returns raw transaction objects) → `SheetService.buildBankRow()` (formats each into a Gastos row) → `SheetService.flushReceiptRows()` (writes to sheet).

### 3. Receipt-to-Transaction Matching

Two-tier matching after both CSVs and receipts are processed:

**Tier 1 — Deterministic:** Fuzzy store name (normalized, lowercased, branch codes stripped) + exact date + exact amount. No API calls.

**Tier 2 — Gemini fallback:** For unmatched receipts, send receipt data + candidate CSV rows (same date range, similar amount) to Gemini to identify the match.

New file `ReconciliationService.gs` handles matching logic, keeping SheetService focused on I/O.

#### Match outcomes

| Receipt | CSV Row | Result |
|---|---|---|
| Matched | Matched | CSV row enriched with receipt File ID. Items classified into Incluidos/Excluidos/No Comestible. Estado=`Conciliado` |
| None | Exists | Gastos row stays. Estado=`Sin Comprobante` |
| Exists | None | Receipt creates its own Gastos row. Estado=`Sin Extracto` |

### 4. Gastos Column Layout

Full HEADERS array with new columns appended at end:

| Index | Column | Description |
|---|---|---|
| A | Fecha | Transaction date (YYYY-MM-DD) |
| B | Tipo | `boleta` or `transferencia` |
| C | Comercio / Destinatario | Store name or transfer recipient |
| D | Categoria | Auto-categorized |
| E | Descripcion | Item summary or transfer concept |
| F | Total | Total in CLP ($#,##0) |
| G | Archivo | Original filename |
| H | File ID | Drive file ID (dedup key) |
| I | Procesado | Processing timestamp |
| J | Origen | `recibo`, `transferencia`, `extracto` |
| K | Estado | `Conciliado`, `Sin Comprobante`, `Sin Extracto` |

**Origen values** (distinct from Tipo to avoid ambiguity):
- `recibo` — from a receipt image/PDF
- `transferencia` — from a transfer screenshot
- `extracto` — from a CSV/Excel bank statement

**Row builders:**
- `buildReceiptRow()` — updated to return 11 values, sets Origen=`recibo` or `transferencia`, Estado=`Sin Extracto` (updated to `Conciliado` after matching)
- `buildBankRow()` — new function for CSV rows, sets Origen=`extracto`, Estado=`Sin Comprobante` (updated to `Conciliado` after matching). Uses bank description for Archivo, CSV File ID for File ID.

### 5. Processing Order

1. Process CSVs/Excel first (creates bank transaction rows)
2. Process receipts (images/PDFs)
3. Match & merge (link receipts to CSV rows, classify items)
4. Refresh Resumen

**Deduplication:**
- CSVs: MD5 hash of `date + description + amount` stored in a **"Dedup"** column in the Cache sheet (tipo=`csv_dedup`). Bank name is detected from the CSV content by Gemini during column detection.
- Images/PDFs: existing File ID dedup
- Matched receipts enrich existing CSV rows instead of creating duplicates

### 6. Gemini Response Cache

New hidden **"Cache"** sheet tab:

| Column | Description |
|---|---|
| Hash | `fileId + ':' + lastModified` (avoids loading entire file into memory) |
| Tipo | `receipt`, `csv_columns`, `csv_categories`, `matching`, `csv_dedup` |
| Resultado | JSON string of Gemini response |
| Timestamp | Cache creation time |

**Flow:** Before any Gemini call, compute hash from file metadata → check Cache → hit: return cached → miss: call API, store result, return.

**Cache key strategy:** Uses `fileId + ':' + file.getLastUpdated()` instead of content MD5 to avoid loading large files into memory (Apps Script has ~50MB limit). If the file is re-uploaded with the same content, it gets a new File ID so the cache naturally misses.

**Cache lifecycle:**
- "Limpiar Datos" does NOT clear cache (cached Gemini responses are still valid)
- "Invalidar Cache" clears only the Cache sheet
- "Reprocesar Todo" clears data but preserves cache (fast reprocess from cached results)
- No automatic TTL/expiration — "Invalidar Cache" is the cleanup mechanism

**Rate limiting:** All Gemini calls go through `callGemini()` which already respects `RATE_LIMIT_DELAY`. Batched operations (category assignment, keto classification) minimize total call count.

### 7. Error Handling

| Scenario | Behavior |
|---|---|
| CSV column detection fails | Skip file, log warning, continue with next file |
| CSV has < 2 rows | Skip file, log warning |
| Excel conversion fails | Skip file, log warning |
| PDF extraction fails | Same as current image failure — skip, increment error count |
| Gemini quota exhausted (429) | Stop processing, flush partial progress, show error toast |
| Cache sheet corrupted | Treat as cache miss, re-call Gemini |
| Matching finds multiple candidates | Use best match (highest similarity), log alternatives |

### 8. File Changes

| File | Changes |
|---|---|
| **Config.gs** | `CACHE_SHEET_NAME`, expanded MIME types, updated HEADERS (add Origen/Estado), `ORIGEN` and `ESTADO` value constants |
| **DriveService.gs** | `getFiles()` replaces `getImageFiles()`, returns all files with `strategy`. New `getFileHash()`. Excel-to-rows conversion |
| **GeminiService.gs** | `detectCsvColumns()`, `categorizeBankDescriptions()`, `matchReceiptToTransaction()` — all via `callGemini()` |
| **NEW: CacheService.gs** | `cachedCallGemini()` wrapper, `getCached()`, `putCache()`, `clearCache()`. Sits between callers and `callGemini()` to avoid circular dependencies |
| **SheetService.gs** | `buildBankRow()`, updated `buildReceiptRow()` (11 cols), `refreshResumen()` |
| **Code.gs** | Processing loop: CSVs first, receipts second, match, Resumen. New menu items: `Invalidar Cache`. Updated `reprocesarTodo()` |
| **NEW: CsvService.gs** | CSV parsing via `Utilities.parseCsv()`, column detection orchestration, row extraction |
| **NEW: ReconciliationService.gs** | `matchAndMerge()`, fuzzy store matching, Gemini fallback matching |

### 9. Menu Structure

```
Recibos
  Procesar Recibos
  Reprocesar Todo
  Actualizar Resumen
  ─────────────
  Limpiar Datos
  Invalidar Cache
  Configurar API Key
```

### 10. Architecture

```
CSV/Excel ──> CsvService ──────> Gemini (detect columns) ──> Gastos (Origen=extracto)
                                                                   |
Images ──────> GeminiService ──> Gemini (extract receipt) ─────────|
                                                                   |
PDF ─────────> GeminiService ──> Gemini (extract PDF) ─────────────|
                                                                   v
                                                     ReconciliationService
                                                       matchAndMerge()
                                                       |         |
                                                  Conciliado  Sin Comprobante
                                                       |
                                               Incluidos/Excluidos/No Comestible
                                                       |
                                                    Resumen

All Gemini calls ──> cachedCallGemini() ──> Cache hit? return : call API + store
```
