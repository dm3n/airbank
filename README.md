# Airbank — AI-Powered Quality of Earnings Platform

**Airbank** is an AI-native Quality of Earnings (QoE) platform built for CPAs, M&A advisors, and financial due diligence professionals. Upload financial documents, let the AI extract and normalize the data, then review a fully interactive workbook — complete with source citations, confidence scores, flag management, and export to Excel or Google Sheets.

**[→ Try the Live Sandbox](https://airbank-platform.vercel.app/)**

---

## What It Does

A traditional QoE engagement takes weeks of manual data wrangling. Airbank compresses that into minutes. Upload a company's financial documents — tax returns, audited statements, general ledger exports — and the AI engine extracts every line item, maps it to the right section of the QoE report, and flags anything that looks inconsistent.

The result is a living, auditable workbook where every cell traces back to its source document. CPAs can review, override, and annotate the data rather than entering it from scratch.

---

## Sandbox

A fully-populated demo workbook is available with live mock data — no signup required.

**[https://airbank-platform.vercel.app/](https://airbank-platform.vercel.app/)**

> Login with `user@test.com` / `test` to explore the complete platform with Alpine Outdoor Co. — a realistic mock engagement with 3 years of financials.

---

## Features

### AI Analysis Engine
- Upload PDFs, Excel files, and GL exports
- Gemini 2.0 Flash extracts financial data across 11 report sections
- RAG (Retrieval-Augmented Generation) via Google Cloud Vertex AI grounds every extraction in the source documents
- Confidence scores and source excerpts attached to every extracted cell
- AI-generated flags for low-confidence data, discrepancies, and anomalies
- Live SSE progress stream during analysis so you can watch the workbook populate in real time

### Interactive QoE Workbook
- **Complete QoE Report** — full normalized quality of earnings with sticky table of contents
- **Quality of Earnings** — EBITDA bridge with management and diligence adjustments
- **Income Statement** — multi-period P&L with gross margin and operating income analysis
- **Margins by Month** — trailing 12-month revenue / COGS / OpEx with trend charts
- **Balance Sheet** — normalized assets, liabilities, and equity across periods
- **Working Capital** — adjusted net working capital with seasonal analysis
- **Net Debt & Debt-Like Items** — full debt schedule including capital leases and deferred revenue
- **Sales by Channel** — revenue breakdown with mix charts
- **Customer Concentration** — top customer analysis with Herfindahl index
- **Proof of Revenue** — three-way match: GL vs. bank deposits vs. tax returns
- **Proof of Cash** — monthly reconciliation of bank deposits to recognized revenue
- **Three-Way Match** — bank / GL / tax variance analysis with PASS / REVIEW / FAIL thresholds
- **Risk & Diligence** — AP testing, accrual analysis, and diligence checklist

### Flag & Review System
- Flag any metric cell for review from its source popover
- Four flag types: Needs Review, Discrepancy, Low Confidence, AI Note
- Flagged values turn red with an inline flag list in the popover
- Global flags dropdown in the workbook header — click any flag to jump straight to it
- AI-generated flags surfaced automatically after analysis
- Full flag lifecycle: create → comment → resolve

### Auditability
- Every cell opens a source popover with document name, page number, and excerpt
- Confidence indicator on every AI-extracted value (high / medium / low)
- Edit any cell — overrides persist with a full audit trail
- "Ask Workbook AI" button injects the cell into the AI chat in context

### AI Chat Panel
- Workbook-scoped AI assistant in a persistent sidebar
- Answers grounded in uploaded documents via RAG
- Cell references automatically injected when you ask about a specific metric

### Export
- **Excel (.xlsx)** — fully formatted workbook download
- **Google Sheets** — publishes directly to a new spreadsheet in your Drive
- **PDF** — print-ready report

---

## Tech Stack

| Layer | Technology |
|---|---|
| Framework | Next.js 16 (App Router) + React 19 + TypeScript |
| Styling | Tailwind CSS v4 + shadcn/ui |
| Database | Supabase (PostgreSQL + Row Level Security) |
| Auth | Supabase Auth via `@supabase/ssr` (cookie sessions) |
| File Storage | Supabase Storage |
| Document Storage | Google Cloud Storage |
| AI / Extraction | Gemini 2.0 Flash via Vertex AI |
| RAG Engine | Vertex AI RAG API (corpus per workbook) |
| Charts | Recharts |
| Export | SheetJS (xlsx) + Google Sheets API |
| Deployment | Vercel |

---

## Architecture

```
┌─────────────────────────────────────────────────────────┐
│                    Next.js App Router                    │
│                                                         │
│  /dashboard                → Workbook list + sidebar    │
│  /dashboard/workbook/[id]  → Interactive workbook       │
│  /api/workbooks/[id]/analyze  → SSE analysis stream     │
│  /api/workbooks/[id]/cells    → Cells CRUD + flags      │
│  /api/workbooks/[id]/flags    → Flag management         │
└──────────────────────┬──────────────────────────────────┘
                       │
          ┌────────────┼─────────────┐
          ▼            ▼             ▼
     Supabase      Vertex AI      Google Cloud
     (DB + Auth    (Gemini 2.0    Storage
     + Storage)    + RAG Corpus)  (documents)
```

**Analysis flow:**
1. User uploads documents → Supabase Storage → GCS → RAG corpus import
2. User triggers analysis → SSE stream opens, progress events fire per section
3. For each of 11 sections: RAG query → Gemini extraction → cells + flags written to DB
4. On completion, workbook fetches all cells and populates the interactive report

---

## Database Schema

Core tables (all with Row Level Security):

```sql
workbooks          -- one per engagement
workbook_cells     -- extracted cells (section, row_key, period, value, confidence)
cell_flags         -- flags on cells (AI-generated or manual)
flag_comments      -- threaded comments per flag
audit_entries      -- edit history for cell overrides
source_documents   -- uploaded file metadata
missing_data       -- AI-identified data gaps
```

---

## Getting Started

### Prerequisites
- Node.js 20+
- Supabase project
- Google Cloud project with Vertex AI API enabled
- GCS bucket

### Environment Variables

```env
NEXT_PUBLIC_SUPABASE_URL=
NEXT_PUBLIC_SUPABASE_ANON_KEY=
SUPABASE_SERVICE_ROLE_KEY=
GCP_PROJECT_ID=
GCP_REGION=us-central1
GCS_BUCKET_NAME=
GOOGLE_APPLICATION_CREDENTIALS=   # path to service account JSON
GOOGLE_CLIENT_EMAIL=               # for Sheets export
GOOGLE_PRIVATE_KEY=
```

### Install & Run

```bash
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000).

### Database

Run `supabase-schema.sql` in your Supabase SQL editor to create all tables and RLS policies.

---

## Project Structure

```
app/
  page.tsx                             # Login
  dashboard/
    layout.tsx                         # Sidebar + AI chat panel
    page.tsx                           # Workbook list
    workbook/[id]/page.tsx             # Main workbook
  api/
    auth/                              # Signup + login
    workbooks/[id]/
      analyze/                         # SSE analysis stream
      cells/[cellId]/                  # Cell overrides
      flags/[flagId]/                  # Flag lifecycle
      flags/[flagId]/comments/         # Comment threads
      export/                          # Excel / Sheets / PDF

components/
  auditable-cell.tsx                   # Cell popover — source, flags, edit, AI
  complete-qoe-section.tsx             # Full QoE report component
  risk-diligence-section.tsx           # Risk & diligence section
  workbook-settings-dialog.tsx         # Settings + upload + re-analyze
  ai-chat-panel.tsx                    # Workbook AI sidebar
  document-viewer-panel.tsx            # Source document viewer

lib/
  gemini.ts                            # Gemini extraction + flag detection
  vertex-rag.ts                        # RAG corpus management
  gcs.ts                               # Google Cloud Storage
  section-prompts.ts                   # Per-section extraction prompts
  supabase-server.ts                   # Server Supabase client
```

---

## License

MIT
