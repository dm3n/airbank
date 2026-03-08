-- QoE Platform Database Schema
-- Run this in the Supabase SQL editor

-- Enable UUID extension
create extension if not exists "uuid-ossp";

-- ============================================================
-- WORKBOOKS
-- ============================================================
create table public.workbooks (
  id uuid primary key default uuid_generate_v4(),
  company_name text not null,
  status text not null default 'uploading'
    check (status in ('uploading','analyzing','needs_input','ready','error')),
  missing_fields jsonb default '[]'::jsonb,
  periods text[] default array['FY20','FY21','FY22','TTM'],
  created_by uuid references auth.users(id) on delete cascade,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

alter table public.workbooks enable row level security;

create policy "Users can view their own workbooks"
  on public.workbooks for select
  using (auth.uid() = created_by);

create policy "Users can insert their own workbooks"
  on public.workbooks for insert
  with check (auth.uid() = created_by);

create policy "Users can update their own workbooks"
  on public.workbooks for update
  using (auth.uid() = created_by);

create policy "Users can delete their own workbooks"
  on public.workbooks for delete
  using (auth.uid() = created_by);

-- ============================================================
-- DOCUMENTS
-- ============================================================
create table public.documents (
  id uuid primary key default uuid_generate_v4(),
  workbook_id uuid references public.workbooks(id) on delete cascade not null,
  file_name text not null,
  doc_type text check (doc_type in ('general_ledger','bank_statements','trial_balance','financials','other')),
  storage_path text,
  gcs_uri text,
  rag_file_id text,
  ingestion_status text not null default 'pending'
    check (ingestion_status in ('pending','uploading','ingesting','ready','error')),
  file_size bigint,
  content_type text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

alter table public.documents enable row level security;

create policy "Users can view documents for their workbooks"
  on public.documents for select
  using (exists (
    select 1 from public.workbooks w
    where w.id = documents.workbook_id and w.created_by = auth.uid()
  ));

create policy "Users can insert documents for their workbooks"
  on public.documents for insert
  with check (exists (
    select 1 from public.workbooks w
    where w.id = documents.workbook_id and w.created_by = auth.uid()
  ));

create policy "Users can update documents for their workbooks"
  on public.documents for update
  using (exists (
    select 1 from public.workbooks w
    where w.id = documents.workbook_id and w.created_by = auth.uid()
  ));

create policy "Users can delete documents for their workbooks"
  on public.documents for delete
  using (exists (
    select 1 from public.workbooks w
    where w.id = documents.workbook_id and w.created_by = auth.uid()
  ));

-- ============================================================
-- RAG CORPORA
-- ============================================================
create table public.rag_corpora (
  id uuid primary key default uuid_generate_v4(),
  workbook_id uuid references public.workbooks(id) on delete cascade not null unique,
  corpus_name text not null,
  created_at timestamptz default now()
);

alter table public.rag_corpora enable row level security;

create policy "Users can view RAG corpora for their workbooks"
  on public.rag_corpora for select
  using (exists (
    select 1 from public.workbooks w
    where w.id = rag_corpora.workbook_id and w.created_by = auth.uid()
  ));

create policy "Users can insert RAG corpora for their workbooks"
  on public.rag_corpora for insert
  with check (exists (
    select 1 from public.workbooks w
    where w.id = rag_corpora.workbook_id and w.created_by = auth.uid()
  ));

-- ============================================================
-- WORKBOOK CELLS
-- ============================================================
create table public.workbook_cells (
  id uuid primary key default uuid_generate_v4(),
  workbook_id uuid references public.workbooks(id) on delete cascade not null,
  section text not null,
  row_key text not null,
  period text not null,
  raw_value numeric,
  display_value text,
  is_calculated boolean default false,
  formula text,
  source_document_id uuid references public.documents(id) on delete set null,
  source_page integer,
  source_excerpt text,
  confidence numeric check (confidence >= 0 and confidence <= 1),
  is_overridden boolean default false,
  created_at timestamptz default now(),
  updated_at timestamptz default now(),
  unique (workbook_id, section, row_key, period)
);

alter table public.workbook_cells enable row level security;

create policy "Users can view cells for their workbooks"
  on public.workbook_cells for select
  using (exists (
    select 1 from public.workbooks w
    where w.id = workbook_cells.workbook_id and w.created_by = auth.uid()
  ));

create policy "Users can insert cells for their workbooks"
  on public.workbook_cells for insert
  with check (exists (
    select 1 from public.workbooks w
    where w.id = workbook_cells.workbook_id and w.created_by = auth.uid()
  ));

create policy "Users can update cells for their workbooks"
  on public.workbook_cells for update
  using (exists (
    select 1 from public.workbooks w
    where w.id = workbook_cells.workbook_id and w.created_by = auth.uid()
  ));

-- ============================================================
-- AUDIT ENTRIES
-- ============================================================
create table public.audit_entries (
  id uuid primary key default uuid_generate_v4(),
  cell_id uuid references public.workbook_cells(id) on delete cascade not null,
  workbook_id uuid references public.workbooks(id) on delete cascade not null,
  old_value numeric,
  new_value numeric,
  note text,
  edited_by uuid references auth.users(id) on delete set null,
  edited_at timestamptz default now()
);

alter table public.audit_entries enable row level security;

create policy "Users can view audit entries for their workbooks"
  on public.audit_entries for select
  using (exists (
    select 1 from public.workbooks w
    where w.id = audit_entries.workbook_id and w.created_by = auth.uid()
  ));

create policy "Users can insert audit entries for their workbooks"
  on public.audit_entries for insert
  with check (exists (
    select 1 from public.workbooks w
    where w.id = audit_entries.workbook_id and w.created_by = auth.uid()
  ));

-- ============================================================
-- MISSING DATA REQUESTS
-- ============================================================
create table public.missing_data_requests (
  id uuid primary key default uuid_generate_v4(),
  workbook_id uuid references public.workbooks(id) on delete cascade not null,
  section text not null,
  field_key text not null,
  period text,
  reason text,
  suggested_doc text,
  status text not null default 'open'
    check (status in ('open','resolved','skipped')),
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

alter table public.missing_data_requests enable row level security;

create policy "Users can view missing requests for their workbooks"
  on public.missing_data_requests for select
  using (exists (
    select 1 from public.workbooks w
    where w.id = missing_data_requests.workbook_id and w.created_by = auth.uid()
  ));

create policy "Users can insert missing requests for their workbooks"
  on public.missing_data_requests for insert
  with check (exists (
    select 1 from public.workbooks w
    where w.id = missing_data_requests.workbook_id and w.created_by = auth.uid()
  ));

create policy "Users can update missing requests for their workbooks"
  on public.missing_data_requests for update
  using (exists (
    select 1 from public.workbooks w
    where w.id = missing_data_requests.workbook_id and w.created_by = auth.uid()
  ));

-- ============================================================
-- STORAGE BUCKET
-- ============================================================
-- Create in Supabase Dashboard → Storage → New bucket
-- Name: qoe-documents
-- Private: true
-- Path pattern: {workbook_id}/{doc_id}/{filename}

-- ============================================================
-- UPDATED_AT TRIGGER
-- ============================================================
create or replace function public.handle_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

create trigger handle_workbooks_updated_at
  before update on public.workbooks
  for each row execute procedure public.handle_updated_at();

create trigger handle_documents_updated_at
  before update on public.documents
  for each row execute procedure public.handle_updated_at();

create trigger handle_workbook_cells_updated_at
  before update on public.workbook_cells
  for each row execute procedure public.handle_updated_at();

create trigger handle_missing_data_requests_updated_at
  before update on public.missing_data_requests
  for each row execute procedure public.handle_updated_at();

-- ============================================================
-- INDEXES
-- ============================================================
create index idx_documents_workbook_id on public.documents(workbook_id);
create index idx_workbook_cells_workbook_id on public.workbook_cells(workbook_id);
create index idx_workbook_cells_section on public.workbook_cells(workbook_id, section);
create index idx_audit_entries_cell_id on public.audit_entries(cell_id);
create index idx_audit_entries_workbook_id on public.audit_entries(workbook_id);
create index idx_missing_data_requests_workbook_id on public.missing_data_requests(workbook_id);

-- ============================================================
-- CELL FLAGS
-- ============================================================
create table public.cell_flags (
  id uuid primary key default gen_random_uuid(),
  workbook_id uuid not null references public.workbooks(id) on delete cascade,
  cell_id uuid references public.workbook_cells(id) on delete set null,
  section text not null,
  row_key text not null,
  period text,
  flag_type text not null default 'needs_review'
    check (flag_type in ('low_confidence','discrepancy','missing','ai_note','needs_review')),
  severity text not null default 'warning'
    check (severity in ('info','warning','critical')),
  title text not null,
  body text,
  created_by_ai boolean not null default false,
  resolved_at timestamptz,
  resolved_by uuid references auth.users(id),
  created_at timestamptz not null default now()
);

alter table public.cell_flags enable row level security;

create policy "Users manage own flags"
  on public.cell_flags for all
  using (workbook_id in (select id from public.workbooks where created_by = auth.uid()));

create index cell_flags_workbook_idx on public.cell_flags(workbook_id);
create index cell_flags_cell_idx on public.cell_flags(cell_id);

-- ============================================================
-- FLAG COMMENTS
-- ============================================================
create table public.flag_comments (
  id uuid primary key default gen_random_uuid(),
  flag_id uuid not null references public.cell_flags(id) on delete cascade,
  workbook_id uuid not null references public.workbooks(id) on delete cascade,
  author_id uuid references auth.users(id),
  author_name text not null,
  body text not null,
  created_at timestamptz not null default now()
);

alter table public.flag_comments enable row level security;

create policy "Users manage own comments"
  on public.flag_comments for all
  using (workbook_id in (select id from public.workbooks where created_by = auth.uid()));

create index flag_comments_flag_idx on public.flag_comments(flag_id);
