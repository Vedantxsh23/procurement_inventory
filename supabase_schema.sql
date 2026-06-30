-- WeRoCon Lab — Procurement & Inventory System
-- Supabase schema. Run this once in SQL Editor.

-- ========== COMPONENTS (inventory) ==========
create table if not exists components (
  id bigint generated always as identity primary key,
  name text not null,
  category text default 'Other',
  item_type text default 'Non-Recurring',     -- 'Non-Recurring' | 'Recurring'
  qty integer default 1,
  unit_price numeric default 0,
  vendor text default '',
  gem_status text default 'Not checked',       -- 'Not checked' | 'Available on GeM' | 'Non-GeM certified'
  gem_search_ref text default '',
  invoice_no text default '',
  payment_status text default 'Pending',
  payment_method text default 'Not selected',
  payment_ref text default '',
  remarks text default '',
  fund_approval_ref text default '',
  pi_name text default '',
  project_title text default '',
  project_no text default '',
  created_by text default '',                  -- who added it (Vedant / Dr. Saurav Kumar etc.)
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

-- ========== FILES (quotations, invoices, certificates - store as Supabase Storage path) ==========
create table if not exists component_files (
  id bigint generated always as identity primary key,
  component_id bigint references components(id) on delete cascade,
  file_type text not null,                      -- 'quotation' | 'invoice' | 'non_gem_certificate' | 'payment_receipt' | 'po'
  file_name text not null,
  storage_path text not null,                   -- path inside Supabase Storage bucket
  uploaded_by text default '',
  uploaded_at timestamptz default now()
);

-- ========== TRACKING (shipment) ==========
create table if not exists shipment_tracking (
  id bigint generated always as identity primary key,
  component_id bigint references components(id) on delete cascade unique,
  courier text default '',
  tracking_id text default '',
  status text default 'Order placed',
  expected_delivery date,
  last_synced_at timestamptz,
  auto_tracked boolean default false,
  raw_status_payload jsonb,                     -- last raw response from courier API, for debugging
  updated_at timestamptz default now()
);

-- ========== TRACKING HISTORY (audit trail of status changes) ==========
create table if not exists tracking_history (
  id bigint generated always as identity primary key,
  component_id bigint references components(id) on delete cascade,
  status text not null,
  source text default 'manual',                 -- 'manual' | 'trackcourier'
  occurred_at timestamptz default now()
);

-- ========== DOCUMENT LOG (record every generated doc, for re-download / audit) ==========
create table if not exists document_log (
  id bigint generated always as identity primary key,
  doc_type text not null,                        -- 'fund-approval' | 'quotation' | 'non-gem' | 'payment-receipt' | 'bundle'
  ref_no text,
  component_ids bigint[] default '{}',
  generated_by text default '',
  created_at timestamptz default now()
);

-- ========== updated_at auto-touch ==========
create or replace function touch_updated_at() returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists trg_components_updated on components;
create trigger trg_components_updated before update on components
  for each row execute function touch_updated_at();

drop trigger if exists trg_tracking_updated on shipment_tracking;
create trigger trg_tracking_updated before update on shipment_tracking
  for each row execute function touch_updated_at();

-- ========== Row Level Security ==========
-- This app uses a shared app-level password gate (not Supabase Auth) per
-- earlier design decision, so RLS here just allows the anon/publishable
-- key to read & write — protection is the password screen in the app,
-- not the database. If you later add Supabase Auth (real per-person
-- login), tighten these policies to check auth.uid().

alter table components enable row level security;
alter table component_files enable row level security;
alter table shipment_tracking enable row level security;
alter table tracking_history enable row level security;
alter table document_log enable row level security;

create policy "allow all - components" on components for all using (true) with check (true);
create policy "allow all - component_files" on component_files for all using (true) with check (true);
create policy "allow all - shipment_tracking" on shipment_tracking for all using (true) with check (true);
create policy "allow all - tracking_history" on tracking_history for all using (true) with check (true);
create policy "allow all - document_log" on document_log for all using (true) with check (true);

-- ========== Realtime (so Dr. Saurav's screen updates live when you add/edit) ==========
alter publication supabase_realtime add table components;
alter publication supabase_realtime add table shipment_tracking;
