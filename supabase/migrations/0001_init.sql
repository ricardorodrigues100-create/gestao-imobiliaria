-- Gestão Imobiliária — schema inicial
-- Corre isto no SQL Editor do teu projeto Supabase (https://supabase.com/dashboard/project/ldlgxnalskrehfwmdqqv/sql/new)
-- se a Claude não tiver conseguido aplicar a migração automaticamente.

create extension if not exists pgcrypto;

create table if not exists public.imob_properties (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  name text not null,
  type text not null default 'long_term_rental',
  address text default '',
  purchase_value numeric default 0,
  current_value numeric,
  purchase_date date,
  status text default 'Ativo',
  notes text default '',
  loans jsonb not null default '[]',
  insurances jsonb not null default '[]',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.imob_transactions (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  property_id uuid references public.imob_properties(id) on delete set null,
  type text not null check (type in ('income','expense')),
  category text not null,
  amount numeric not null default 0,
  date date not null,
  description text default '',
  recurring text default 'none',
  reservation_id uuid,
  created_at timestamptz not null default now()
);

create table if not exists public.imob_reservations (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  property_id uuid references public.imob_properties(id) on delete set null,
  guest_name text default '',
  platform text default 'Airbnb',
  check_in date,
  check_out date,
  gross_amount numeric default 0,
  platform_fee numeric default 0,
  cleaning_fee numeric default 0,
  notes text default '',
  created_at timestamptz not null default now()
);

create table if not exists public.imob_settings (
  owner_id uuid primary key default auth.uid() references auth.users(id) on delete cascade,
  inflation_rate numeric not null default 2.5,
  custom_income_categories text[] not null default '{}',
  custom_expense_categories text[] not null default '{}',
  updated_at timestamptz not null default now()
);

alter table public.imob_properties enable row level security;
alter table public.imob_transactions enable row level security;
alter table public.imob_reservations enable row level security;
alter table public.imob_settings enable row level security;

drop policy if exists "own rows select" on public.imob_properties;
drop policy if exists "own rows insert" on public.imob_properties;
drop policy if exists "own rows update" on public.imob_properties;
drop policy if exists "own rows delete" on public.imob_properties;
create policy "own rows select" on public.imob_properties for select using (owner_id = auth.uid());
create policy "own rows insert" on public.imob_properties for insert with check (owner_id = auth.uid());
create policy "own rows update" on public.imob_properties for update using (owner_id = auth.uid()) with check (owner_id = auth.uid());
create policy "own rows delete" on public.imob_properties for delete using (owner_id = auth.uid());

drop policy if exists "own rows select" on public.imob_transactions;
drop policy if exists "own rows insert" on public.imob_transactions;
drop policy if exists "own rows update" on public.imob_transactions;
drop policy if exists "own rows delete" on public.imob_transactions;
create policy "own rows select" on public.imob_transactions for select using (owner_id = auth.uid());
create policy "own rows insert" on public.imob_transactions for insert with check (owner_id = auth.uid());
create policy "own rows update" on public.imob_transactions for update using (owner_id = auth.uid()) with check (owner_id = auth.uid());
create policy "own rows delete" on public.imob_transactions for delete using (owner_id = auth.uid());

drop policy if exists "own rows select" on public.imob_reservations;
drop policy if exists "own rows insert" on public.imob_reservations;
drop policy if exists "own rows update" on public.imob_reservations;
drop policy if exists "own rows delete" on public.imob_reservations;
create policy "own rows select" on public.imob_reservations for select using (owner_id = auth.uid());
create policy "own rows insert" on public.imob_reservations for insert with check (owner_id = auth.uid());
create policy "own rows update" on public.imob_reservations for update using (owner_id = auth.uid()) with check (owner_id = auth.uid());
create policy "own rows delete" on public.imob_reservations for delete using (owner_id = auth.uid());

drop policy if exists "own rows select" on public.imob_settings;
drop policy if exists "own rows insert" on public.imob_settings;
drop policy if exists "own rows update" on public.imob_settings;
create policy "own rows select" on public.imob_settings for select using (owner_id = auth.uid());
create policy "own rows insert" on public.imob_settings for insert with check (owner_id = auth.uid());
create policy "own rows update" on public.imob_settings for update using (owner_id = auth.uid()) with check (owner_id = auth.uid());
