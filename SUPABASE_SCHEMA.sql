-- AnnotatePro online storage
-- Run this once in Supabase SQL Editor.
--
-- This first migration stores the existing dashboard state as JSONB.
-- It keeps the current React dashboard structure intact while moving
-- the data online. Authentication/role tables can be added next.

create table if not exists public.dashboard_state (
  id bigint primary key,
  data jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

alter table public.dashboard_state enable row level security;

-- Temporary setup policy:
-- The dashboard can read/write the single state row while you are
-- setting up the application. Replace these policies with authenticated
-- role-based policies when Login is added.
drop policy if exists "dashboard state read" on public.dashboard_state;
drop policy if exists "dashboard state write" on public.dashboard_state;

create policy "dashboard state read"
on public.dashboard_state
for select
to anon, authenticated
using (id = 1);

create policy "dashboard state write"
on public.dashboard_state
for all
to anon, authenticated
using (id = 1)
with check (id = 1);

insert into public.dashboard_state (id, data)
values (1, '{}'::jsonb)
on conflict (id) do nothing;
