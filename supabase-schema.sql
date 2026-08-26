-- Run this once in Supabase SQL Editor.
create table if not exists public.dashboard_state (
  id text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.dashboard_state enable row level security;

drop policy if exists "dashboard_state_public_read" on public.dashboard_state;
create policy "dashboard_state_public_read"
  on public.dashboard_state
  for select
  to anon, authenticated
  using (true);

drop policy if exists "dashboard_state_authenticated_insert" on public.dashboard_state;
create policy "dashboard_state_authenticated_insert"
  on public.dashboard_state
  for insert
  to authenticated
  with check (true);

drop policy if exists "dashboard_state_authenticated_update" on public.dashboard_state;
create policy "dashboard_state_authenticated_update"
  on public.dashboard_state
  for update
  to authenticated
  using (true)
  with check (true);

-- Optional: allow authenticated administrators to remove the shared snapshot.
drop policy if exists "dashboard_state_authenticated_delete" on public.dashboard_state;
create policy "dashboard_state_authenticated_delete"
  on public.dashboard_state
  for delete
  to authenticated
  using (true);
