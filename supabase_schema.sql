create extension if not exists pgcrypto;

create table if not exists public.segments (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  segment_id text not null,
  snippet text not null,
  theme text,
  theme_code text,
  theme_label text,
  open_code text,
  cipq_domain text check (cipq_domain in ('Creation', 'Production', 'Distribution', 'Access')),
  quadrant_primary text check (quadrant_primary in ('Creation', 'Production', 'Distribution', 'Access')),
  secondary_domain text,
  quadrant_secondary text,
  indicator_code text not null,
  indicator_name text not null,
  indicator_label text,
  severity integer not null check (severity between 1 and 5),
  stakeholder text,
  stakeholder_group text,
  respondent_type text,
  region text,
  source_type text,
  source_id text,
  record_confidence text check (record_confidence in ('low', 'medium', 'high')),
  is_cross_quadrant boolean not null default false,
  linked_quadrants text[] default '{}',
  analysis_notes text,
  session_id text,
  encoded_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now()),
  created_at timestamptz not null default timezone('utc', now()),
  unique (user_id, segment_id)
);

alter table public.segments
  add column if not exists theme_code text,
  add column if not exists theme_label text,
  add column if not exists quadrant_primary text,
  add column if not exists quadrant_secondary text,
  add column if not exists indicator_label text,
  add column if not exists stakeholder_group text,
  add column if not exists respondent_type text,
  add column if not exists source_id text,
  add column if not exists record_confidence text,
  add column if not exists is_cross_quadrant boolean not null default false,
  add column if not exists linked_quadrants text[] default '{}',
  add column if not exists analysis_notes text,
  add column if not exists encoded_at timestamptz not null default timezone('utc', now()),
  add column if not exists updated_at timestamptz not null default timezone('utc', now());

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'segments_quadrant_primary_check'
  ) then
    alter table public.segments
      add constraint segments_quadrant_primary_check
      check (quadrant_primary in ('Creation', 'Production', 'Distribution', 'Access'));
  end if;

  if not exists (
    select 1
    from pg_constraint
    where conname = 'segments_record_confidence_check'
  ) then
    alter table public.segments
      add constraint segments_record_confidence_check
      check (record_confidence in ('low', 'medium', 'high'));
  end if;
end $$;

create index if not exists segments_user_created_idx
  on public.segments (user_id, created_at);

create index if not exists segments_user_encoded_idx
  on public.segments (user_id, encoded_at);

alter table public.segments enable row level security;

drop policy if exists "segments_select_own" on public.segments;
create policy "segments_select_own"
on public.segments
for select
to authenticated
using ((select auth.uid()) = user_id);

drop policy if exists "segments_insert_own" on public.segments;
create policy "segments_insert_own"
on public.segments
for insert
to authenticated
with check ((select auth.uid()) = user_id);

drop policy if exists "segments_update_own" on public.segments;
create policy "segments_update_own"
on public.segments
for update
to authenticated
using ((select auth.uid()) = user_id)
with check ((select auth.uid()) = user_id);

drop policy if exists "segments_delete_own" on public.segments;
create policy "segments_delete_own"
on public.segments
for delete
to authenticated
using ((select auth.uid()) = user_id);
