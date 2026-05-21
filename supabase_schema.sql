-- CIPQ Segments Table
-- Idempotent: safe to run multiple times on an existing database.

create extension if not exists pgcrypto;

create table if not exists public.segments (
  id                uuid        primary key default gen_random_uuid(),
  user_id           uuid        not null references auth.users(id) on delete cascade,
  segment_id        text        not null,
  snippet           text        not null,
  theme             text,
  theme_code        text,
  theme_label       text,
  open_code         text,
  cipq_domain       text        check (cipq_domain in ('Creation', 'Production', 'Distribution', 'Access')),
  quadrant_primary  text        check (quadrant_primary in ('Creation', 'Production', 'Distribution', 'Access')),
  secondary_domain  text,
  quadrant_secondary text,
  indicator_code    text        not null,
  indicator_name    text        not null,
  indicator_label   text,
  severity          integer     not null check (severity between 1 and 5),
  stakeholder       text,
  stakeholder_group text,
  respondent_type   text,
  region            text,
  source_type       text,
  source_id         text,
  value_chain_stage text        check (value_chain_stage in ('Development', 'Production', 'Distribution', 'Market Access')),
  pestle_tags       text[]      default '{}'
                                check (pestle_tags <@ array['Political','Economic','Social','Technological','Legal','Environmental']::text[]),
  record_confidence text        check (record_confidence in ('low', 'medium', 'high')),
  is_cross_quadrant boolean     not null default false,
  linked_quadrants  text[]      default '{}',
  analysis_notes    text,
  session_id        text,
  encoded_at        timestamptz not null default timezone('utc', now()),
  updated_at        timestamptz not null default timezone('utc', now()),
  created_at        timestamptz not null default timezone('utc', now()),

  unique (user_id, segment_id)
);

-- Indexes
create index if not exists segments_user_created_idx on public.segments (user_id, created_at);
create index if not exists segments_user_encoded_idx on public.segments (user_id, encoded_at);

-- Row-level security
alter table public.segments enable row level security;

-- Drop all existing policies before recreating
drop policy if exists "segments_select_own"    on public.segments;
drop policy if exists "segments_insert_own"    on public.segments;
drop policy if exists "segments_update_own"    on public.segments;
drop policy if exists "segments_delete_own"    on public.segments;
drop policy if exists "segments_select_public" on public.segments;

-- Authenticated users: full CRUD on their own rows
create policy "segments_select_own"
  on public.segments for select to authenticated
  using ((select auth.uid()) = user_id);

create policy "segments_insert_own"
  on public.segments for insert to authenticated
  with check ((select auth.uid()) = user_id);

create policy "segments_update_own"
  on public.segments for update to authenticated
  using ((select auth.uid()) = user_id)
  with check ((select auth.uid()) = user_id);

create policy "segments_delete_own"
  on public.segments for delete to authenticated
  using ((select auth.uid()) = user_id);

-- Anonymous users: read-only view of all published segments (guest mode)
create policy "segments_select_public"
  on public.segments for select to anon
  using (true);

-- ─────────────────────────────────────────────────────────────────────────────
-- SURVEY MODULE TABLES
-- Analytically separate from CIPQ. Never joined into CIPQ aggregations.
-- ─────────────────────────────────────────────────────────────────────────────

create table if not exists public.survey_questions (
  id          uuid  primary key default gen_random_uuid(),
  code        text  not null unique,
  text        text  not null,
  cipq_domain text  check (cipq_domain in ('Creation','Production','Distribution','Access')),
  category    text,
  created_at  timestamptz not null default timezone('utc', now())
);

create table if not exists public.survey_responses (
  id               uuid    primary key default gen_random_uuid(),
  respondent_id    text,
  respondent_group text,
  region           text,
  source_id        text,
  question_code    text    not null references public.survey_questions(code) on delete cascade,
  score            integer not null check (score between 1 and 5),
  recorded_at      timestamptz not null default timezone('utc', now())
);

create index if not exists survey_responses_qcode_idx on public.survey_responses (question_code);
create index if not exists survey_responses_group_idx on public.survey_responses (respondent_group);

-- RLS
alter table public.survey_questions  enable row level security;
alter table public.survey_responses  enable row level security;

drop policy if exists "survey_questions_select_all"  on public.survey_questions;
drop policy if exists "survey_questions_write"       on public.survey_questions;
drop policy if exists "survey_responses_select_all"  on public.survey_responses;
drop policy if exists "survey_responses_write"       on public.survey_responses;

-- Anyone (including guests) can read survey data
create policy "survey_questions_select_all"
  on public.survey_questions for select to authenticated, anon using (true);

create policy "survey_responses_select_all"
  on public.survey_responses for select to authenticated, anon using (true);

-- Only authenticated users can write
create policy "survey_questions_write"
  on public.survey_questions for all to authenticated using (true) with check (true);

create policy "survey_responses_write"
  on public.survey_responses for all to authenticated using (true) with check (true);