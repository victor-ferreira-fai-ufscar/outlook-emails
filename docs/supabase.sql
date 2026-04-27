-- Schema minimo para persistencia do backend no Supabase

create table if not exists public.sessions (
  file_name text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

create table if not exists public.profiles (
  id uuid primary key default gen_random_uuid(),
  user_id text not null,
  path text not null,
  payload jsonb not null,
  created_at timestamptz not null default now()
);

create index if not exists profiles_user_id_idx on public.profiles (user_id);
create index if not exists profiles_created_at_idx on public.profiles (created_at desc);
