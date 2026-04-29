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

create table if not exists public.user_settings (
  user_id text primary key,
  max_emails_in_summary int not null default 10,
  include_read_emails boolean not null default false,
  preferred_channel text not null default 'whatsapp',
  priority_senders text[] not null default '{}',
  updated_at timestamptz not null default now()
);

create index if not exists user_settings_updated_at_idx on public.user_settings (updated_at desc);

create table if not exists public.whatsapp_links (
  whatsapp_number text primary key,
  user_id text not null,
  updated_at timestamptz not null default now()
);

create unique index if not exists whatsapp_links_user_id_idx on public.whatsapp_links (user_id);
