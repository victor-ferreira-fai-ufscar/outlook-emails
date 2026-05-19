create table public.sessions (
  file_name text not null,
  payload jsonb not null,
  updated_at timestamp with time zone not null default now(),
  constraint sessions_pkey primary key (file_name)
) TABLESPACE pg_default;