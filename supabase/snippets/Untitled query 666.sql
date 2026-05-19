create table public.whatsapp_links (
  whatsapp_number text not null,
  user_id uuid null,
  updated_at timestamp with time zone null default now(),
  constraint whatsapp_links_pkey primary key (whatsapp_number)
) TABLESPACE pg_default;