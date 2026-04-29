create table public.whatsapp_inbound (
  id uuid default gen_random_uuid() primary key,
  payload jsonb not null,
  status text not null default 'pending',
  error text,
  created_at timestamp with time zone default timezone('utc'::text, now()) not null,
  processed_at timestamp with time zone
);

-- Habilitar RLS (opcional, mas recomendado para o Supabase)
alter table public.whatsapp_inbound enable row level security;

-- Apenas funções com permissão de service_role podem inserir
create policy "Service role can insert webhook events"
  on public.whatsapp_inbound for insert
  with check (true);

-- Apenas service role pode ler/atualizar
create policy "Service role can select webhook events"
  on public.whatsapp_inbound for select
  using (true);

create policy "Service role can update webhook events"
  on public.whatsapp_inbound for update
  using (true);
