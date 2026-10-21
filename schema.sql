-- Protocolo de Iluminación · esquema de base de datos
-- Pegar completo en Supabase → SQL Editor → New query → Run.

create table if not exists public.mediciones (
  id          uuid primary key,
  user_id     uuid not null default auth.uid() references auth.users(id) on delete cascade,
  empresa     text,
  fecha       date,
  data        jsonb not null,
  deleted     boolean not null default false,
  updated_at  timestamptz not null default now(),
  created_at  timestamptz not null default now()
);

create index if not exists mediciones_user_idx on public.mediciones (user_id, updated_at desc);

-- Seguridad: cada usuario solo puede ver y modificar sus propias mediciones.
alter table public.mediciones enable row level security;

drop policy if exists "ver propias"       on public.mediciones;
drop policy if exists "crear propias"     on public.mediciones;
drop policy if exists "modificar propias" on public.mediciones;
drop policy if exists "borrar propias"    on public.mediciones;

create policy "ver propias"       on public.mediciones for select using (auth.uid() = user_id);
create policy "crear propias"     on public.mediciones for insert with check (auth.uid() = user_id);
create policy "modificar propias" on public.mediciones for update using (auth.uid() = user_id) with check (auth.uid() = user_id);
create policy "borrar propias"    on public.mediciones for delete using (auth.uid() = user_id);
