-- Vendedor 360° — gestão de preços, vigência e auditoria
-- Execute uma única vez no SQL Editor do Supabase.

alter table public.products
  add column if not exists valid_from date,
  add column if not exists valid_until date,
  add column if not exists pricing_mode text not null default 'manual'
    check (pricing_mode in ('automatic', 'manual'));

create table if not exists public.product_price_history (
  id uuid primary key default gen_random_uuid(),
  product_id uuid not null references public.products(id) on delete cascade,
  product_code text not null,
  changed_by uuid references auth.users(id) on delete set null,
  changed_at timestamptz not null default now(),
  change_source text not null default 'manual'
    check (change_source in ('manual', 'import', 'system')),
  old_values jsonb,
  new_values jsonb not null
);

create index if not exists product_price_history_product_idx
  on public.product_price_history(product_id, changed_at desc);

create index if not exists products_validity_idx
  on public.products(valid_from, valid_until);

alter table public.product_price_history enable row level security;

drop policy if exists "product_price_history_read_managers"
  on public.product_price_history;
create policy "product_price_history_read_managers"
  on public.product_price_history
  for select to authenticated
  using (public.my_can_manage_catalog());

-- A gravação do histórico ocorre pelo trigger, não diretamente pelo navegador.
revoke insert, update, delete on public.product_price_history from anon, authenticated;
grant select on public.product_price_history to authenticated;

create or replace function public.audit_product_price_changes()
returns trigger
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_source text;
begin
  v_source := coalesce(
    nullif(current_setting('app.product_change_source', true), ''),
    'manual'
  );

  insert into public.product_price_history (
    product_id,
    product_code,
    changed_by,
    change_source,
    old_values,
    new_values
  )
  values (
    new.id,
    new.product_code,
    auth.uid(),
    case when v_source in ('manual', 'import', 'system') then v_source else 'manual' end,
    case when tg_op = 'UPDATE' then to_jsonb(old) else null end,
    to_jsonb(new)
  );

  return new;
end;
$$;

create or replace function public.touch_product_updated_at()
returns trigger
language plpgsql
set search_path = ''
as $$
begin
  new.updated_at := now();
  return new;
end;
$$;

drop trigger if exists products_touch_updated_at on public.products;
create trigger products_touch_updated_at
  before update on public.products
  for each row execute function public.touch_product_updated_at();

drop trigger if exists products_price_audit on public.products;
create trigger products_price_audit
  after insert or update on public.products
  for each row execute function public.audit_product_price_changes();
