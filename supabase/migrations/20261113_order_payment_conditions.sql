-- Vendedor 360° — condições de pagamento nos pedidos
-- Execute uma única vez no SQL Editor do Supabase.

alter table public.orders
  add column if not exists payment_term_days smallint not null default 7,
  add column if not exists payment_term_label text not null default '7 dias',
  add column if not exists finance_rate numeric(6,4) not null default 0,
  add column if not exists installment_days smallint[] not null default array[7]::smallint[];

do $$
begin
  if not exists (
    select 1 from pg_constraint
    where conname = 'orders_payment_term_days_check'
      and conrelid = 'public.orders'::regclass
  ) then
    alter table public.orders
      add constraint orders_payment_term_days_check
      check (payment_term_days in (0, 7, 14, 21, 28));
  end if;

  if not exists (
    select 1 from pg_constraint
    where conname = 'orders_finance_rate_check'
      and conrelid = 'public.orders'::regclass
  ) then
    alter table public.orders
      add constraint orders_finance_rate_check
      check (finance_rate in (0, 0.007, 0.014, 0.021));
  end if;
end;
$$;

comment on column public.orders.payment_term_days is
  'Prazo comercial selecionado: 0, 7, 14, 21 ou 28 dias.';
comment on column public.orders.payment_term_label is
  'Rótulo da condição comercial no momento da criação do pedido.';
comment on column public.orders.finance_rate is
  'Acréscimo financeiro aplicado ao pedido: 0%, 0,70%, 1,40% ou 2,10%.';
comment on column public.orders.installment_days is
  'Dias de vencimento das parcelas, contados da emissão, preservados como snapshot do pedido.';
