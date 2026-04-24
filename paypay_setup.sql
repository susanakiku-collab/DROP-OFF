-- PayPay単発1か月利用 用の追加SQL
alter table public.dropoff_teams
  add column if not exists manual_paid_until timestamptz null,
  add column if not exists payment_method text null;

create table if not exists public.dropoff_payments (
  id uuid primary key default gen_random_uuid(),
  team_id uuid not null references public.dropoff_teams(id) on delete cascade,
  provider text not null check (provider in ('stripe','paypay')),
  payment_type text not null check (payment_type in ('subscription','one_time_paid_month')),
  status text not null check (status in ('created','pending','completed','failed','cancelled','expired')),
  amount integer not null,
  currency text not null default 'JPY',
  provider_payment_id text null,
  merchant_payment_id text not null unique,
  valid_until timestamptz null,
  paid_at timestamptz null,
  created_by uuid null,
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
