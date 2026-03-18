-- ============================================================
-- Portfolio Tracker — Supabase schema
-- Run this in the Supabase SQL editor (Dashboard → SQL Editor)
-- ============================================================

-- ── profiles ──────────────────────────────────────────────────────────────────
create table if not exists public.profiles (
    user_id                         uuid references auth.users(id) on delete cascade primary key,
    email                           text not null,
    stripe_customer_id              text unique,
    subscription_status             text not null default 'inactive',
    -- possible values: inactive | active | trialing | past_due | canceled | unpaid
    subscription_plan               text,
    -- possible values: null | lite | full
    strategy_limit                  integer,
    -- null = unlimited (full plan); 20 = lite plan; 0 = no active subscription
    subscription_current_period_end timestamptz,
    created_at                      timestamptz not null default now(),
    updated_at                      timestamptz not null default now()
);

-- ── Row-Level Security ────────────────────────────────────────────────────────
alter table public.profiles enable row level security;

-- Users can read their own profile (used by the Streamlit anon-key client)
create policy "Users can view own profile"
    on public.profiles for select
    using (auth.uid() = user_id);

-- Users cannot write their own profile directly — all writes go through
-- the FastAPI backend which uses the service-role key (bypasses RLS)

-- ── Auto-create profile on signup ────────────────────────────────────────────
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
    insert into public.profiles (user_id, email)
    values (new.id, new.email)
    on conflict (user_id) do nothing;
    return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
    after insert on auth.users
    for each row execute procedure public.handle_new_user();

-- ── updated_at auto-bump ──────────────────────────────────────────────────────
create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
    new.updated_at = now();
    return new;
end;
$$;

drop trigger if exists set_profiles_updated_at on public.profiles;
create trigger set_profiles_updated_at
    before update on public.profiles
    for each row execute procedure public.set_updated_at();

-- ── Useful indexes ────────────────────────────────────────────────────────────
create index if not exists profiles_stripe_customer_id_idx
    on public.profiles (stripe_customer_id);

create index if not exists profiles_subscription_status_idx
    on public.profiles (subscription_status);
