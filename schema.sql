-- BOTEC App: Supabase Schema
-- Run this in your Supabase SQL editor

-- Documents table
create table if not exists botec_documents (
  id          uuid default gen_random_uuid() primary key,
  user_id     uuid references auth.users(id) on delete cascade not null,
  name        text not null,
  programme   text,
  data        jsonb not null default '{}',
  created_at  timestamptz default now(),
  updated_at  timestamptz default now()
);

-- Auto-update updated_at
create or replace function update_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

create trigger botec_documents_updated_at
  before update on botec_documents
  for each row execute function update_updated_at();

-- Row-level security
alter table botec_documents enable row level security;

create policy "Users can view own documents"
  on botec_documents for select
  using (auth.uid() = user_id);

create policy "Users can create own documents"
  on botec_documents for insert
  with check (auth.uid() = user_id);

create policy "Users can update own documents"
  on botec_documents for update
  using (auth.uid() = user_id);

create policy "Users can delete own documents"
  on botec_documents for delete
  using (auth.uid() = user_id);
