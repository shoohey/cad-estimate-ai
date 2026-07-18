-- アプリデータ永続化用 key-value ストア（適用済み）
-- ポリシーは意図的に作成しない: anon/authenticated は全拒否、
-- アクセスは service_role キー（サーバーサイドのみ）に限定する。

create table if not exists public.app_store (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.app_store enable row level security;
