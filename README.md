# Gestão Imobiliária

App de gestão de carteira imobiliária (Next.js 14 + Supabase), com login, dashboard de
rentabilidade, imóveis, alojamento local, movimentos com importação de Excel, e definições.

## 1. Base de dados (Supabase)

O projeto Supabase já está criado e as chaves em `.env.local.example` já apontam para ele
(o mesmo projeto que usas noutras apps, com tabelas prefixadas `imob_`).

Falta garantir que as tabelas existem. Se a Claude não confirmar que aplicou a migração:

1. Abre https://supabase.com/dashboard/project/ldlgxnalskrehfwmdqqv/sql/new
2. Cola o conteúdo de `supabase/migrations/0001_init.sql`
3. Corre (Run)

Isto cria as tabelas `imob_properties`, `imob_transactions`, `imob_reservations`,
`imob_settings`, todas com Row Level Security: cada utilizador só vê as suas próprias linhas.

## 2. Local (opcional, para testar antes de publicar)

```bash
cp .env.local.example .env.local
npm install
npm run dev
```

Abre http://localhost:3000 — vai redirecionar para `/login`. Usa "Primeira vez? Criar conta"
para criares a tua conta (o Supabase pode pedir para confirmares por email, dependendo das
definições de Auth do projeto).

## 3. GitHub

```bash
git init
git add .
git commit -m "Gestão imobiliária"
gh repo create gestao-imobiliaria --private --source=. --push
```

(ou cria o repo manualmente em github.com e faz `git remote add origin ...` + `git push`)

## 4. Vercel

Opção A — dashboard: em vercel.com, "Add New Project" → importa o repo → antes de fazer
deploy, adiciona as environment variables (Settings → Environment Variables):

- `NEXT_PUBLIC_SUPABASE_URL`
- `NEXT_PUBLIC_SUPABASE_ANON_KEY`

(os valores estão em `.env.local.example`) → Deploy.

Opção B — CLI:

```bash
npm i -g vercel
vercel link
vercel env add NEXT_PUBLIC_SUPABASE_URL
vercel env add NEXT_PUBLIC_SUPABASE_ANON_KEY
vercel --prod
```

## Estrutura

- `app/login` — login e criação de conta (Supabase Auth)
- `app/(app)` — área protegida (redireciona para `/login` sem sessão)
- `components/ImobiliarioApp.tsx` — toda a aplicação (dashboard, imóveis, alojamento local,
  movimentos, definições)
- `middleware.ts` — protege as rotas e mantém a sessão viva
- `supabase/migrations/0001_init.sql` — esquema da base de dados

## Notas

- Os dados são financeiros e cada conta só vê os seus próprios registos (RLS). Não partilhes
  a tua password.
- A importação de Excel (Movimentos → Importar Excel) aceita `.xlsx`, `.xls` e `.csv` com
  mapeamento de colunas.
