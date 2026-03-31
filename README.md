# PPTX Lecture Portal (monorepo)

Per `Cursor_AI_작업지시서.md`: Next.js + Prisma + Supabase + FastAPI parser.

## Layout

- `frontend/` — Next.js 14 (App Router), Tailwind, Prisma, Supabase Auth
- `parser-api/` — FastAPI, `python-pptx`, optional Storage upload

## Prerequisites

- Node 20+
- Python 3.11+
- Supabase project (`vector` extension enabled)
- Storage bucket `lecture-assets` (public read recommended for draft image URLs) if using parser uploads

## Environment

1. Copy `frontend/.env.example` to `frontend/.env.local` and fill values (see your `project_info.txt`; do not commit secrets).
2. DB password special characters must be **URL-encoded** inside `DATABASE_URL` (`@` -> `%40`, `!` -> `%21`, etc.).
3. Optional: `parser-api/.env` from `parser-api/.env.example` for Storage uploads.
4. Parser: use `PARSER_API_URL=http://127.0.0.1:8010` in `frontend/.env.local` (default port 8010 avoids clashes with other apps on 8000). Upload uses same-origin `/api/parse-pptx`, which proxies to that URL.

## Database

From `frontend/`:

```powershell
cd frontend; npm install; npx prisma migrate deploy
```

If `migrate deploy` fails on pooler, temporarily set `DATABASE_URL` to the **Direct** connection string from Supabase Connect, run migrate, then restore the pooler URL.

After first migrate, run SQL once (Supabase SQL Editor or `psql`):

- Contents of `frontend/prisma/sql/add_embedding_column.sql` (adds `vector(3072)` + HNSW index)

## Run locally

**Terminal 1 — parser**

```powershell
cd parser-api; python -m venv .venv; .\.venv\Scripts\Activate.ps1; pip install -r requirements.txt; python -m uvicorn app.main:app --reload --host 127.0.0.1 --port 8010
```

**Terminal 2 — frontend**

```powershell
cd frontend; npm install; npm run dev
```

Open http://localhost:3000 — sign in with Supabase (admin email must match `ADMIN_EMAILS`). `/admin/upload` sends `.pptx` to Next `/api/parse-pptx`, which forwards to FastAPI (draft only, no DB save yet).

## Auth

- `ADMIN_EMAILS` comma-separated → admin
- Or `app_metadata.role === "admin"` on the JWT

### Password reset (Supabase email link)

1. In Supabase: **Authentication → URL Configuration → Redirect URLs**, add:
   - `http://localhost:3000/auth/update-password`
   - (production domain when deployed)
2. **Site URL** should be `http://localhost:3000` for local dev.
3. Reset links open the site with `#...type=recovery`. The app redirects that hash to `/auth/update-password` and shows the new-password form.

## Next steps (지시서)

- Next.js route: LLM meta (title/description/tags) from `plainText`
- Publish: slides + `LectureChunk` + embeddings (`text-embedding-3-large`)
- Portal SSG + Global/Lecture RAG chat
