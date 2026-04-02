# PPTX Lecture Portal — 시작 가이드 (한국어)

PPTX 파일을 업로드하면 텍스트·도형·이미지를 추출해 HTML로 재현하고,
AI 어시스턴트(RAG)가 강의 내용에 답변하는 강의 포털입니다.

---

## 프로젝트 구조

```
PPTX_Parsing/
├── frontend/        # Next.js 14 (App Router) — UI, 인증, DB
├── parser-api/      # FastAPI + python-pptx — PPTX 파싱 서버
└── docs/            # 배포 가이드 등 문서
```

---

## 1. 사전 설치 항목

### 공통

| 소프트웨어 | 권장 버전 | 다운로드 |
|------------|-----------|----------|
| Node.js | 20 이상 | https://nodejs.org |
| Python | 3.11 이상 | https://python.org |

### 슬라이드 래스터(이미지 변환)용 — Windows 권장

| 소프트웨어 | 역할 | 비고 |
|------------|------|------|
| Microsoft PowerPoint | 고품질 슬라이드 이미지 생성 | 설치되어 있으면 자동 사용 |
| LibreOffice | PowerPoint 없을 때 대체 | https://libreoffice.org |
| Freesentation 폰트 | PPT 폰트 누락 방지 | 이미 설치되어 있으면 별도 설치 불필요. 없으면 → https://freesentation.blog/freesentation |

> PowerPoint 없이도 동작하지만, 래스터 품질은 LibreOffice 기준으로 낮아질 수 있습니다.

---

## 2. 필수 계정 준비

### Supabase (무료 플랜으로 시작 가능)

1. https://supabase.com 에서 프로젝트 생성
2. **Settings → Database → Extensions** 에서 `vector` 확장 활성화
3. **Storage → New bucket** 에서 `lecture-assets` 버킷 생성 (Public 읽기 허용)
4. 아래 값을 메모해 둡니다:
   - Project URL (예: `https://abc123.supabase.co`)
   - anon key
   - service_role key
   - Database 연결 문자열 (Connection String → **Transaction/Pooler** 방식)

### OpenAI (나중 단계 — 임베딩·RAG용)

- https://platform.openai.com 에서 API 키 발급
- 현재 드래프트(검수) 단계에서는 없어도 됩니다.

---

## 3. 환경 변수 설정

`.env.local` 파일은 git에 올라가지 않으므로 **`.env.example`을 복사해서 만드세요.**

```powershell
# Windows
copy frontend\.env.example frontend\.env.local
copy parser-api\.env.example parser-api\.env
```

```bash
# Mac / Linux
cp frontend/.env.example frontend/.env.local
cp parser-api/.env.example parser-api/.env
```

이후 각 파일을 열어 `<...>` 로 표시된 부분만 실제 값으로 채웁니다.

### frontend/.env.local 항목 안내

| 항목 | 찾는 곳 | 비고 |
|------|---------|------|
| `NEXT_PUBLIC_SUPABASE_URL` | Supabase → Settings → API | `https://xxx.supabase.co` |
| `NEXT_PUBLIC_SUPABASE_ANON_KEY` | 위와 동일 → anon public | `eyJ...` |
| `SUPABASE_SERVICE_ROLE_KEY` | 위와 동일 → service_role | 서버 전용, 절대 노출 금지 |
| `DATABASE_URL` | Supabase → Settings → Database → **Transaction(Pooler)** | 비밀번호 특수문자 인코딩 필요 |
| `ADMIN_EMAILS` | 직접 입력 | 쉼표로 여러 개 가능 |
| `OPENAI_API_KEY` | platform.openai.com/api-keys | RAG 기능 사용 시만 필요 |
| `PARSER_API_URL` | 기본값 유지 | `http://127.0.0.1:8010` |

> **DATABASE_URL 비밀번호 특수문자 인코딩**  
> `@` → `%40` · `!` → `%21` · `#` → `%23` · `$` → `%24`  
> 예: 비밀번호가 `abc!@#` → URL에서 `abc%21%40%23`

---

## 4. 데이터베이스 초기화

```powershell
cd frontend
npm install
npx prisma migrate deploy
```

마이그레이션 후 Supabase SQL Editor에서 아래 파일 내용을 한 번 실행합니다:

```
frontend/prisma/sql/add_embedding_column.sql
```

(pgvector `vector(3072)` 컬럼과 HNSW 인덱스를 추가합니다.)

---

## 5. 실행

### 터미널 1 — parser-api (PPTX 파싱 서버)

```powershell
cd parser-api
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m uvicorn app.main:app --reload --host 127.0.0.1 --port 8010
```

### 터미널 2 — frontend (Next.js)

```powershell
cd frontend
npm install
npm run dev
```

브라우저에서 http://localhost:3000 접속

---

## 6. 처음 사용 방법

1. http://localhost:3000 에서 **이메일 로그인** (`.env.local`의 `ADMIN_EMAILS`에 등록된 이메일 필요)
2. 로그인 후 http://localhost:3000/admin/upload 이동
3. `.pptx` 파일을 선택하면 자동으로 파싱 시작
4. 파싱 완료 후 **HTML 렌더링 미리보기**와 **원본 슬라이드 이미지** 비교 검수
5. 슬라이드별 설명 입력 → Description 자동 합본 확인
6. (추후) 최종 등록 시 DB에 저장 + 임베딩 생성

---

## 7. 인증 관련

### 관리자 권한 부여 방법

- `.env.local`의 `ADMIN_EMAILS`에 이메일 추가 (가장 간단)
- 또는 Supabase Dashboard → Authentication → 해당 유저 → `app_metadata`에 `{ "role": "admin" }` 추가

### 비밀번호 재설정 이메일 설정

Supabase Dashboard → **Authentication → URL Configuration** 에서:

- **Site URL**: `http://localhost:3000`
- **Redirect URLs**에 추가: `http://localhost:3000/auth/update-password`

---

## 8. 문제 해결

| 증상 | 확인 사항 |
|------|-----------|
| 파싱 후 슬라이드 이미지가 검거나 텍스트 누락 | LibreOffice 설치 여부, 폰트 설치 여부 확인 |
| parser-api 연결 안 됨 | `.env.local`의 `PARSER_API_URL` 확인, parser 서버 실행 중인지 확인 |
| DB 마이그레이션 실패 | `DATABASE_URL`을 Supabase의 **Direct(Transaction 아닌)** 연결로 변경 후 재시도 |
| 로그인 후 /admin 접근 안 됨 | `ADMIN_EMAILS`에 해당 이메일 등록 여부 확인 |
| 래스터가 너무 어둡거나 흐림 | Freesentation 폰트 설치, LibreOffice 최신 버전 사용 |

---

## 9. 상세 배포 가이드

서버 배포(Fly.io, Railway, Vercel 등)가 필요하면 `docs/DEPLOYMENT.md`를 참고하세요.

---

## 향후 기능 (개발 중)

- 강의 포털 메인 페이지 (SSG)
- 슬라이드 기반 AI 질의응답 (RAG 채팅)
- 최종 등록 → DB + 임베딩 영구 저장
- pgvector 기반 유사 슬라이드 검색
