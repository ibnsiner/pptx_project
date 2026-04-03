# PPTX → 웹페이지 퍼블리시 파이프라인

> 작성일: 2026-04-01  
> 목표: PPTX 파일을 파싱해 Supabase에 구조화 데이터를 저장하고,  
> Vercel(Next.js)이 이를 읽어 웹페이지로 서빙한다.  
> 텍스트 선택 시 GPT-4o Vision이 멀티모달 맥락으로 보충 설명을 제공한다.

---

## 1. 전체 아키텍처

```
[로컬 / 서버]
PPTX 파일 업로드
    ↓
parser-api (FastAPI + python-pptx + PowerPoint COM)
    ├── 래스터 JPEG 생성 (슬라이드별)
    ├── JSON elements 추출 (좌표·스타일·텍스트·도형·표)
    ├── 슬라이드 배경색 추출
    └── 전체 텍스트 추출

    ↓ 최종 등록(Publish) 시

[Supabase]
    ├── Storage: 이미지 파일 저장
    └── DB: 구조화 데이터 저장

    ↓

[Vercel / Next.js]
    ├── 공개 페이지: /lectures/[id]
    │     DB + Storage URL 읽어서 렌더링
    └── AI 보충 설명: /api/chat
          텍스트 + 이미지 URL + 맥락 → GPT-4o Vision → 응답
```

---

## 2. Supabase 저장 구조

### 2-1. Storage (파일)

```
lecture-assets/
└── {lectureId}/
    ├── thumbnail.jpg              ← 첫 슬라이드 (강의 대표 이미지)
    ├── slides/
    │   ├── 001.jpg                ← 슬라이드 1 래스터
    │   ├── 002.jpg
    │   └── ...
    └── images/
        ├── s01_img01.jpg          ← 슬라이드 내 추출 이미지
        └── ...
```

- **버킷**: `lecture-assets` (Public 읽기 허용)
- **URL 형식**: `https://{project}.supabase.co/storage/v1/object/public/lecture-assets/{path}`

---

### 2-2. DB 스키마 (Prisma)

#### Lecture 테이블

| 필드 | 타입 | 설명 |
|------|------|------|
| id | String (cuid) | PK |
| title | String | 강의 제목 |
| description | String? | 슬라이드 설명 합본 |
| tags | String[] | 태그 목록 |
| author | String | 강사명 |
| slideCount | Int | 전체 슬라이드 수 |
| thumbnailUrl | String? | 대표 이미지 URL (Storage) |
| slideWidthEmu | Int? | 슬라이드 가로 (EMU) |
| slideHeightEmu | Int? | 슬라이드 세로 (EMU) |
| createdAt | DateTime | 등록일 |

#### Slide 테이블

| 필드 | 타입 | 설명 |
|------|------|------|
| id | String (cuid) | PK |
| lectureId | String | FK → Lecture |
| slideNumber | Int | 슬라이드 번호 (1-based) |
| rasterUrl | String? | 래스터 JPEG URL (Storage) |
| elements | Json | 파싱된 요소 배열 (텍스트·도형·표·이미지) |
| plainText | String? | 전체 텍스트 (검색용) |
| slideNote | String? | 관리자 입력 슬라이드 설명 |
| backgroundColor | String? | 슬라이드 배경색 (#rrggbb) |

#### LectureChunk 테이블 (RAG용)

| 필드 | 타입 | 설명 |
|------|------|------|
| id | String (cuid) | PK |
| lectureId | String | FK → Lecture |
| slideId | String? | FK → Slide |
| slideNumber | Int | 슬라이드 번호 |
| content | String | 텍스트 청크 |
| tokenCount | Int? | 토큰 수 |
| embedding | vector(3072) | text-embedding-3-large 벡터 |

---

## 3. 퍼블리시 파이프라인 (최종 등록 클릭 시)

```
관리자: 검수 완료 → "최종 등록" 버튼 클릭

Step 1: 래스터 이미지 → Supabase Storage 업로드
  - 슬라이드별 JPEG → slides/00N.jpg
  - 첫 슬라이드 → thumbnail.jpg
  - 슬라이드 내 이미지 → images/sNN_imgNN.jpg
  → 각 URL 획득

Step 2: Lecture 레코드 생성
  - 제목, 태그, 설명, 작성자, 슬라이드 수, thumbnailUrl 저장

Step 3: Slide 레코드 생성 (슬라이드별)
  - slideNumber, rasterUrl, elements(JSON), plainText, slideNote, backgroundColor

Step 4: LectureChunk 생성 + 임베딩
  - 슬라이드별 plainText → text-embedding-3-large API 호출
  - 벡터 저장 (pgvector)

Step 5: 공개 URL 반환
  - /lectures/{lectureId}
```

---

## 4. Vercel 공개 페이지 렌더링

### 4-1. /lectures/[id] (강의 목록 → 슬라이드 뷰어)

```
Next.js SSR/ISR:
  1. DB에서 Lecture + Slides 읽기
  2. 슬라이드 렌더링:
     - 배경색 적용 (backgroundColor)
     - JSON elements로 HTML 요소 배치 (SlideViewer 컴포넌트)
     - rasterUrl은 시각 보완용 (배경 또는 이미지 fallback)
  3. 슬라이드 네비게이션 (← N/전체 →)
```

### 4-2. AI 보충 설명 (텍스트 선택 시)

```
사용자가 슬라이드 텍스트 드래그 선택
    ↓
POST /api/chat
  {
    lectureId: "...",
    slideNumber: N,
    selectedText: "사용자가 선택한 텍스트",
    rasterUrl: "https://supabase.co/.../slides/00N.jpg",  ← 시각 맥락
    slideElements: [...],   ← 구조 맥락
    slideNote: "...",       ← 관리자 큐레이션 맥락
    plainText: "...",       ← 전체 텍스트 맥락
  }
    ↓
GPT-4o Vision (멀티모달):
  - 이미지 URL로 슬라이드 시각 이해
  - 선택 텍스트 + 맥락으로 보충 설명 생성
    ↓
사이드 패널 / 모달에 응답 표시
```

---

## 5. Supabase ↔ Vercel 연결

### 환경 변수 (Vercel 대시보드에 등록)

```env
# Supabase
NEXT_PUBLIC_SUPABASE_URL=https://{project}.supabase.co
NEXT_PUBLIC_SUPABASE_ANON_KEY=eyJ...
SUPABASE_SERVICE_ROLE_KEY=eyJ...
DATABASE_URL=postgresql://...pooler.supabase.com:6543/postgres?pgbouncer=true

# OpenAI
OPENAI_API_KEY=sk-...

# Admin
ADMIN_EMAILS=you@example.com
```

### Supabase Storage 버킷 설정

```sql
-- lecture-assets 버킷 공개 읽기 허용
INSERT INTO storage.policies (name, bucket_id, operation, definition)
VALUES (
  'Public read for lecture-assets',
  'lecture-assets',
  'SELECT',
  'true'
);
```

---

## 6. 구현 단계별 체크리스트

### Phase 1 — 저장 인프라
- [ ] Supabase Storage `lecture-assets` 버킷 Public 설정
- [ ] parser-api에 Supabase 연결 (`SUPABASE_URL`, `SUPABASE_SERVICE_ROLE_KEY`)
- [ ] 래스터 JPEG → Storage 업로드 함수 구현
- [ ] Prisma 스키마에 `rasterUrl`, `backgroundColor`, `slideNote` 추가
- [ ] Prisma 마이그레이션 실행

### Phase 2 — 퍼블리시 API
- [ ] `POST /api/lectures` — 최종 등록 엔드포인트
  - 래스터 Storage 업로드
  - Lecture + Slide 레코드 생성
  - LectureChunk + 임베딩 생성
- [ ] 관리자 페이지 "최종 등록" 버튼 활성화

### Phase 3 — 공개 페이지
- [ ] `/lectures` — 강의 목록 (카드 그리드, SSG)
- [ ] `/lectures/[id]` — 슬라이드 뷰어
  - 배경색 적용된 SlideViewer
  - 슬라이드 네비게이션
  - rasterUrl 이미지 보완 렌더링

### Phase 4 — AI 보충 설명
- [ ] 텍스트 선택 감지 (onMouseUp / onTouchEnd)
- [ ] "AI에게 물어보기" 버튼 팝업
- [ ] `POST /api/chat` — GPT-4o Vision 연동
- [ ] 응답 사이드 패널 UI

### Phase 5 — 배포
- [ ] Vercel 환경 변수 등록
- [ ] `frontend/` Vercel 배포
- [ ] parser-api Docker 이미지 → Fly.io / Railway 배포
- [ ] 도메인 설정

---

## 7. 기술 스택 요약

| 구성 요소 | 기술 | 역할 |
|----------|------|------|
| PPTX 파싱 | FastAPI + python-pptx + PowerPoint COM | 데이터 추출 |
| 이미지 저장 | Supabase Storage | 래스터·이미지 CDN |
| 데이터 저장 | Supabase PostgreSQL + Prisma | 구조화 데이터 |
| 벡터 검색 | pgvector (text-embedding-3-large) | RAG |
| 웹 서빙 | Next.js 14 App Router (Vercel) | 공개 페이지 |
| AI 보충 설명 | GPT-4o Vision + OpenAI API | 멀티모달 Q&A |
| 인증 | Supabase Auth | 관리자 접근 제어 |
