# [Cursor AI 작업 지시서] PPTX to Web Portal with AI Assistant

## 1. 프로젝트 개요 및 비전 (Project Overview)
본 프로젝트는 사내/교내 환경에서 2명의 관리자(업로더)와 약 500명의 뷰어가 사용하는 **"넷플릭스/노션 스타일의 강의안(PPTX) 포털 웹 서비스"**를 구축하는 것이다.

* **핵심 가치 1:** PPT 슬라이드를 통이미지로 변환하지 않는다. 텍스트, 도형, 이미지를 개별 추출하여 100% HTML/CSS 기반의 반응형 웹 컴포넌트로 렌더링한다.
* **핵심 가치 2:** 추출된 텍스트 데이터를 바탕으로 환각(Hallucination) 없는 RAG 기반 AI 어시스턴트를 제공한다.
* **핵심 가치 3:** 500명의 뷰어가 접속해도 서버 부하가 없도록 Next.js의 SSG(Static Site Generation)를 적극 활용하여 로딩 속도를 극대화하고 비용을 최소화한다.
* **핵심 가치 4:** 관리자가 최종 업로드 전, 파싱된 HTML 결과물과 AI가 생성한 메타데이터를 확인하고 수정할 수 있는 **"프리뷰 및 검수(Preview & Verify)"** 단계를 반드시 거쳐야 한다.

## 2. 기술 스택 (Tech Stack)
* **Frontend & Main Backend:** Next.js (App Router), TypeScript, Tailwind CSS
* **Parsing Microservice:** Python (FastAPI), `python-pptx`
* **Database & Vector DB:** Supabase (PostgreSQL + `pgvector`)
* **ORM:** Prisma
* **Storage:** Supabase Storage
* **AI / LLM:** OpenAI API (gpt-4o) 혹은 Anthropic (Claude 3.5), LangChain
* **임베딩(확정):** OpenAI **`text-embedding-3-large`** (벡터 차원 **3072**, DB `vector(3072)`와 일치)
* **Deployment:** Vercel (Next.js), Render/Railway (Python FastAPI)
* **인증:** Supabase Auth (역할은 §3 참고)

## 3. 인증·권한 (Admin / Viewer)

* **전제:** 인증은 **Supabase Auth**를 쓴다. 사내 SAML/OIDC가 있으면 Supabase SSO·IdP 연동으로 동일하게 `role`만 맞추면 된다.
* **역할은 두 가지:** `admin`, `viewer` (다른 역할은 두지 않는다).
  * **admin (2명):** `/admin/**` 접근, PPTX 업로드·파서(FastAPI) 호출·프리뷰/검수·**최종 등록**(임베딩·DB·Storage 확정). 필요 시 Storage에 원본/에셋 업로드. 포털·뷰어 URL은 viewer와 동일하게 이용 가능.
  * **viewer (약 500명):** 공개 **강의 목록·강의 뷰어**, Global/Lecture **AI 채팅(질의)**. **`/admin/**` 및 관리자 전용 API(파서 프록시·최종 등록 등)는 403.**
* **역할 저장 위치:** Supabase `app_metadata.role` = `"admin"` | `"viewer"` 권장. 관리자 계정만 초대 시 admin, 나머지 기본 viewer.
* **간단 대안:** 배포 환경의 `ADMIN_EMAILS`(쉼표 구분) 또는 테이블 `AdminAllowlist(email)`로 로그인 이메일이 목록에 있으면 admin으로 간주.
* **Next.js:** Middleware 또는 Server Component에서 세션 확인 후 **`/admin` 트리는 admin만** 통과.
* **API:** 관리자 전용 Route Handler는 **세션 + role 검증** 후 동작. **Supabase `service_role` 키는 서버에서만** 사용하고 클라이언트 번들에 넣지 않는다.
* **RLS:** viewer는 공개 강의에 매핑된 `Lecture` / `Slide` / `LectureChunk` 등 **SELECT만**. 강의 생성·수정·삭제·청크 쓰기는 서버(admin 경로)가 `service_role` 또는 별도 admin 정책으로 수행.
* **Draft 단계:** 파싱 직후 JSON은 DB에 없으므로 RLS 대상이 아니다(브라우저 상태만).

## 4. 시스템 아키텍처 및 데이터 흐름 (System Architecture)
Vercel의 Serverless 함수 타임아웃(10~15초)을 방지하기 위해 무거운 PPT 파싱 로직은 별도의 Python 서버로 분리하며, Draft -> Preview -> Publish 흐름을 따른다.

* **[Admin 업로드 및 임시 파싱]:** 관리자 2명이 Next.js 웹(`/admin/upload`)에서 `.pptx` 파일을 업로드하면, Python FastAPI 서버로 파일을 전송한다. Python 서버는 파싱된 JSON 배열과 AI가 추출한 메타데이터(title, description, tags)를 DB에 저장하지 않고 Next.js 프론트엔드로 즉시 반환(Return)한다.
* **[Python 파싱 상세]:** `python-pptx`가 슬라이드 크기(EMU)를 추출하고, 모든 요소의 절대 좌표(x, y, w, h)를 비율(%)로 변환한다.
* **[리소스 분리]:** 이미지/도형 객체는 이미지로 추출하여 Supabase Storage에 업로드하고 URL을 반환받는다.
* **[AI 자동화]:** 추출된 전체 텍스트를 LLM에 보내 강의의 title, description(3줄 요약), tags(키워드 5개)를 자동 생성한다. RAG 임베딩용 청크는 **슬라이드 단위(1슬라이드 = 1청크, 1순위 전략)**로 만든다(§5.3).
* **[프리뷰 렌더링 & 검수]:** 프론트엔드는 응답받은 JSON 데이터를 상태(State)에 담아 `SlideViewer` 컴포넌트로 HTML/CSS 변환 결과를 뷰어 화면과 100% 동일하게 미리 보여준다. 화면 측면에는 AI가 자동 생성한 메타데이터를 수정할 수 있는 폼(Form)을 제공한다.
* **[최종 등록 및 DB 저장]:** 관리자가 이상 없음을 확인하고 "최종 등록"을 누르면, 프론트엔드가 파싱된 JSON 결과물(화면 렌더링용)과 AI 벡터 데이터(`pgvector`)를 Supabase에 영구 저장(`INSERT`)한다.
* **[Viewer 렌더링]:** 500명의 뷰어가 접속하면 Next.js가 DB의 JSON을 가져와 `position: absolute` 기반의 HTML/CSS로 원본과 동일하게 렌더링한다.

### 4.1 AI·임베딩 호출 위치 (제안)

| 작업 | 권장 실행 위치 | 이유 |
|------|----------------|------|
| PPTX 바이너리 파싱, EMU→%, 이미지 바이트 추출 | **parser-api (FastAPI)** | Vercel 타임아웃 회피, CPU·파일 I/O에 적합 |
| 메타 생성 (title, description, tags) | **Next.js Route Handler** (파서가 돌려준 `plainText` 또는 슬라이드 텍스트를 body로 POST) | API 키·프롬프트 버전·감사 로그를 한곳에서 관리; 파서는 LLM 키를 몰라도 됨 |
| (대안) 메타를 파서와 한 번에 | parser-api에서 LLM 호출 | 왕복 감소. Render/Railway에 API 키를 두고 타임아웃을 넉넉히 잡을 때만 권장 |
| 청킹 + **임베딩** + `LectureChunk` INSERT | **Next.js** “최종 등록” API | `Lecture`/`Slide` 생성 직후 같은 트랜잭션·id로 청크를 묶기 쉬움; Prisma + raw SQL 혼용 |
| RAG 검색 + 채팅 응답(스트리밍) | **Next.js Route Handler** | `lecture_id` 필터·rate limit·viewer 세션과 결합 |

요약: **파서는 파일→JSON(+선택 스토리지 업로드)** 까지, **모든 LLM·임베딩·챗은 Next 서버**가 담당하는 구성이 운영·보안상 단순하다.

## 5. 데이터베이스 스키마 (Prisma + RAG 벡터)

다음 구조를 바탕으로 `schema.prisma`를 구성한다. **`vector` 컬럼**은 Prisma 네이티브 타입이 약하므로 Supabase에서 확장 설치 후 **마이그레이션 SQL로 컬럼 생성**하거나, Prisma `Unsupported` / **raw query**로 읽기·쓰기한다.

### 5.1 관계형 모델 (Prisma 예시)

```prisma
// 1. 강의 단위 (포털 노출용)
model Lecture {
  id          String   @id @default(cuid())
  title       String
  description String?  // LLM이 자동 생성한 3줄 요약 (검수 후 확정)
  thumbnail   String?  // 첫 슬라이드 썸네일 URL
  author      String   // 업로더 이름
  tags        String[] // LLM이 자동 생성한 태그 배열
  createdAt   DateTime @default(now())
  slides      Slide[]
  chunks      LectureChunk[]
}

// 2. 개별 슬라이드 데이터 (웹 렌더링용)
model Slide {
  id          String   @id @default(cuid())
  slideNumber Int
  lectureId   String
  lecture     Lecture  @relation(fields: [lectureId], references: [id], onDelete: Cascade)
  elements    Json     // 파싱된 요소들 (text, coordinates, font, image URL 등)
  chunks      LectureChunk[]
}

// 3. RAG용 텍스트 청크 (임베딩은 SQL에서 vector 타입으로 보관)
model LectureChunk {
  id          String   @id @default(cuid())
  lectureId   String
  lecture     Lecture  @relation(fields: [lectureId], references: [id], onDelete: Cascade)
  slideId     String?
  slide       Slide?   @relation(fields: [slideId], references: [id], onDelete: SetNull)
  slideNumber Int      // 출처 링크용 비정규화 (뷰어 /lectures/[id]?slide=n)
  chunkIndex  Int      // 강의 내 순서; 슬라이드당 1청크 전략이면 slideNumber와 동일하게 두거나 0부터 연번
  content     String   // 해당 슬라이드에서 추출한 텍스트를 순서대로 이은 본문(검색·스니펫용)
  // embedding: Supabase에서 "vector(3072)" — OpenAI text-embedding-3-large — Prisma Unsupported 또는 raw INSERT/SELECT
  tokenCount  Int?     // 선택: 과금·디버깅
  createdAt   DateTime @default(now())

  @@unique([lectureId, slideNumber]) // 슬라이드당 1청크 전략
  @@index([lectureId])
}
```

### 5.2 Supabase / pgvector (SQL 예시)

Prisma 마이그레이션에 포함하거나 SQL 에디터에서 실행한다. **임베딩 모델: OpenAI `text-embedding-3-large` → 벡터 차원 3072.**

```sql
-- 확장 (한 번만)
create extension if not exists vector;

-- LectureChunk에 벡터 컬럼 추가 (Prisma 모델과 동기화)
alter table "LectureChunk" add column if not exists embedding vector(3072);

-- 유사도 검색 인덱스 (데이터 쌓인 뒤 CONCURRENTLY 권장)
create index if not exists lecturechunk_embedding_hnsw
  on "LectureChunk"
  using hnsw (embedding vector_cosine_ops);
```

* **검색 쿼리:** `order by embedding <=> $query_vector` (cosine distance) + **Lecture Mode** 시 `where "lectureId" = $id`. Global Mode는 `lectureId` 조건 없음.
* **출처 UI:** `lectureId`, `slideNumber`, (선택) `slideId`로 강의 상세·슬라이드 앵커 URL 생성.

### 5.3 청크 전략 (확정)

* **1순위(기본): 슬라이드 경계 청크** — 각 슬라이드에서 추출한 텍스트만 순서대로 이어 **슬라이드당 `LectureChunk` 1행**을 만든다.
* `slideId`·`slideNumber`·`content`가 출처 링크와 1:1로 대응되어 Global/Lecture RAG UI에 유리하다.
* 동일 슬라이드에 행이 여러 개 생기지 않도록 한다(분할·오버랩 전략은 사용하지 않음). 추후 토큰 한도 등으로 분할이 필요해지면 그때 별도 옵션으로 확장한다.

## 6. 핵심 구현 로직 (Crucial Implementation Rules)
A. Python 좌표 퍼센트(%) 변환 공식
PPTX의 EMU 단위를 웹 반응형에 맞게 퍼센트로 변환하여 JSON으로 출력해야 한다.

left_percent = (shape.left / slide_width) * 100

top_percent = (shape.top / slide_height) * 100

width_percent = (shape.width / slide_width) * 100

height_percent = (shape.height / slide_height) * 100

B. 프론트엔드 UI 렌더링 (airppt-parser 벤치마킹 주의사항)
[가장 중요한 제약사항]: GitHub의 airppt-parser 리포지토리를 벤치마킹하되, 절대 npm install로 직접 설치하지 마라 (8년 전 레거시 코드임). 해당 라이브러리가 JSON 데이터를 HTML/CSS로 매핑하는 핵심 알고리즘만 분석하여, 최신 React/Next.js 컴포넌트(SlideViewer.tsx)로 **직접 포팅(재구현)**해야 한다.

텍스트는 드래그 가능한 HTML 요소여야 한다.

폰트 크기는 부모 컨테이너 대비 비율(vw/vh) 또는 컨테이너 쿼리를 사용한다.

슬라이드 컨테이너는 aspect-ratio: 16/9; position: relative;를 기본으로 하고, 내부 요소는 position: absolute; left: {x}%; top: {y}%; 형태로 배치한다.

C. RAG 기반 AI Assistant (Metadata Filtering)
두 가지 모드의 텍스트 기반 RAG를 구현한다.

Global Mode (전체 검색): 메인 포털 화면의 챗봇. lecture_id 필터 없이 전체 Vector DB를 검색하며, 답변 시 "어느 강의, 몇 페이지"인지 이동 가능한 출처(Source) 링크를 제공해야 한다.

Lecture Mode (특정 강의 딥다이브): 개별 강의 뷰어 화면의 챗봇. 현재 보고 있는 강의의 lecture_id를 메타데이터 필터로 적용하여 해당 강의 내용에만 집중된 답변을 제공한다.

## 7. 참조용 프로토타입 코드 (Reference Only)
아래 코드는 아키텍처(Python의 비율 변환 및 React의 절대 좌표 렌더링) 이해를 돕기 위한 예시 뼈대 코드이므로, 이를 참고하여 프로젝트에 맞게 고도화할 것.

### [참조 1: Python FastAPI 백엔드 파서 핵심 로직]

```python
from fastapi import FastAPI, UploadFile, File
from pptx import Presentation
import io

app = FastAPI()

@app.post("/api/parse-pptx")
async def parse_pptx(file: UploadFile = File(...)):
    # ... 파일 로드 ...
    content = await file.read()
    prs = Presentation(io.BytesIO(content))

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    slides_data = []
    for i, slide in enumerate(prs.slides):
        elements = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            # 절대 좌표를 %로 변환
            left_pct = (shape.left / slide_width) * 100
            top_pct = (shape.top / slide_height) * 100
            width_pct = (shape.width / slide_width) * 100
            height_pct = (shape.height / slide_height) * 100

            elements.append({
                "type": "text",
                "content": shape.text,
                "style": {
                    "left": f"{left_pct}%", "top": f"{top_pct}%",
                    "width": f"{width_pct}%", "height": f"{height_pct}%",
                }
            })
        slides_data.append({"slideNumber": i + 1, "elements": elements})
    return {"slides": slides_data, "meta": {"title": "AI 자동생성 제목", "tags": []}}
```

### [참조 2: Next.js SlideViewer 컴포넌트 렌더링 로직]

```tsx
export default function SlideViewer({ elements }: { elements: SlideElement[] }) {
  return (
    <div className="relative w-full max-w-5xl mx-auto bg-white shadow-md overflow-hidden" style={{ aspectRatio: '16/9' }}>
      {elements.map((el, index) => (
        <div
          key={index}
          style={{
            position: 'absolute',
            left: el.style.left, top: el.style.top, width: el.style.width, height: el.style.height,
            fontSize: el.style.fontSize || '1.5vw', 
            display: 'flex', flexDirection: 'column',
          }}
        >
          {el.type === 'text' && el.content && <span className="whitespace-pre-wrap">{el.content}</span>}
        </div>
      ))}
    </div>
  );
}
```

## 8. 단계별 개발 계획 (Step-by-Step Execution Plan)
Cursor는 다음 순서대로 작업을 진행하고, 각 단계가 끝날 때마다 사용자에게 확인을 받아야 한다. 오버엔지니어링을 피하고 직관적이고 가벼운 코드를 작성해라.

Step 1: 프로젝트 폴더 구조 세팅 (Monorepo 형태: /frontend Next.js 앱과 /parser-api Python FastAPI 앱 분리).

Step 2: Supabase 프로젝트 연결 및 Prisma 스키마 초기화 & 마이그레이션(Auth, RLS, `Lecture`/`Slide`/`LectureChunk`, pgvector 컬럼·인덱스).

Step 3: Python FastAPI 셋업 및 python-pptx를 활용한 파싱 로직 + 퍼센트(%) 계산 + 이미지 스토리지 업로드 + JSON 반환 API 작성 (데이터베이스에 직접 저장하지 않음).

Step 4: 프론트엔드 - 관리자용 업로드 페이지(/admin/upload) 작성. 파일 업로드 폼 및 Python API 호출 로직 연동.

Step 5: 프론트엔드 - 파싱된 JSON을 HTML/CSS로 그려내는 SlideViewer 컴포넌트 개발 (airppt-parser 로직 포팅) 및 업로드 페이지 내 '프리뷰(미리보기) 대시보드' 조립. AI 생성 메타데이터 수정 Form 함께 구현.

Step 6: 프론트엔드/백엔드 - 프리뷰 화면 폼에서 "최종 등록" Submit 시, 추출된 텍스트를 AI 임베딩 처리하고 최종 검증된 데이터를 Supabase DB 및 pgvector에 영구 저장하는 로직 작성.

Step 7: Next.js 포털 메인 UI(강의 리스트 카드, SSG 적용) 및 두 가지 모드(Global/Lecture)의 AI 채팅창 구현.

시작할 준비가 되면 "네, 지시서를 완벽히 이해했습니다. 프로젝트 셋업(Step 1)부터 진행하겠습니다."라고 대답해 줘.

일부 빠진 내용이 있어서 위와 같이 원문을 수정했어.