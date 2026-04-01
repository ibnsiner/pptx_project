# 배포 가이드 (PPTX Portal)

> 이 문서는 `parser-api`(FastAPI) 서버를 **로컬 Windows 개발 환경** 또는
> **Linux 프로덕션 서버(컨테이너 포함)** 에 배포할 때 필요한 사항을 정리합니다.
> Next.js 프론트는 Vercel에 배포하되, **parser-api는 반드시 별도 서버**에 두어야 합니다.

---

## 1. 아키텍처 개요

```
[브라우저]
    |
    v
[Vercel] Next.js App Router
    |  - 인증, 메타 CRUD, Prisma/Supabase 접근
    |  - /api/parse-pptx → parser-api 프록시
    |
    v
[별도 서버] parser-api  (FastAPI + python-pptx + LibreOffice + PyMuPDF)
    |  - PPTX → JSON 파싱
    |  - PPTX → PDF(LibreOffice) → JPEG 래스터 생성
    |
    v
[Supabase] PostgreSQL + Storage
```

**Vercel에 LibreOffice를 올릴 수 없습니다.**
서버리스 함수 크기·타임아웃(최대 300초)·실행 환경 모두 제약이 있습니다.

---

## 2. parser-api 서버 요구 사항

### 2-1. 런타임

| 항목 | 요구 사항 |
|------|-----------|
| Python | 3.11 이상 |
| pip 패키지 | `parser-api/requirements.txt` 참조 |
| LibreOffice | **7.x 이상** (headless 변환용, 필수) |
| 메모리 | 최소 1 GB (PPTX 변환 시 피크 500 MB 이상 가능) |
| 디스크 임시 공간 | 업로드 1건당 최대 ~200 MB (변환 후 자동 삭제) |
| 타임아웃 | 단일 요청 최대 **280초** 허용 필요 |

### 2-2. Python 패키지

```
fastapi>=0.115.0
uvicorn[standard]>=0.32.0
python-pptx>=1.0.2
python-multipart>=0.0.12
python-dotenv>=1.0.1
supabase>=2.10.0
pymupdf>=1.24.0
```

설치:

```bash
cd parser-api
pip install -r requirements.txt
```

---

## 3. LibreOffice 설치

### Windows (개발)

1. [LibreOffice 공식 사이트](https://www.libreoffice.org/download/download-libreoffice/)에서 설치 파일 다운로드
2. 기본 경로에 설치: `C:\Program Files\LibreOffice\program\soffice.exe`
3. 다른 경로에 설치한 경우 환경 변수로 지정:
   ```
   PPTX_LIBREOFFICE_PATH=C:\custom\path\soffice.exe
   ```

### Ubuntu / Debian (프로덕션)

```bash
apt-get update
apt-get install -y libreoffice libreoffice-writer libreoffice-impress

# 한글 렌더링을 위한 추가 패키지
apt-get install -y fonts-nanum fonts-noto-cjk

# 설치 확인
soffice --version
```

### Docker

```dockerfile
FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-impress \
    fonts-nanum \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY parser-api/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY parser-api/ .
CMD ["python", "-m", "uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
```

---

## 4. 폰트 설치 (래스터 텍스트 누락 방지)

LibreOffice headless 변환 시 **PPTX에 사용된 폰트가 서버에 없으면 해당 텍스트 블록 전체가 래스터에서 사라집니다.**
폰트가 누락된 경우, 파싱 응답의 `meta.slideRaster.missingFonts` 필드에 목록이 표시됩니다.

### 4-1. Freesentation (프리젠테이션) 폰트

이 프로젝트의 PPT 파일은 [Freesentation](https://freesentation.blog/freesentation) 폰트를 사용합니다.
SIL 오픈폰트라이선스(OFL) - 상업 사용 무료.

**다운로드 (ver 1.006 권장 - PPT 최적화 경량 버전):**

```
https://github.com/Freesentation/freesentation/raw/refs/heads/main/Freesentation_1.006.zip
```

#### Windows에 설치

탐색기에서 `Freesentation-*.ttf` 9개 파일 선택 → 우클릭 → **"모든 사용자를 위해 설치"**

LibreOffice 폰트 폴더에도 추가(관리자 권한):
```
C:\Program Files\LibreOffice\program\resource\common\fonts\
```

#### Linux / Docker에 설치

```bash
# TTF 파일을 서버에 업로드한 뒤
mkdir -p /usr/share/fonts/truetype/freesentation
cp Freesentation-*.ttf /usr/share/fonts/truetype/freesentation/
fc-cache -f -v

# 설치 확인
fc-list | grep -i freesentation
```

Dockerfile에 포함하려면:

```dockerfile
COPY fonts/Freesentation-*.ttf /usr/share/fonts/truetype/freesentation/
RUN fc-cache -f -v
```

### 4-2. 기타 한글 폰트 (Linux 서버 필수)

```bash
# Nanum 폰트 (나눔고딕, 나눔명조 등)
apt-get install -y fonts-nanum

# Noto CJK (구글 범용 한중일 폰트)
apt-get install -y fonts-noto-cjk

# 폰트 캐시 갱신
fc-cache -f -v
```

### 4-3. 폰트 추가 시 일반 절차

```bash
# 1. TTF/OTF 파일을 폰트 디렉터리에 복사
cp MyFont.ttf /usr/share/fonts/truetype/
# 또는 사용자 전용
cp MyFont.ttf ~/.local/share/fonts/

# 2. 캐시 갱신
fc-cache -f -v

# 3. 확인
fc-list | grep "MyFont"

# 4. LibreOffice 재시작 (daemon이 아닌 headless 방식은 매 실행마다 폰트 재탐색)
```

---

## 5. 환경 변수

### parser-api (`parser-api/.env`)

| 변수 | 기본값 | 설명 |
|------|--------|------|
| `PPTX_LIBREOFFICE_PATH` | 자동 탐색 | `soffice` 실행 파일 경로 (기본 위치가 아닌 경우) |
| `PPTX_SLIDE_RASTER` | `1` | `0`으로 설정하면 래스터 생성 비활성화 |
| `PPTX_RASTER_MAX_LONG_EDGE` | `1280` | 래스터 이미지 긴 변 최대 픽셀 (320~4096) |
| `PPTX_RASTER_JPEG_QUALITY` | `82` | JPEG 품질 (50~95) |
| `PPTX_RASTER_MAX_BYTES_PER_SLIDE` | `2000000` | 슬라이드당 JPEG 최대 바이트 (초과 시 자동 품질 하향) |
| `PPTX_RASTER_TIMEOUT_SEC` | `240` | LibreOffice 변환 타임아웃(초) |
| `PPTX_IMAGE_INLINE_MAX_BYTES` | `12000000` | 도형 이미지 inline data URL 최대 바이트 |
| `SUPABASE_URL` | - | Supabase 프로젝트 URL (이미지 Storage 업로드용) |
| `SUPABASE_SERVICE_ROLE_KEY` | - | Supabase service role key |
| `STORAGE_BUCKET` | `lecture-assets` | Supabase Storage 버킷명 |

### Next.js 프론트 (`frontend/.env.local` → Vercel 환경 변수)

| 변수 | 설명 |
|------|------|
| `NEXT_PUBLIC_SUPABASE_URL` | Supabase 프로젝트 URL |
| `NEXT_PUBLIC_SUPABASE_ANON_KEY` | Supabase anon key |
| `NEXT_PUBLIC_PARSER_API_URL` | parser-api 서버 URL (브라우저 → Next 프록시용 표시) |
| `PARSER_API_URL` | parser-api 서버 URL (서버 사이드 프록시 실제 호출용, `NEXT_PUBLIC_` 없음) |
| `SUPABASE_SERVICE_ROLE_KEY` | Supabase service role key (서버 전용, 노출 금지) |
| `DATABASE_URL` | Prisma 연결 문자열 (pgBouncer 포함) |
| `OPENAI_API_KEY` | OpenAI API 키 (RAG·임베딩용, 서버 전용) |
| `ADMIN_EMAILS` | 관리자 이메일 목록 (쉼표 구분) |

---

## 6. 실행

### 개발 (Windows)

```powershell
# parser-api
cd parser-api
..\venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m uvicorn main:app --reload --port 8010

# frontend (별도 터미널)
cd frontend
npm install
npm run dev
```

### 프로덕션 (Linux)

```bash
# parser-api (gunicorn + uvicorn worker 권장)
cd /app
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8010 --workers 2 --timeout-keep-alive 300

# 또는 systemd / Docker Compose로 관리
```

---

## 7. 권장 호스팅 옵션 (parser-api)

| 서비스 | 특징 | 비고 |
|--------|------|------|
| **Fly.io** | 컨테이너, 저비용, 긴 타임아웃 | `fly.toml`에 `[http_service] internal_port = 8010` |
| **Railway** | Git 연동 자동 배포, 간단 | 메모리 제한 확인 필요 |
| **Render** | Docker 지원, 무료 tier 존재 | Cold start 주의 |
| **Google Cloud Run** | 요청당 과금, 자동 스케일 | `--timeout=300` 필수 |
| **AWS App Runner / ECS** | 엔터프라이즈급 | 설정 복잡도 높음 |

> LibreOffice가 포함된 Docker 이미지는 약 **1.5~2 GB** 수준입니다.
> 이미지 크기가 부담되면 `PPTX_SLIDE_RASTER=0`으로 래스터를 비활성화하고
> 텍스트/도형 JSON만 제공하는 경량 모드로 운용할 수 있습니다.

---

## 8. 래스터 품질 트러블슈팅

| 증상 | 원인 | 조치 |
|------|------|------|
| 텍스트 블록 통째 누락 | 해당 폰트 미설치 | `meta.slideRaster.missingFonts` 확인 후 폰트 설치 |
| 전체 슬라이드 검정/흰색 | LO PDF 변환 실패 | `meta.slideRaster.reason` 확인, LO 재설치 |
| 일부 도형 색상 이상 | LO 테마/효과 미지원 | 래스터 한계 (검수용으로만 활용) |
| JPEG 0장, status=partial | PyMuPDF alpha 채널 문제 | `pymupdf>=1.24` 업그레이드 |
| 슬라이드·PDF 페이지 수 불일치 | LO 슬라이드 병합/분리 | `meta.slideRaster.pageCountMismatch` 확인 |

---

*최종 수정: 2026-04-01*
