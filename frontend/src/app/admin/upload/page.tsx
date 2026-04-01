"use client";

import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
} from "react";
import { SlideViewer } from "@/components/SlideViewer";
import { slideCanvasPaddingBottomPercent } from "@/lib/slideCanvasAspect";
import type { ParsePptxResponse, SlideData } from "@/types/parse";

/** Same-origin proxy → `src/app/api/parse-pptx/route.ts` (Network 탭에 `parse-pptx`로 보임). */
const PARSE_ENDPOINT = "/api/parse-pptx";

/** 미리보기용: data URL은 본문 대신 크기만 표시해 UI·DevTools가 멈추지 않게 함 */
function slideJsonForPreview(slide: SlideData): string {
  return JSON.stringify(
    slide,
    (key, value) => {
      if (
        key === "rasterPreview" &&
        typeof value === "string" &&
        value.startsWith("data:")
      ) {
        return `[슬라이드 전체 래스터(JPEG) 포함, 약 ${Math.round(value.length / 1024)} KB — 가운데 패널 하단 이미지 / 전체 JSON 다운로드]`;
      }
      if (
        typeof value === "string" &&
        value.startsWith("data:") &&
        value.length > 120
      ) {
        return `[도형·이미지 data URL 축약, ${Math.round(value.length / 1024)} KB — 전체 JSON 다운로드]`;
      }
      return value;
    },
    2,
  );
}

/** 슬라이드 순서대로 비어 있지 않은 슬라이드 메모만 이어붙여 강의 단위 Description으로 씀 */
function aggregateDescriptionFromSlideNotes(
  slides: SlideData[],
  notes: Record<number, string>,
): string {
  const chunks: string[] = [];
  for (const s of slides) {
    const text = (notes[s.slideNumber] ?? "").trim();
    if (text) {
      chunks.push(`[슬라이드 ${s.slideNumber}]\n${text}`);
    }
  }
  return chunks.join("\n\n");
}

/** 비어 있지 않은 첫 슬라이드 메모의 첫 줄을 제목 후보로 사용 (최대 길이 제한) */
function deriveTitleFromSlideNotes(
  slides: SlideData[],
  notes: Record<number, string>,
  maxLen = 120,
): string {
  for (const s of slides) {
    const raw = (notes[s.slideNumber] ?? "").trim();
    if (!raw) continue;
    const line = raw.split(/\r?\n/)[0].trim();
    if (!line) continue;
    if (line.length <= maxLen) return line;
    return `${line.slice(0, maxLen - 1)}…`;
  }
  return "";
}

export default function AdminUploadPage() {
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [parserStale, setParserStale] = useState(false);
  const [data, setData] = useState<ParsePptxResponse | null>(null);
  const [slideIndex, setSlideIndex] = useState(0);
  const [slideNotes, setSlideNotes] = useState<Record<number, string>>({});
  const [title, setTitle] = useState("");
  const [tags, setTags] = useState("");
  const titleEditedByUser = useRef(false);
  const previewCardRef = useRef<HTMLDivElement>(null);
  const [previewColumnHeightPx, setPreviewColumnHeightPx] = useState<
    number | null
  >(null);
  const [isLgLayout, setIsLgLayout] = useState(false);

  const onFile = useCallback(
    async (file: File | null) => {
      if (!file) return;
      setError(null);
      setParserStale(false);
      setBusy(true);
      setData(null);
      try {
        const fd = new FormData();
        fd.append("file", file);
        const res = await fetch(PARSE_ENDPOINT, {
          method: "POST",
          body: fd,
        });
        if (!res.ok) {
          const raw = await res.text();
          let msg = raw || `HTTP ${res.status}`;
          try {
            const j = JSON.parse(raw) as { error?: string; detail?: string };
            if (j?.error) msg = j.error;
            else if (typeof j?.detail === "string") msg = j.detail;
          } catch {
            /* keep raw */
          }
          throw new Error(msg);
        }
        const json = (await res.json()) as ParsePptxResponse;
        const hdrBuild = res.headers.get("x-pptx-parser-build");
        const metaBuild = json.meta?.parserApiBuild;
        setParserStale(!metaBuild && !hdrBuild);
        setData(json);
        setSlideIndex(0);
        setSlideNotes({});
        titleEditedByUser.current = false;
        setTitle(json.meta?.title ?? "");
        setTags((json.meta?.tags ?? []).join(", "));
      } catch (e) {
        setError(e instanceof Error ? e.message : "parse failed");
      } finally {
        setBusy(false);
      }
    },
    [],
  );

  const slides: SlideData[] = data?.slides ?? [];
  const current: SlideData | undefined = slides[slideIndex];
  const canPrevSlide = slideIndex > 0;
  const canNextSlide = slideIndex < slides.length - 1;
  const hasParsedData = Boolean(data && slides.length > 0);
  const currentSlideNumber = current?.slideNumber;
  const currentSlideNote =
    currentSlideNumber !== undefined ? (slideNotes[currentSlideNumber] ?? "") : "";
  const descriptionFromSlides = useMemo(
    () => aggregateDescriptionFromSlideNotes(slides, slideNotes),
    [slides, slideNotes],
  );
  const fallbackTitleFromParser = data?.meta?.title ?? "";
  const slideRasterMeta = data?.meta?.slideRaster;
  const { slideAspectRatio, slideCanvasPaddingBottomPct } = useMemo(() => {
    const emu = data?.meta?.slideSizeEmu;
    const w = emu?.width;
    const h = emu?.height;
    const slideCanvasPaddingBottomPct = slideCanvasPaddingBottomPercent(emu);
    if (typeof w === "number" && typeof h === "number" && w > 0 && h > 0) {
      return {
        slideAspectRatio: `${w} / ${h}`,
        slideCanvasPaddingBottomPct,
      };
    }
    return {
      slideAspectRatio: "16 / 9",
      slideCanvasPaddingBottomPct,
    };
  }, [data?.meta?.slideSizeEmu]);

  useEffect(() => {
    if (titleEditedByUser.current) return;
    const fromNotes = deriveTitleFromSlideNotes(slides, slideNotes);
    setTitle(fromNotes || fallbackTitleFromParser);
  }, [slides, slideNotes, fallbackTitleFromParser]);

  useEffect(() => {
    const mq = window.matchMedia("(min-width: 1024px)");
    const sync = () => setIsLgLayout(mq.matches);
    sync();
    mq.addEventListener("change", sync);
    return () => mq.removeEventListener("change", sync);
  }, []);

  useLayoutEffect(() => {
    if (!hasParsedData || !current) {
      setPreviewColumnHeightPx(null);
      return;
    }
    const el = previewCardRef.current;
    if (!el) return;
    const measure = () => {
      const h = el.getBoundingClientRect().height;
      if (h > 0) setPreviewColumnHeightPx(Math.round(h));
    };
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [
    hasParsedData,
    current,
    slideIndex,
    data,
    slideCanvasPaddingBottomPct,
    slideAspectRatio,
  ]);

  const sideCardHeightStyle: CSSProperties | undefined =
    isLgLayout && previewColumnHeightPx != null
      ? { height: previewColumnHeightPx, minHeight: 0 }
      : undefined;

  const sideCardHeightClass =
    isLgLayout && previewColumnHeightPx != null
      ? ""
      : "max-h-[calc(100vh-6rem)]";

  return (
    <main className="min-h-screen bg-zinc-50">
      {/* 상단 헤더 */}
      <div className="border-b border-zinc-200 bg-white">
        <div className="mx-auto flex max-w-[1600px] items-center justify-between px-6 py-4">
          <h1 className="text-base font-semibold tracking-tight text-zinc-900">
            PPTX Upload &amp; Preview
          </h1>
          {hasParsedData ? (
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => {
                  const blob = new Blob([JSON.stringify(data, null, 2)], {
                    type: "application/json;charset=utf-8",
                  });
                  const a = document.createElement("a");
                  a.href = URL.createObjectURL(blob);
                  a.download = "parse-pptx-full-response.json";
                  a.click();
                  URL.revokeObjectURL(a.href);
                }}
                className="inline-flex items-center gap-1.5 rounded-lg border border-zinc-200 bg-white px-3 py-2 text-xs font-medium text-zinc-600 shadow-sm transition hover:border-zinc-300 hover:bg-zinc-50"
              >
                <svg className="h-3.5 w-3.5" viewBox="0 0 16 16" fill="currentColor" aria-hidden>
                  <path d="M7.25 10.44V2.75a.75.75 0 0 1 1.5 0v7.69l2.72-2.72a.75.75 0 1 1 1.06 1.06l-4 4a.75.75 0 0 1-1.06 0l-4-4a.75.75 0 1 1 1.06-1.06l2.72 2.72Z" />
                  <path d="M2.75 13.25a.75.75 0 0 0 0 1.5h10.5a.75.75 0 0 0 0-1.5H2.75Z" />
                </svg>
                JSON 다운로드
              </button>

              <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg bg-zinc-900 px-3 py-2 text-xs font-medium text-white shadow-sm transition hover:bg-zinc-700">
                <svg className="h-3.5 w-3.5" viewBox="0 0 16 16" fill="currentColor" aria-hidden>
                  <path fillRule="evenodd" d="M14 8a.75.75 0 0 1-.75.75H4.56l3.22 3.22a.75.75 0 1 1-1.06 1.06l-4.5-4.5a.75.75 0 0 1 0-1.06l4.5-4.5a.75.75 0 0 1 1.06 1.06L4.56 7.25h8.69A.75.75 0 0 1 14 8Z" clipRule="evenodd" />
                </svg>
                새 파일 업로드
                <input
                  type="file"
                  accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
                  disabled={busy}
                  onChange={(e) => {
                    setData(null);
                    setSlideIndex(0);
                    setSlideNotes({});
                    setTitle("");
                    setTags("");
                    titleEditedByUser.current = false;
                    void onFile(e.target.files?.[0] ?? null);
                  }}
                  className="sr-only"
                />
              </label>
            </div>
          ) : null}
        </div>
      </div>

      <div className="mx-auto max-w-[1600px] px-6 py-6">

      {/* 에러 */}
      {error ? (
        <div className="mb-4 flex items-start gap-3 rounded-xl border border-red-200 bg-red-50 px-4 py-3">
          <svg className="mt-0.5 h-4 w-4 shrink-0 text-red-500" viewBox="0 0 16 16" fill="currentColor" aria-hidden>
            <path fillRule="evenodd" d="M8 1a7 7 0 1 0 0 14A7 7 0 0 0 8 1ZM7.25 5.75a.75.75 0 0 1 1.5 0v3.5a.75.75 0 0 1-1.5 0v-3.5Zm.75 6.5a.75.75 0 1 0 0-1.5.75.75 0 0 0 0 1.5Z" clipRule="evenodd" />
          </svg>
          <p className="text-sm text-red-800">{error}</p>
        </div>
      ) : null}

      {parserStale && data ? (
        <div className="mb-4 flex items-start gap-3 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3">
          <svg className="mt-0.5 h-4 w-4 shrink-0 text-amber-500" viewBox="0 0 16 16" fill="currentColor" aria-hidden>
            <path fillRule="evenodd" d="M8 1a7 7 0 1 0 0 14A7 7 0 0 0 8 1ZM7.25 5.75a.75.75 0 0 1 1.5 0v3.5a.75.75 0 0 1-1.5 0v-3.5Zm.75 6.5a.75.75 0 1 0 0-1.5.75.75 0 0 0 0 1.5Z" clipRule="evenodd" />
          </svg>
          <div className="text-sm text-amber-900">
            <strong>parser-api 버전 불일치</strong> — 최신 코드로 백엔드를 재시작하세요.
          </div>
        </div>
      ) : null}

      {/* 파싱 결과 요약 배너 */}
      {hasParsedData && data?.meta?.slideRaster ? (
        <div className="mb-5 flex flex-wrap items-center gap-2 rounded-xl border border-zinc-200 bg-white px-4 py-2.5 shadow-sm">
          {/* 엔진 배지 */}
          {data.meta.slideRaster.engine === "powerpoint-com" ? (
            <span className="inline-flex items-center gap-1 rounded-full bg-emerald-50 px-2.5 py-0.5 text-xs font-medium text-emerald-700 ring-1 ring-emerald-200">
              <span className="h-1.5 w-1.5 rounded-full bg-emerald-500" />
              PowerPoint COM
            </span>
          ) : data.meta.slideRaster.engine === "libreoffice" ? (
            <span className="inline-flex items-center gap-1 rounded-full bg-sky-50 px-2.5 py-0.5 text-xs font-medium text-sky-700 ring-1 ring-sky-200">
              <span className="h-1.5 w-1.5 rounded-full bg-sky-500" />
              LibreOffice
            </span>
          ) : null}
          {/* 슬라이드 수 */}
          <span className="text-xs text-zinc-500">
            <span className="font-medium text-zinc-800">{slides.length}</span>장 슬라이드
          </span>
          {typeof data.meta.slideRaster.slidesRendered === "number" ? (
            <span className="text-xs text-zinc-500">
              래스터 <span className="font-medium text-zinc-800">{data.meta.slideRaster.slidesRendered}</span>/{slides.length}장
            </span>
          ) : null}
          {/* 파일명 */}
          {data.meta.title ? (
            <span className="ml-auto truncate text-xs text-zinc-400">{data.meta.title}</span>
          ) : null}

          {/* PPT COM fallback 경고 */}
          {data.meta.slideRaster.pptComFallbackReason ? (
            <div className="w-full border-t border-zinc-100 pt-2 text-xs text-amber-700">
              PowerPoint COM 실패, LibreOffice 사용 중
              <span className="ml-1 font-mono text-[11px] opacity-75">({data.meta.slideRaster.pptComFallbackReason})</span>
            </div>
          ) : null}
          {/* 폰트 누락 경고 */}
          {data.meta.slideRaster.missingFonts?.length ? (
            <div className="w-full border-t border-zinc-100 pt-2 text-xs text-red-700">
              미설치 폰트 감지: <span className="font-mono">{data.meta.slideRaster.missingFonts.join(", ")}</span>
            </div>
          ) : data.meta.slideRaster.pptxFonts?.length ? (
            <details className="w-full border-t border-zinc-100 pt-2">
              <summary className="cursor-pointer text-xs text-zinc-400">
                폰트 {data.meta.slideRaster.pptxFonts.length}종 (모두 설치됨)
              </summary>
              <p className="mt-1 font-mono text-[11px] text-zinc-500">{data.meta.slideRaster.pptxFonts.join(", ")}</p>
            </details>
          ) : null}
        </div>
      ) : null}

      {hasParsedData ? (
        <div className="grid grid-cols-1 gap-5 lg:grid-cols-[minmax(0,18rem)_minmax(0,1fr)_minmax(0,22rem)] lg:items-stretch">

          {/* ── 좌: 메타데이터 카드 ── */}
          <aside
            className={`order-2 flex min-h-0 w-full min-w-0 flex-col overflow-hidden rounded-2xl border border-zinc-200 bg-white shadow-sm lg:sticky lg:top-4 lg:order-1 ${sideCardHeightClass}`}
            style={sideCardHeightStyle}
          >
            <div className="flex h-14 shrink-0 items-center border-b border-zinc-100 px-5">
              <div>
                <h2 className="text-[13px] font-semibold text-zinc-800">Metadata</h2>
                <p className="text-[11px] text-zinc-400">제목·태그·슬라이드 설명 입력</p>
              </div>
            </div>

            <div className="flex min-h-0 flex-1 flex-col gap-4 overflow-y-auto overscroll-contain p-5">
              <div>
                <label className="mb-1 block text-[11px] font-medium uppercase tracking-wide text-zinc-400">
                  Title
                </label>
                <input
                  value={title}
                  onChange={(e) => {
                    titleEditedByUser.current = true;
                    setTitle(e.target.value);
                  }}
                  placeholder="강의 제목"
                  className="w-full rounded-lg border border-zinc-200 bg-zinc-50 px-3 py-2 text-sm text-zinc-800 placeholder-zinc-300 transition focus:border-zinc-400 focus:bg-white focus:outline-none focus:ring-2 focus:ring-zinc-200"
                />
              </div>

              <div>
                <label className="mb-1 block text-[11px] font-medium uppercase tracking-wide text-zinc-400">
                  Tags
                </label>
                <input
                  value={tags}
                  onChange={(e) => setTags(e.target.value)}
                  placeholder="ai, 분석, agent"
                  className="w-full rounded-lg border border-zinc-200 bg-zinc-50 px-3 py-2 text-sm text-zinc-800 placeholder-zinc-300 transition focus:border-zinc-400 focus:bg-white focus:outline-none focus:ring-2 focus:ring-zinc-200"
                />
              </div>

              {current ? (
                <div>
                  <label className="mb-1 block text-[11px] font-medium uppercase tracking-wide text-zinc-400">
                    슬라이드 {current.slideNumber} 설명
                  </label>
                  <textarea
                    value={currentSlideNote}
                    onChange={(e) => {
                      const next = e.target.value;
                      setSlideNotes((prev) => ({
                        ...prev,
                        [current.slideNumber]: next,
                      }));
                    }}
                    rows={5}
                    placeholder="강사가 이 슬라이드에서 전달할 핵심 메시지를 작성하세요."
                    className="w-full resize-none rounded-lg border border-zinc-200 bg-zinc-50 px-3 py-2 text-sm text-zinc-800 placeholder-zinc-300 transition focus:border-zinc-400 focus:bg-white focus:outline-none focus:ring-2 focus:ring-zinc-200"
                  />
                </div>
              ) : null}

              <div>
                <label className="mb-1 block text-[11px] font-medium uppercase tracking-wide text-zinc-400">
                  Description <span className="normal-case font-normal text-zinc-300">(자동 합본)</span>
                </label>
                <textarea
                  readOnly
                  value={descriptionFromSlides}
                  rows={4}
                  placeholder="슬라이드마다 설명을 입력하면 여기에 합쳐집니다."
                  className="w-full resize-none cursor-default rounded-lg border border-dashed border-zinc-200 bg-zinc-50/60 px-3 py-2 text-sm text-zinc-600 placeholder-zinc-300"
                />
              </div>

              {/* 유틸 버튼 */}
              <div className="mt-auto space-y-2 border-t border-zinc-100 pt-4">
                <div className="grid grid-cols-2 gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setTitle(data?.meta?.title ?? "");
                      setTags((data?.meta?.tags ?? []).join(", "));
                      titleEditedByUser.current = true;
                    }}
                    className="rounded-lg border border-zinc-200 bg-white py-1.5 text-xs font-medium text-zinc-600 transition hover:border-zinc-300 hover:bg-zinc-50"
                  >
                    초안값 복원
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      titleEditedByUser.current = false;
                      setTitle("");
                      setTags("");
                      setSlideNotes({});
                    }}
                    className="rounded-lg border border-zinc-200 bg-white py-1.5 text-xs font-medium text-zinc-600 transition hover:border-zinc-300 hover:bg-zinc-50"
                  >
                    전체 초기화
                  </button>
                </div>
                <button
                  type="button"
                  onClick={() => {
                    titleEditedByUser.current = false;
                    const t = deriveTitleFromSlideNotes(slides, slideNotes) || fallbackTitleFromParser;
                    setTitle(t);
                  }}
                  className="w-full rounded-lg border border-zinc-200 bg-white py-1.5 text-xs text-zinc-500 transition hover:border-zinc-300 hover:bg-zinc-50"
                >
                  슬라이드 설명 기준으로 제목 맞추기
                </button>
                <p className="text-center text-[11px] text-zinc-300">
                  Publish · embeddings · DB insert — 다음 단계
                </p>
              </div>
            </div>
          </aside>

          {/* ── 가운데: 슬라이드 미리보기 카드 ── */}
          <div className="order-1 flex h-full min-h-0 w-full min-w-0 items-start lg:order-2">
            {current ? (
              <div
                ref={previewCardRef}
                className="flex w-full min-w-0 flex-col overflow-x-hidden rounded-2xl border border-zinc-200 bg-white shadow-sm"
              >
                {/* 슬라이드 정보 바 */}
                <div className="flex h-14 shrink-0 items-center justify-between border-b border-zinc-100 px-5">
                  <span className="text-[13px] font-semibold text-zinc-800">
                    슬라이드 {current.slideNumber}
                    <span className="ml-2 text-xs font-normal text-zinc-400">
                      도형 {(current.elements ?? []).length}개
                      {(current.plainText ?? "").trim()
                        ? ` · ${(current.plainText ?? "").length}자`
                        : ""}
                    </span>
                  </span>
                  {/* 네비게이션 */}
                  <nav className="flex items-center gap-1.5" aria-label="슬라이드 이동">
                    <button
                      type="button"
                      disabled={!canPrevSlide}
                      onClick={() => setSlideIndex((i) => Math.max(0, i - 1))}
                      aria-label="이전 슬라이드"
                      className="inline-flex h-7 w-7 items-center justify-center rounded-lg border border-zinc-200 bg-white text-zinc-500 transition hover:border-zinc-300 hover:bg-zinc-50 disabled:pointer-events-none disabled:opacity-30"
                    >
                      <svg className="h-4 w-4" viewBox="0 0 20 20" fill="currentColor" aria-hidden>
                        <path fillRule="evenodd" d="M12.79 5.23a.75.75 0 01-.02 1.06L8.832 10l3.938 3.71a.75.75 0 11-1.04 1.08l-4.5-4.25a.75.75 0 010-1.08l4.5-4.25a.75.75 0 011.06.02z" clipRule="evenodd" />
                      </svg>
                    </button>
                    <span className="min-w-[4.5rem] select-none text-center text-xs tabular-nums text-zinc-500">
                      <span className="font-semibold text-zinc-800">{slideIndex + 1}</span>
                      <span className="text-zinc-300"> / </span>
                      {slides.length}
                    </span>
                    <button
                      type="button"
                      disabled={!canNextSlide}
                      onClick={() => setSlideIndex((i) => Math.min(slides.length - 1, i + 1))}
                      aria-label="다음 슬라이드"
                      className="inline-flex h-7 w-7 items-center justify-center rounded-lg border border-zinc-200 bg-white text-zinc-500 transition hover:border-zinc-300 hover:bg-zinc-50 disabled:pointer-events-none disabled:opacity-30"
                    >
                      <svg className="h-4 w-4" viewBox="0 0 20 20" fill="currentColor" aria-hidden>
                        <path fillRule="evenodd" d="M7.21 14.77a.75.75 0 01.02-1.06L11.168 10 7.23 6.29a.75.75 0 111.04-1.08l4.5 4.25a.75.75 0 010 1.08l-4.5 4.25a.75.75 0 01-1.06-.02z" clipRule="evenodd" />
                      </svg>
                    </button>
                  </nav>
                </div>

                <div className="p-4 space-y-4">
                  {/* 추출 미리보기 */}
                  <div>
                    <p className="mb-2 text-[11px] font-medium uppercase tracking-wide text-zinc-400">
                      HTML 렌더링 <span className="normal-case font-normal">— 텍스트·도형을 퍼센트 좌표로 재현</span>
                    </p>
                    <SlideViewer
                      elements={current.elements ?? []}
                      className="mx-auto w-full max-w-full"
                      aspectRatio={slideAspectRatio}
                      canvasPaddingBottomPercent={slideCanvasPaddingBottomPct}
                    />
                  </div>

                  {/* 래스터 미리보기 */}
                  {current.rasterPreview ? (
                    <div>
                      <p className="mb-2 text-[11px] font-medium uppercase tracking-wide text-zinc-400">
                        원본 슬라이드 이미지 <span className="normal-case font-normal">— PowerPoint가 직접 렌더링한 결과</span>
                      </p>
                      <div
                        className="relative isolate h-0 w-full overflow-hidden rounded-xl border border-zinc-100 bg-zinc-900 shadow-inner"
                        style={{ paddingBottom: `${slideCanvasPaddingBottomPct}%` }}
                      >
                        {/* eslint-disable-next-line @next/next/no-img-element */}
                        <img
                          src={current.rasterPreview}
                          alt=""
                          className="absolute inset-0 h-full w-full object-contain"
                        />
                      </div>
                    </div>
                  ) : null}

                  {/* 래스터 오류 메시지 */}
                  {slideRasterMeta?.status === "ok" && !current.rasterPreview ? (
                    <p className="rounded-lg bg-amber-50 px-3 py-2 text-center text-[11px] text-amber-700">
                      이 슬라이드의 래스터가 없습니다. (렌더 실패 또는 응답 불완전)
                    </p>
                  ) : null}
                  {slideRasterMeta?.status && slideRasterMeta.status !== "ok" ? (
                    <p className={`rounded-lg px-3 py-2 text-center text-[11px] ${
                      slideRasterMeta.status === "error" ? "bg-red-50 text-red-800" : "bg-amber-50 text-amber-800"
                    }`}>
                      래스터 {slideRasterMeta.status}: {slideRasterMeta.reason || "(사유 없음)"}
                    </p>
                  ) : null}
                </div>
              </div>
            ) : null}
          </div>

          {/* ── 우: JSON 뷰어 카드 ── */}
          <aside
            className={`order-3 flex min-h-0 w-full min-w-0 flex-col overflow-hidden rounded-2xl border border-zinc-200 bg-white shadow-sm lg:order-3 lg:sticky lg:top-4 ${sideCardHeightClass}`}
            style={sideCardHeightStyle}
          >
            {current ? (
              <>
                <div className="flex h-14 shrink-0 items-center border-b border-zinc-100 px-5">
                  <div>
                    <h2 className="text-[13px] font-semibold text-zinc-800">JSON</h2>
                    <p className="text-[11px] text-zinc-400">파서가 추출한 좌표·텍스트·이미지 구조 데이터</p>
                  </div>
                </div>
                <pre className="min-h-0 min-w-0 flex-1 overflow-x-auto overflow-y-auto overscroll-contain break-all bg-zinc-950 p-4 font-mono text-[11px] leading-relaxed text-emerald-300">
                  {slideJsonForPreview(current)}
                </pre>
              </>
            ) : null}
          </aside>
        </div>
      ) : null}

      {/* 업로드 드롭존 */}
      {!hasParsedData ? (
        <label
          className={`mx-auto mb-6 flex max-w-xl cursor-pointer flex-col items-center justify-center gap-4 rounded-2xl border-2 border-dashed py-14 text-center transition ${
            busy
              ? "cursor-not-allowed border-zinc-200 bg-zinc-50 opacity-60"
              : "border-zinc-300 bg-white hover:border-zinc-400 hover:bg-zinc-50"
          }`}
        >
          <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-zinc-100 text-zinc-400">
            {busy ? (
              <svg className="h-7 w-7 animate-spin" viewBox="0 0 24 24" fill="none" aria-hidden>
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" />
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v4l3-3-3-3v4a8 8 0 00-8 8h4z" />
              </svg>
            ) : (
              <svg className="h-7 w-7" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={1.5} aria-hidden>
                <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 7.5m0 0L7.5 12M12 7.5v9" />
              </svg>
            )}
          </div>
          <div>
            <p className="text-sm font-semibold text-zinc-700">
              {busy ? "파싱 중..." : "PPTX 파일을 선택하세요"}
            </p>
            {!busy ? (
              <p className="mt-1 text-xs text-zinc-400">.pptx 파일만 지원합니다</p>
            ) : null}
          </div>
          {!busy ? (
            <span className="rounded-lg bg-zinc-900 px-5 py-2 text-sm font-medium text-white transition hover:bg-zinc-700">
              파일 선택
            </span>
          ) : null}
          <input
            type="file"
            accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
            disabled={busy}
            onChange={(e) => void onFile(e.target.files?.[0] ?? null)}
            className="sr-only"
          />
        </label>
      ) : null}
      </div>
    </main>
  );
}
