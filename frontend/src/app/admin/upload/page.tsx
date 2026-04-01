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
    <main className="mx-auto max-w-[1600px] px-4 py-8">
      <h1 className="text-xl font-semibold">Upload / Preview</h1>
      <p className="mt-1 text-sm text-zinc-600">
        .pptx 파일을 선택하면 parser-api가 JSON을 반환합니다. DB에는 아직 저장하지
        않습니다.
      </p>

      <div className="mt-6 flex flex-wrap items-center gap-4">
        <input
          type="file"
          accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
          disabled={busy}
          onChange={(e) => void onFile(e.target.files?.[0] ?? null)}
          className="text-sm"
        />
        {busy ? (
          <span className="text-sm text-zinc-500">Parsing...</span>
        ) : null}
        {hasParsedData ? (
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
            className="rounded border border-zinc-300 bg-white px-3 py-2 text-xs text-zinc-700 hover:bg-zinc-50"
          >
            전체 응답 JSON 다운로드
          </button>
        ) : null}
      </div>

      {hasParsedData ? (
        <div className="mt-3 rounded-lg border border-zinc-200 bg-zinc-50 px-3 py-2 text-sm text-zinc-700">
          <span className="font-medium text-zinc-800">슬라이드 래스터</span>
          {data?.meta?.slideRaster ? (
            <>
              {": "}
              <code className="rounded bg-white px-1 text-xs">
                {data.meta.slideRaster.status}
              </code>
              {typeof data.meta.slideRaster.slidesRendered === "number" ? (
                <span className="text-zinc-600">
                  {" "}
                  ({data.meta.slideRaster.slidesRendered}/{slides.length}장 JPEG)
                </span>
              ) : null}
              {data.meta.slideRaster.reason ? (
                <span className="mt-1 block text-xs text-zinc-600">
                  {data.meta.slideRaster.reason}
                </span>
              ) : null}
              {data.meta.slideRaster.missingFonts?.length ? (
                <div className="mt-2 rounded border border-red-200 bg-red-50 px-2 py-1.5 text-xs text-red-900">
                  <strong>래스터 텍스트 누락 원인 — 미설치 폰트 감지됨</strong>
                  <div className="mt-1 font-mono">
                    {data.meta.slideRaster.missingFonts.join(", ")}
                  </div>
                  <div className="mt-1 text-red-700">
                    {data.meta.slideRaster.missingFontHint}
                  </div>
                </div>
              ) : data.meta.slideRaster.pptxFonts?.length ? (
                <details className="mt-1">
                  <summary className="cursor-pointer text-xs text-zinc-500">
                    사용된 폰트 {data.meta.slideRaster.pptxFonts.length}종 (모두 시스템에 있음)
                  </summary>
                  <div className="mt-0.5 font-mono text-[11px] text-zinc-600">
                    {data.meta.slideRaster.pptxFonts.join(", ")}
                  </div>
                </details>
              ) : null}
            </>
          ) : (
            <span className="text-amber-800">
              {" "}
              meta.slideRaster 없음 — parser-api를 최신 코드로 띄웠는지,{" "}
              <code className="rounded bg-amber-100 px-1">pip install -r requirements.txt</code>{" "}
              후 재시작했는지 확인하세요.
            </span>
          )}
          <span className="mt-1 block text-xs text-zinc-500">
            우측 JSON의{" "}
            <code className="rounded bg-zinc-200 px-0.5">elements[].src</code>는
            용량 때문에 축약 표시됩니다. 슬라이드 통째 이미지는{" "}
            <code className="rounded bg-zinc-200 px-0.5">rasterPreview</code> 필드(동일
            축약)와 가운데 하단 미리보기입니다.
          </span>
        </div>
      ) : null}

      {error ? (
        <p className="mt-4 rounded border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-800">
          {error}
        </p>
      ) : null}

      {parserStale && data ? (
        <p className="mt-4 rounded border border-amber-300 bg-amber-50 px-3 py-2 text-sm text-amber-950">
          연결된 파서가 이 프로젝트의 최신 parser-api가 아닌 것 같습니다. (
          <code className="rounded bg-amber-100 px-1">meta.parserApiBuild</code> 없음)
          <br />
          <span className="mt-1 block text-amber-900/90">
            1) 브라우저에서{" "}
            <code className="rounded bg-amber-100 px-1">http://127.0.0.1:8010/health</code>{" "}
            를 열어 <code className="rounded bg-amber-100 px-1">parserApiBuild</code>가
            보이는지 확인
            <br />
            2) <code className="rounded bg-amber-100 px-1">frontend/.env.local</code>의{" "}
            <code className="rounded bg-amber-100 px-1">PARSER_API_URL</code>이 그 서버(
            보통 <code className="rounded bg-amber-100 px-1">http://127.0.0.1:8010</code>)
            를 가리키는지 확인 후 Next dev 서버 재시작
            <br />
            3) <code className="rounded bg-amber-100 px-1">PPTX_Parsing/parser-api</code>
            에서 <code className="rounded bg-amber-100 px-1">start-backend.bat</code>로 백엔드
            실행
          </span>
        </p>
      ) : null}

      {hasParsedData ? (
        <div className="mt-8">
          {/* lg: 가운데 미리보기 높이(px)를 ResizeObserver로 재서 양옆 카드에 동일 적용 · JSON은 패널 내부 스크롤 */}
          <div className="grid grid-cols-1 gap-6 lg:grid-cols-[minmax(0,17.5rem)_minmax(0,1fr)_minmax(0,22rem)] lg:items-stretch">
          {/* lg: 메타 고정폭 · 미리보기가 남는 폭 전부 · JSON 고정폭(내부 스크롤). minmax(0,…)로 긴 JSON이 가운데 열을 압축하지 않게 함 */}
          <aside
            className={`order-2 flex min-h-0 w-full min-w-0 flex-col overflow-hidden rounded-xl border border-zinc-200 bg-white p-4 shadow-sm lg:sticky lg:top-4 lg:order-1 ${sideCardHeightClass}`}
            style={sideCardHeightStyle}
          >
            <h2 className="shrink-0 text-sm font-medium text-zinc-800">
              Metadata (draft)
            </h2>
            <div className="mt-4 flex min-h-0 flex-1 flex-col gap-4 overflow-y-auto overscroll-contain pr-0.5">
              <label className="block text-xs text-zinc-500">
                Title (슬라이드 설명 첫 줄에서 자동 반영 · 직접 수정하면 고정)
                <input
                  value={title}
                  onChange={(e) => {
                    titleEditedByUser.current = true;
                    setTitle(e.target.value);
                  }}
                  className="mt-1 w-full rounded border border-zinc-300 px-2 py-1.5 text-sm"
                />
              </label>
              <label className="block text-xs text-zinc-500">
                Tags (comma-separated)
                <input
                  value={tags}
                  onChange={(e) => setTags(e.target.value)}
                  className="mt-1 w-full rounded border border-zinc-300 px-2 py-1.5 text-sm"
                />
              </label>
              {current ? (
                <label className="block text-xs text-zinc-500">
                  슬라이드 설명 (Slide {current.slideNumber})
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
                    placeholder="강사가 이 슬라이드에서 전달할 핵심 메시지를 줄글로 작성하세요."
                    className="mt-1 w-full rounded border border-zinc-300 px-2 py-1.5 text-sm"
                  />
                </label>
              ) : null}
              <div className="block text-xs text-zinc-500">
                Description (자동: 슬라이드 설명 합본)
                <textarea
                  readOnly
                  value={descriptionFromSlides}
                  rows={5}
                  placeholder="슬라이드마다 설명을 입력하면 여기에 순서대로 합쳐집니다."
                  className="mt-1 w-full cursor-default whitespace-pre-wrap rounded border border-dashed border-zinc-300 bg-zinc-50 px-2 py-1.5 text-sm text-zinc-800"
                />
              </div>
              <div className="rounded-lg border border-zinc-200 bg-zinc-50 p-2.5">
                <p className="text-[11px] text-zinc-500">
                  Parsed metadata utility
                </p>
                <div className="mt-2 grid grid-cols-2 gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setTitle(data?.meta?.title ?? "");
                      setTags((data?.meta?.tags ?? []).join(", "));
                      titleEditedByUser.current = true;
                    }}
                    className="rounded border border-zinc-300 bg-white px-2 py-1.5 text-xs text-zinc-700 hover:bg-zinc-50"
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
                    className="rounded border border-zinc-300 bg-white px-2 py-1.5 text-xs text-zinc-700 hover:bg-zinc-50"
                  >
                    입력 초기화
                  </button>
                </div>
                <button
                  type="button"
                  onClick={() => {
                    titleEditedByUser.current = false;
                    const t =
                      deriveTitleFromSlideNotes(slides, slideNotes) ||
                      fallbackTitleFromParser;
                    setTitle(t);
                  }}
                  className="mt-2 w-full rounded border border-zinc-200 bg-white px-2 py-1.5 text-[11px] text-zinc-600 hover:bg-zinc-50"
                >
                  제목을 슬라이드 설명 기준으로 다시 맞추기
                </button>
              </div>
              <p className="text-xs text-zinc-400">
                Final publish, embeddings, and DB insert come in a later step.
              </p>
            </div>
          </aside>

          <div className="order-1 flex h-full min-h-0 w-full min-w-0 items-start lg:order-2">
            {current ? (
              <div
                ref={previewCardRef}
                className="flex w-full min-w-0 flex-col overflow-x-hidden rounded-xl border border-zinc-200/80 bg-gradient-to-b from-white to-zinc-50/80 p-4 shadow-sm"
              >
                <p className="mb-3 shrink-0 text-center text-xs text-zinc-500">
                  슬라이드 {current.slideNumber} — 추출된 도형{" "}
                  {(current.elements ?? []).length}개
                  {(current.plainText ?? "").trim()
                    ? ` · plainText ${(current.plainText ?? "").length}자`
                    : ""}
                </p>
                <div className="w-full min-w-0 shrink-0">
                  <SlideViewer
                    elements={current.elements ?? []}
                    className="mx-auto w-full max-w-full"
                    aspectRatio={slideAspectRatio}
                    canvasPaddingBottomPercent={slideCanvasPaddingBottomPct}
                  />
                </div>
                <nav
                  className="mt-4 flex shrink-0 items-center justify-center gap-2"
                  aria-label="슬라이드 이동"
                >
                  <button
                    type="button"
                    disabled={!canPrevSlide}
                    onClick={() => setSlideIndex((i) => Math.max(0, i - 1))}
                    aria-label="이전 슬라이드"
                    className="inline-flex h-10 w-10 shrink-0 items-center justify-center rounded-full border border-zinc-300 bg-white text-zinc-700 shadow-sm transition hover:border-zinc-400 hover:bg-zinc-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-zinc-400 focus-visible:ring-offset-2 disabled:pointer-events-none disabled:opacity-35"
                  >
                    <span className="sr-only">이전</span>
                    <svg
                      className="h-5 w-5"
                      viewBox="0 0 20 20"
                      fill="currentColor"
                      aria-hidden
                    >
                      <path
                        fillRule="evenodd"
                        d="M12.79 5.23a.75.75 0 01-.02 1.06L8.832 10l3.938 3.71a.75.75 0 11-1.04 1.08l-4.5-4.25a.75.75 0 010-1.08l4.5-4.25a.75.75 0 011.06.02z"
                        clipRule="evenodd"
                      />
                    </svg>
                  </button>
                  <span className="min-w-[5.5rem] select-none text-center text-sm tabular-nums text-zinc-600">
                    <span className="font-medium text-zinc-800">
                      {slideIndex + 1}
                    </span>
                    <span className="text-zinc-400"> / </span>
                    {slides.length}
                  </span>
                  <button
                    type="button"
                    disabled={!canNextSlide}
                    onClick={() =>
                      setSlideIndex((i) =>
                        Math.min(slides.length - 1, i + 1),
                      )
                    }
                    aria-label="다음 슬라이드"
                    className="inline-flex h-10 w-10 shrink-0 items-center justify-center rounded-full border border-zinc-300 bg-white text-zinc-700 shadow-sm transition hover:border-zinc-400 hover:bg-zinc-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-zinc-400 focus-visible:ring-offset-2 disabled:pointer-events-none disabled:opacity-35"
                  >
                    <span className="sr-only">다음</span>
                    <svg
                      className="h-5 w-5"
                      viewBox="0 0 20 20"
                      fill="currentColor"
                      aria-hidden
                    >
                      <path
                        fillRule="evenodd"
                        d="M7.21 14.77a.75.75 0 01.02-1.06L11.168 10 7.23 6.29a.75.75 0 111.04-1.08l4.5 4.25a.75.75 0 010 1.08l-4.5 4.25a.75.75 0 01-1.06-.02z"
                        clipRule="evenodd"
                      />
                    </svg>
                  </button>
                </nav>
                {current.rasterPreview ? (
                  /* block 레이아웃: flex-col 자식 + absolute만 있으면 min-height:auto가 0으로 잡혀 padding 박스가 납작해짐 */
                  <div className="mt-4 min-w-0 shrink-0 space-y-1.5">
                    <p className="text-center text-[11px] text-zinc-500">
                      슬라이드 래스터 미리보기 (PPT→PDF→이미지, JSON 퍼센트 좌표와 동일 비율)
                    </p>
                    <div
                      className="relative isolate h-0 w-full min-w-0 max-w-full overflow-hidden rounded-md border border-zinc-200 bg-black shadow-inner"
                      style={{
                        paddingBottom: `${slideCanvasPaddingBottomPct}%`,
                      }}
                    >
                      {/* eslint-disable-next-line @next/next/no-img-element -- parser data URL; absolute so intrinsic JPEG size cannot expand the box */}
                      <img
                        src={current.rasterPreview}
                        alt=""
                        className="absolute inset-0 box-border h-full w-full max-h-full max-w-full object-contain object-center"
                      />
                    </div>
                  </div>
                ) : (
                  <div className="min-h-0 flex-1" aria-hidden="true" />
                )}
                {slideRasterMeta?.status === "ok" && !current.rasterPreview ? (
                  <p className="mt-2 shrink-0 text-center text-[11px] text-amber-800">
                    이 슬라이드에 rasterPreview가 없습니다. (해당 페이지 렌더 실패 또는 응답
                    불완전 가능)
                  </p>
                ) : null}
                {slideRasterMeta &&
                slideRasterMeta.status &&
                slideRasterMeta.status !== "ok" ? (
                  <p
                    className={`mt-3 shrink-0 rounded border px-2 py-1.5 text-center text-[11px] ${
                      slideRasterMeta.status === "error"
                        ? "border-red-200 bg-red-50 text-red-900"
                        : slideRasterMeta.status === "disabled"
                          ? "border-zinc-200 bg-zinc-50 text-zinc-600"
                          : "border-amber-200 bg-amber-50 text-amber-950"
                    }`}
                  >
                    래스터 {slideRasterMeta.status}:{" "}
                    {slideRasterMeta.reason || "(사유 없음)"}
                    {slideRasterMeta.renderErrorsSample?.length ? (
                      <span className="mt-1 block font-mono text-[10px] opacity-90">
                        {slideRasterMeta.renderErrorsSample.join(" | ")}
                      </span>
                    ) : null}
                  </p>
                ) : null}
                {slideRasterMeta?.status === "ok" &&
                slideRasterMeta.pageCountMismatch ? (
                  <p className="mt-2 shrink-0 text-center text-[11px] text-amber-800">
                    PDF 페이지 수와 슬라이드 수 불일치: PDF{" "}
                    {slideRasterMeta.pageCountMismatch.pdfPages}장 / PPTX{" "}
                    {slideRasterMeta.pageCountMismatch.pptxSlides}장
                  </p>
                ) : null}
              </div>
            ) : null}
          </div>

          <aside
            className={`order-3 flex min-h-0 w-full min-w-0 flex-col overflow-hidden lg:order-3 lg:sticky lg:top-4 ${sideCardHeightClass}`}
            style={sideCardHeightStyle}
          >
            {current ? (
              <div className="flex h-full min-h-0 w-full min-w-0 flex-1 flex-col overflow-hidden rounded-xl border border-zinc-200 bg-zinc-50">
                <div className="shrink-0 border-b border-zinc-200 bg-zinc-50 px-3 py-2 text-sm font-medium text-zinc-800">
                  현재 슬라이드 JSON (미리보기 · data URL 축약)
                </div>
                <pre className="min-h-0 min-w-0 flex-1 overflow-x-auto overflow-y-auto overscroll-contain break-all bg-white p-3 font-mono text-[11px] leading-relaxed text-zinc-800">
                  {slideJsonForPreview(current)}
                </pre>
              </div>
            ) : null}
          </aside>
          </div>
        </div>
      ) : null}
    </main>
  );
}
