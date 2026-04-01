"use client";

import { useMemo, useState, type CSSProperties } from "react";
import type { SlideElement } from "@/types/parse";

function parsePercent(s: string): number {
  const m = /^([\d.]+)%$/.exec(String(s).trim());
  return m ? parseFloat(m[1]) : 0;
}

/** PPT 줄바꿈(수직 탭 U+000B 등)을 화면에서 읽히게 */
function normalizeSlideText(content: string): string {
  return content
    .replace(/\u000b/g, "\n")
    .replace(/\u000c/g, "\n")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n");
}

/**
 * 아래쪽 도형을 먼저 그리고, 위쪽(y 작음)을 나중에 그려 겹칠 때 제목이 위에 오게 함.
 */
/** 구 파서의 vw 단위는 뷰포트 기준이라 미리보기 상자보다 글자가 과하게 커짐 → cqw로 통일 */
function normalizeFontSize(raw: string | undefined): string {
  if (!raw || !raw.trim()) {
    return "1.15cqw";
  }
  const s = raw.trim();
  const m = /^([\d.]+)\s*vw$/i.exec(s);
  if (m) {
    return `${m[1]}cqw`;
  }
  return s;
}

function sortElementsForPaintOrder(elements: SlideElement[]): SlideElement[] {
  return [...elements].sort((a, b) => {
    const tb = parsePercent(b.style.top);
    const ta = parsePercent(a.style.top);
    if (tb !== ta) return tb - ta;
    return parsePercent(a.style.left) - parsePercent(b.style.left);
  });
}

function SlideRaster({ src }: { src: string }) {
  const [failed, setFailed] = useState(false);
  if (failed) {
    return (
      <span className="text-xs leading-tight text-red-600">
        이미지 로드 실패(URL·용량·CSP 확인)
      </span>
    );
  }
  return (
    // eslint-disable-next-line @next/next/no-img-element
    <img
      src={src}
      alt=""
      className="h-full w-full object-contain"
      draggable={false}
      onError={() => setFailed(true)}
    />
  );
}

type Props = {
  elements?: SlideElement[] | null;
  className?: string;
  /** 슬라이드 캔버스 종횡비 (CSS aspect-ratio). canvasPaddingBottomPercent 없을 때만 사용 */
  aspectRatio?: string;
  /**
   * (slideHeight/slideWidth)*100 — padding-bottom % 트릭으로 캔버스 높이 고정.
   * 래스터 미리보기와 픽셀 단위로 동일한 방식으로 맞출 때 전달.
   */
  canvasPaddingBottomPercent?: number;
};

export function SlideViewer({
  elements,
  className = "",
  aspectRatio = "16 / 9",
  canvasPaddingBottomPercent,
}: Props) {
  const ordered = useMemo(
    () => sortElementsForPaintOrder(elements ?? []),
    [elements],
  );

  const canvasBoxStyle =
    canvasPaddingBottomPercent != null
      ? ({
          height: 0,
          paddingBottom: `${canvasPaddingBottomPercent}%`,
          containerType: "inline-size" as const,
          overflow: "hidden" as const,
        } satisfies CSSProperties)
      : ({
          aspectRatio,
          containerType: "inline-size" as const,
          overflow: "hidden" as const,
        } satisfies CSSProperties);

  const h0WhenPb =
    canvasPaddingBottomPercent != null ? "h-0 " : "";

  if (!elements?.length) {
    return (
      <div
        className={`relative ${h0WhenPb}min-w-0 w-full max-w-full rounded-md border border-dashed border-zinc-300 bg-zinc-50 text-sm text-zinc-500 ${className}`}
        style={canvasBoxStyle}
      >
        <div className="absolute inset-0 flex items-center justify-center px-2 text-center">
          이 슬라이드에서 추출된 도형이 없습니다. (텍스트만 다른 슬라이드에 있거나 파서 한계일 수
          있습니다.)
        </div>
      </div>
    );
  }

  return (
    <div
      className={`relative ${h0WhenPb}min-w-0 w-full max-w-full rounded-md border border-zinc-200 bg-white shadow-sm ${className}`}
      style={canvasBoxStyle}
    >
      {ordered.map((el, index) => {
        const isImage = el.type === "image";
        return (
          <div
            key={index}
            className="pointer-events-auto select-text box-border min-h-0"
            style={{
              position: "absolute",
              left: el.style.left,
              top: el.style.top,
              width: el.style.width,
              /* PPT 도형 박스 높이에 맞춤. auto+minHeight면 본문이 박스 밖으로 무한 확장됨 */
              height: el.style.height,
              zIndex: 10 + index,
              fontSize: normalizeFontSize(el.style.fontSize),
              lineHeight: 1.25,
              display: "flex",
              flexDirection: "column",
              justifyContent: "flex-start",
              alignItems: "flex-start",
              overflow: "hidden",
            }}
          >
            {el.type === "text" && el.content ? (
              <span className="min-h-0 min-w-0 max-w-full whitespace-pre-wrap break-words text-left text-zinc-900">
                {normalizeSlideText(el.content)}
              </span>
            ) : null}
            {isImage && el.src ? <SlideRaster src={el.src} /> : null}
            {isImage && !el.src && el.skipReason ? (
              <span className="text-xs leading-tight text-amber-800">
                이미지 미표시: {el.skipReason}
                {typeof el.byteLength === "number"
                  ? ` (${Math.round(el.byteLength / 1024)}KB)`
                  : ""}
              </span>
            ) : null}
          </div>
        );
      })}
    </div>
  );
}
