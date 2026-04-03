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
    // shape(채워진 도형)는 항상 text/image 아래에 그림
    const typeOrder = (e: SlideElement) => (e.type === "shape" ? 0 : 1);
    const to = typeOrder(a) - typeOrder(b);
    if (to !== 0) return to;
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
  const { ordered, tablePositions } = useMemo(() => {
    const all = sortElementsForPaintOrder(elements ?? []);
    // 표와 같은 위치의 text 요소(평문 중복)를 건너뛰기 위해 위치 Set 구성
    const tablePositions = new Set(
      all
        .filter((el) => el.type === "table")
        .map((el) => `${el.style.left}|${el.style.top}`),
    );
    return { ordered: all, tablePositions };
  }, [elements]);

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
        const isShape = el.type === "shape";
        const s = el.style;
        const posKey = `${s.left}|${s.top}`;

        // 표와 같은 위치의 text 요소는 중복이므로 건너뜀
        if (el.type === "text" && tablePositions.has(posKey)) return null;

        // 표 — HTML table 렌더링
        if (el.type === "table" && el.rows) {
          const hasHeader = el.rows.length > 1;
          return (
            <div
              key={index}
              style={{
                position: "absolute",
                left: s.left,
                top: s.top,
                width: s.width,
                minHeight: s.height,
                zIndex: 10 + index,
                overflow: "hidden",
              }}
            >
              <table
                style={{
                  width: "100%",
                  borderCollapse: "collapse",
                  fontSize: normalizeFontSize(s.fontSize) ?? "1.2cqw",
                  fontFamily: s.fontFamily ?? undefined,
                  tableLayout: "fixed",
                }}
              >
                <tbody>
                  {el.rows.map((row, ri) => (
                    <tr key={ri}>
                      {row.map((cell, ci) => {
                        const isHead = ri === 0 && hasHeader;
                        const Tag = isHead ? "th" : "td";
                        return (
                          <Tag
                            key={ci}
                            style={{
                              border: "1px solid #555",
                              padding: "0.3em 0.5em",
                              textAlign: "center",
                              verticalAlign: "middle",
                              backgroundColor: isHead ? "rgba(0,0,0,0.12)" : "transparent",
                              fontWeight: isHead ? 600 : undefined,
                              wordBreak: "keep-all",
                              overflowWrap: "anywhere",
                            }}
                          >
                            {cell}
                          </Tag>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          );
        }

        // 채워진 도형 / 커넥터 선 / 삼각형 등 — CSS로 표현
        if (isShape) {
          const fillColor = (s as any).fillColor ?? "transparent";
          const fillOpacity = parseFloat((s as any).fillOpacity ?? "1");
          const borderRadius = (s as any).borderRadius ?? "0%";
          const strokeColor = (s as any).strokeColor as string | undefined;
          const clipPath = (s as any).clipPath as string | undefined;
          const isLine = Boolean((s as any).isLine);
          const heightVal = parseFloat(s.height ?? "0");
          const widthVal = parseFloat(s.width ?? "0");
          const effectiveHeight =
            isLine && heightVal < 0.2 ? "0.25%" : s.height;
          // 수직선: width=0% 이면 최소 너비 부여
          const effectiveWidth =
            isLine && widthVal < 0.2 ? "0.25%" : s.width;
          // 흰색 라인은 짙은 배경용이므로 HTML(흰 배경)에서는 연한 회색으로 대체
          const visibleFill =
            fillColor === "#ffffff" || fillColor === "#fff"
              ? "rgba(0,0,0,0.15)"
              : fillColor === "transparent"
                ? "transparent"
                : fillColor;
          return (
            <div
              key={index}
              style={{
                position: "absolute",
                left: s.left,
                top: s.top,
                width: effectiveWidth,
                height: effectiveHeight,
                backgroundColor: visibleFill,
                opacity: fillOpacity,
                // clipPath가 있으면 borderRadius 무시 (삼각형 등)
                borderRadius: clipPath ? undefined : borderRadius,
                clipPath: clipPath ?? undefined,
                border: strokeColor ? `2px solid ${strokeColor}` : undefined,
                zIndex: 5 + index,
                pointerEvents: "none",
              }}
              aria-hidden
            />
          );
        }

        return (
          <div
            key={index}
            className="pointer-events-auto select-text box-border"
            style={{
              position: "absolute",
              left: s.left,
              top: s.top,
              width: s.width,
              ...(isImage
                ? { height: s.height, overflow: "hidden" }
                : { minHeight: s.height, overflow: "visible" }),
              zIndex: 10 + index,
              fontSize: normalizeFontSize(s.fontSize),
              lineHeight: 1.3,
              display: "flex",
              flexDirection: "column",
              justifyContent: "flex-start",
              alignItems:
                s.textAlign === "center"
                  ? "center"
                  : s.textAlign === "right"
                    ? "flex-end"
                    : "flex-start",
              textAlign: s.textAlign ?? "left",
            }}
          >
            {el.type === "text" && el.content ? (
              el.paragraphStyles ? (
                // 단락별 스타일 적용: \n 기준 분리, \u000b는 줄바꿈
                <span className="min-h-0 min-w-0 max-w-full whitespace-pre-wrap break-words" style={{ width: "100%" }}>
                  {normalizeSlideText(el.content).split("\n").map((para, pi) => {
                    const ps = el.paragraphStyles![pi] ?? {};
                    // 단락에 명시된 color가 없으면 첫 단락과 같은 스타일이 아닌 경우
                    // 부모 color를 상속하지 않고 기본 텍스트 색(inherit)을 사용한다.
                    // 이렇게 해야 제목(주황)과 본문(흰색/기본)을 구분할 수 있다.
                    const isFirstStyle = pi === 0;
                    const effectiveColor = ps.color ?? (isFirstStyle ? s.color : undefined);
                    return (
                      <span
                        key={pi}
                        style={{
                          display: "block",
                          color: effectiveColor || undefined,
                          fontWeight: ps.bold != null ? (ps.bold ? 700 : 400) : (isFirstStyle && s.bold ? 700 : undefined),
                          fontStyle: (ps.italic ?? (isFirstStyle ? s.italic : undefined)) ? "italic" : undefined,
                          textDecoration: (ps.underline ?? (isFirstStyle ? s.underline : undefined)) ? "underline" : undefined,
                          fontFamily: ps.fontFamily ?? s.fontFamily ?? undefined,
                          textAlign: ps.textAlign ?? s.textAlign ?? "left",
                          fontSize: ps.fontSize ? normalizeFontSize(ps.fontSize) : (isFirstStyle ? normalizeFontSize(s.fontSize) : undefined),
                        }}
                      >
                        {para || "\u00a0"}
                      </span>
                    );
                  })}
                </span>
              ) : (
                <span
                  className="min-h-0 min-w-0 max-w-full whitespace-pre-wrap break-words"
                  style={{
                    color: s.color ?? undefined,
                    fontWeight: s.bold ? 700 : undefined,
                    fontStyle: s.italic ? "italic" : undefined,
                    textDecoration: s.underline ? "underline" : undefined,
                    fontFamily: s.fontFamily ?? undefined,
                    textAlign: s.textAlign ?? "left",
                    width: "100%",
                  }}
                >
                  {normalizeSlideText(el.content)}
                </span>
              )
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
