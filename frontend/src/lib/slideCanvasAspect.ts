/** PPT 슬라이드 캔버스 높이/너비 비율을 padding-bottom %로 (absolute 자식만 있어도 박스 높이 고정) */
export function slideCanvasPaddingBottomPercent(
  slideSizeEmu?: { width: number; height: number } | null,
): number {
  const w = slideSizeEmu?.width;
  const h = slideSizeEmu?.height;
  if (typeof w === "number" && typeof h === "number" && w > 0 && h > 0) {
    return (h / w) * 100;
  }
  return (9 / 16) * 100;
}
