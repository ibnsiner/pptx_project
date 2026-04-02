export type SlideElementStyle = {
  left: string;
  top: string;
  width: string;
  height: string;
  // text
  fontSize?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  textAlign?: "left" | "center" | "right" | "justify";
  // shape (filled shape / connector line)
  fillColor?: string;
  fillOpacity?: string;
  borderRadius?: string;
  strokeColor?: string;
  isLine?: boolean;
  clipPath?: string;
};

export type ParagraphStyle = Partial<Pick<SlideElementStyle,
  "color" | "bold" | "italic" | "underline" | "fontSize" | "fontFamily" | "textAlign"
>>;

export type SlideElement = {
  type: "text" | "image" | "shape" | "table";
  content?: string;
  src?: string;
  /** 단락별 스타일 배열. \n으로 구분된 단락과 1:1 대응, 빈 단락은 {} */
  paragraphStyles?: ParagraphStyle[];
  /** 표 데이터: rows[행][열] = 셀 텍스트 */
  rows?: string[][];
  /** parser-api: 인라인 한도 초과 등 */
  skipReason?: string;
  byteLength?: number;
  contentType?: string;
  style: SlideElementStyle;
};

export type SlideData = {
  slideNumber: number;
  elements: SlideElement[];
  plainText?: string;
  /** LibreOffice+PyMuPDF JPEG data URL (승인 전 미리보기·AI 시각 근거) */
  rasterPreview?: string;
  /** parser-api: ZIP/추출 진단(백엔드 버전 확인용) */
  extractStats?: Record<string, unknown>;
};

export type SlideRasterMeta = {
  enabled?: boolean;
  status?: string;
  reason?: string;
  /** 래스터 엔진: "powerpoint-com" | "libreoffice" */
  engine?: string;
  /** PowerPoint COM 실패 후 LibreOffice fallback 시 PPT COM 실패 사유 */
  pptComFallbackReason?: string;
  longEdgePx?: number;
  jpegQuality?: number;
  slidesRendered?: number;
  pdfPageCount?: number;
  pageCountMismatch?: { pdfPages: number; pptxSlides: number };
  renderErrorsSample?: string[];
  /** PPTX 파일에서 추출한 폰트 목록 */
  pptxFonts?: string[];
  /** 시스템에 설치되지 않은 폰트 (래스터 텍스트 누락 주요 원인) */
  missingFonts?: string[];
  missingFontHint?: string;
};

export type ParseMeta = {
  title: string;
  description: string;
  tags: string[];
  /** 이 필드가 없으면 Next가 연결한 FastAPI가 이 저장소 parser-api가 아닐 수 있음 */
  parserApiBuild?: string;
  /** python-pptx 슬라이드 크기(EMU), 래스터와 동일 종횡비 정합 */
  slideSizeEmu?: { width: number; height: number };
  slideRaster?: SlideRasterMeta;
};

export type ParsePptxResponse = {
  slides: SlideData[];
  meta: ParseMeta;
  plainText?: string;
};
