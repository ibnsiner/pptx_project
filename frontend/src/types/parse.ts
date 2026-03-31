export type SlideElementStyle = {
  left: string;
  top: string;
  width: string;
  height: string;
  fontSize?: string;
};

export type SlideElement = {
  type: "text" | "image";
  content?: string;
  src?: string;
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
  /** parser-api: ZIP/추출 진단(백엔드 버전 확인용) */
  extractStats?: Record<string, unknown>;
};

export type ParseMeta = {
  title: string;
  description: string;
  tags: string[];
  /** 이 필드가 없으면 Next가 연결한 FastAPI가 이 저장소 parser-api가 아닐 수 있음 */
  parserApiBuild?: string;
};

export type ParsePptxResponse = {
  slides: SlideData[];
  meta: ParseMeta;
  plainText?: string;
};
