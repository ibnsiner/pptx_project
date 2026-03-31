import { NextRequest, NextResponse } from "next/server";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

function parserBaseUrl(): string {
  const raw =
    process.env.PARSER_API_URL?.trim() ||
    process.env.NEXT_PUBLIC_PARSER_API_URL?.trim() ||
    "http://127.0.0.1:8010";
  return raw.replace(/\/$/, "");
}

/** Proxies to FastAPI so the browser only hits same-origin `/api/parse-pptx` (easy to see in Network). */
export async function POST(request: NextRequest) {
  const base = parserBaseUrl();
  let formData: FormData;
  try {
    formData = await request.formData();
  } catch {
    return NextResponse.json(
      { error: "Invalid multipart body (file too large or malformed)." },
      { status: 400 },
    );
  }

  const file = formData.get("file");
  if (!file || !(file instanceof Blob) || file.size === 0) {
    return NextResponse.json(
      { error: "Missing or empty `file` field in form data." },
      { status: 400 },
    );
  }

  try {
    const upstream = await fetch(`${base}/api/parse-pptx`, {
      method: "POST",
      body: formData,
    });
    const body = await upstream.text();
    const outHeaders = new Headers();
    outHeaders.set(
      "Content-Type",
      upstream.headers.get("content-type") ?? "application/json",
    );
    const build = upstream.headers.get("x-pptx-parser-build");
    if (build) {
      outHeaders.set("X-PPTX-Parser-Build", build);
    }
    return new NextResponse(body, {
      status: upstream.status,
      headers: outHeaders,
    });
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    const hint =
      msg.includes("fetch failed") || msg.includes("ECONNREFUSED")
        ? ` parser-api가 켜져 있는지 확인하세요 (${base}).`
        : "";
    return NextResponse.json(
      {
        error: `Upstream parser unreachable: ${msg}.${hint}`,
      },
      { status: 502 },
    );
  }
}
