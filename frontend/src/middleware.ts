import { type NextRequest, NextResponse } from "next/server";
import { createServerClient } from "@supabase/ssr";
import { isUserAdmin } from "@/lib/auth/admin";

export async function middleware(request: NextRequest) {
  let supabaseResponse = NextResponse.next({
    request,
  });

  const supabase = createServerClient(
    process.env.NEXT_PUBLIC_SUPABASE_URL!,
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!,
    {
      cookies: {
        getAll() {
          return request.cookies.getAll();
        },
        setAll(
          cookiesToSet: {
            name: string;
            value: string;
            options?: Record<string, unknown>;
          }[],
        ) {
          cookiesToSet.forEach(({ name, value }) =>
            request.cookies.set(name, value),
          );
          supabaseResponse = NextResponse.next({
            request,
          });
          cookiesToSet.forEach(({ name, value, options }) =>
            supabaseResponse.cookies.set(name, value, options),
          );
        },
      },
    },
  );

  const {
    data: { user },
  } = await supabase.auth.getUser();

  if (request.nextUrl.pathname.startsWith("/admin")) {
    const appMeta = user?.app_metadata as Record<string, unknown> | undefined;
    if (!isUserAdmin(user?.email, appMeta)) {
      const url = request.nextUrl.clone();
      url.pathname = "/";
      url.searchParams.set("error", "forbidden");
      return NextResponse.redirect(url);
    }
  }

  return supabaseResponse;
}

/** `/admin`만 검사. `/api/*`는 제외해 프록시·파싱 요청이 미들웨어를 거치지 않게 함. */
export const config = {
  matcher: ["/admin/:path*"],
};
