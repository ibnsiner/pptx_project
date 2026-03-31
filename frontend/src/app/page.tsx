import Link from "next/link";

export default function HomePage() {
  return (
    <main className="mx-auto max-w-2xl px-4 py-16">
      <h1 className="text-2xl font-semibold tracking-tight">
        PPTX Lecture Portal
      </h1>
      <p className="mt-2 text-zinc-600">
        강의 목록과 뷰어는 추후 연결됩니다. 관리자는 업로드·검수부터 사용할 수
        있습니다.
      </p>
      <ul className="mt-8 space-y-2 text-blue-600">
        <li>
          <Link href="/login" className="underline hover:text-blue-800">
            Sign in (admin)
          </Link>
        </li>
        <li>
          <Link href="/admin/upload" className="underline hover:text-blue-800">
            Admin: 업로드 / 프리뷰 (로그인 필요)
          </Link>
        </li>
      </ul>
    </main>
  );
}
