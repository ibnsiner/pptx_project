"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";
import { createClient } from "@/lib/supabase/client";

export default function LoginPage() {
  const router = useRouter();
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError(null);
    setBusy(true);
    const supabase = createClient();
    const { error: err } = await supabase.auth.signInWithPassword({
      email,
      password,
    });
    setBusy(false);
    if (err) {
      setError(err.message);
      return;
    }
    router.push("/admin/upload");
    router.refresh();
  }

  return (
    <main className="mx-auto max-w-sm px-4 py-16">
      <h1 className="text-lg font-semibold">Sign in</h1>
      <form onSubmit={(e) => void onSubmit(e)} className="mt-6 space-y-4">
        <label className="block text-sm text-zinc-600">
          Email
          <input
            type="email"
            autoComplete="email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            required
            className="mt-1 w-full rounded border border-zinc-300 px-3 py-2 text-zinc-900"
          />
        </label>
        <label className="block text-sm text-zinc-600">
          Password
          <input
            type="password"
            autoComplete="current-password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            required
            className="mt-1 w-full rounded border border-zinc-300 px-3 py-2 text-zinc-900"
          />
        </label>
        {error ? (
          <p className="text-sm text-red-600">{error}</p>
        ) : null}
        <button
          type="submit"
          disabled={busy}
          className="w-full rounded bg-zinc-900 py-2 text-sm font-medium text-white hover:bg-zinc-800 disabled:opacity-50"
        >
          {busy ? "Signing in..." : "Sign in"}
        </button>
      </form>
      <p className="mt-4 text-center text-sm text-zinc-500">
        <a href="/" className="underline">
          Home
        </a>
      </p>
    </main>
  );
}
