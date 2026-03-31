"use client";

import { useEffect, useState } from "react";
import { useRouter } from "next/navigation";
import { createClient } from "@/lib/supabase/client";

export default function UpdatePasswordPage() {
  const router = useRouter();
  const [ready, setReady] = useState(false);
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    const supabase = createClient();

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((event) => {
      if (event === "PASSWORD_RECOVERY") {
        setReady(true);
      }
    });

    void supabase.auth.getSession().then(({ data: { session } }) => {
      if (typeof window === "undefined") return;
      const hash = window.location.hash;
      if (hash.includes("type=recovery") && session) {
        setReady(true);
      }
    });

    return () => {
      subscription.unsubscribe();
    };
  }, []);

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError(null);
    if (password.length < 8) {
      setError("Password must be at least 8 characters.");
      return;
    }
    if (password !== confirm) {
      setError("Passwords do not match.");
      return;
    }
    setBusy(true);
    const supabase = createClient();
    const { error: err } = await supabase.auth.updateUser({ password });
    setBusy(false);
    if (err) {
      setError(err.message);
      return;
    }
    router.push("/login");
    router.refresh();
  }

  return (
    <main className="mx-auto max-w-sm px-4 py-16">
      <h1 className="text-lg font-semibold">Set new password</h1>
      <p className="mt-2 text-sm text-zinc-600">
        Enter a new password for your account.
      </p>

      {!ready ? (
        <p className="mt-6 text-sm text-zinc-500">Loading session...</p>
      ) : (
        <form onSubmit={(e) => void onSubmit(e)} className="mt-6 space-y-4">
          <label className="block text-sm text-zinc-600">
            New password
            <input
              type="password"
              autoComplete="new-password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              required
              minLength={8}
              className="mt-1 w-full rounded border border-zinc-300 px-3 py-2 text-zinc-900"
            />
          </label>
          <label className="block text-sm text-zinc-600">
            Confirm
            <input
              type="password"
              autoComplete="new-password"
              value={confirm}
              onChange={(e) => setConfirm(e.target.value)}
              required
              minLength={8}
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
            {busy ? "Saving..." : "Update password"}
          </button>
        </form>
      )}

      <p className="mt-6 text-center text-sm text-zinc-500">
        <a href="/login" className="underline">
          Back to sign in
        </a>
      </p>
    </main>
  );
}
