"use client";

import { useEffect } from "react";

/**
 * Password reset emails open Site URL with #...&type=recovery.
 * Root "/" has no handler, so redirect to /auth/update-password with the same hash.
 */
export function AuthRecoveryHashRedirect() {
  useEffect(() => {
    if (typeof window === "undefined") return;
    const h = window.location.hash;
    if (!h || h.length < 2) return;
    const type = new URLSearchParams(h.slice(1)).get("type");
    if (type !== "recovery") return;
    if (window.location.pathname.startsWith("/auth/update-password")) return;
    window.location.replace(
      `${window.location.origin}/auth/update-password${h}`,
    );
  }, []);

  return null;
}
