/**
 * Admin if email is listed in ADMIN_EMAILS (comma-separated) or
 * JWT app_metadata.role === "admin".
 */
export function isAdminEmail(email: string | undefined): boolean {
  if (!email) return false;
  const list = process.env.ADMIN_EMAILS?.split(",").map((s) => s.trim().toLowerCase()) ?? [];
  return list.includes(email.toLowerCase());
}

export function isAdminFromAppMetadata(
  appMetadata: Record<string, unknown> | undefined,
): boolean {
  const role = appMetadata?.role;
  return role === "admin";
}

export function isUserAdmin(
  email: string | undefined,
  appMetadata: Record<string, unknown> | undefined,
): boolean {
  return isAdminFromAppMetadata(appMetadata) || isAdminEmail(email);
}
