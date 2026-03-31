import type { Metadata } from "next";
import { AuthRecoveryHashRedirect } from "@/components/AuthRecoveryHashRedirect";
import "./globals.css";

export const metadata: Metadata = {
  title: "PPTX Lecture Portal",
  description: "Lecture slides portal",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko">
      <body className="min-h-screen bg-zinc-50 text-zinc-900 antialiased">
        <AuthRecoveryHashRedirect />
        {children}
      </body>
    </html>
  );
}
