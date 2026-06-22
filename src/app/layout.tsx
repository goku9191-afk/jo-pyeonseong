import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "조편성 프로그램",
  description: "학급 조편성을 쉽고 빠르게!",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko" className="h-full">
      <body className="min-h-full flex flex-col">{children}</body>
    </html>
  );
}
