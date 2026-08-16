import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "DeskDuty Pro",
  description:
    "Smart desk duty roster generator for SRM MUN Society — allocate heads and members across hourly slots with fairness-based scheduling.",
  icons: {
    icon: [
      { url: "/favicon.ico", sizes: "any" },
      { url: "/favicon-16x16.png", sizes: "16x16", type: "image/png" },
      { url: "/favicon-32x32.png", sizes: "32x32", type: "image/png" },
    ],
    apple: "/apple-touch-icon.png",
  },
  manifest: "/site.webmanifest",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" className="h-full antialiased">
      <body className="min-h-full flex flex-col bg-mesh">{children}</body>
    </html>
  );
}
