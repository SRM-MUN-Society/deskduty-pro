import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "DeskDuty Pro",
  description:
    "Smart desk duty roster generator for SRM MUN Society — allocate heads and members across hourly slots with fairness-based scheduling.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" className="h-full antialiased">
      <body className="min-h-full flex flex-col bg-mesh">{children}</body>
    </html>
  );
}
