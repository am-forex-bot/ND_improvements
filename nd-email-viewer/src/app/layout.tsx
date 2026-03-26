import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "ND Email Viewer — NetDocuments Done Right",
  description: "A better front-end for NetDocuments email and document management",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className="antialiased">{children}</body>
    </html>
  );
}
