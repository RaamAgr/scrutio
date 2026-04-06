import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Scrutio — AI Call Evaluation',
  description: 'Scrutio: Examine every call with precision. Batch AI-powered quality analysis using Gemini.',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
