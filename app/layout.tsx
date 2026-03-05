import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Vital Score Card Builder',
  description: 'Generate Perigee field visit score card reports for Vital reps.',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body className="min-h-screen bg-[#f8f8f8] text-[#32373C]">{children}</body>
    </html>
  );
}
