import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'RELFOOD Reports Engine',
  description: 'RELFOOD Settlement & Remittance Dashboard',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body className="antialiased bg-slate-900 text-slate-100">
        {children}
      </body>
    </html>
  );
}
