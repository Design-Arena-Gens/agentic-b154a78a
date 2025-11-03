import './globals.css';
import type { Metadata } from 'next';

export const metadata: Metadata = {
  title: 'Logistics Operations Excel Dashboard',
  description: 'Generate an interactive Excel dashboard for logistics reviews',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
