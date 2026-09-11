import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Gestão Imobiliária',
  description: 'Gestão de carteira imobiliária — imóveis, movimentos e rentabilidade',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="pt-PT">
      <body className="bg-stone-100 text-stone-900">{children}</body>
    </html>
  );
}
