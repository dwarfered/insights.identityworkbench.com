import AppProviders from '@/quickstart/ui/AppProviders';
import './globals.css';

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang='en'>
      <body>
        <AppProviders>{children}</AppProviders>
      </body>
    </html>
  );
}
