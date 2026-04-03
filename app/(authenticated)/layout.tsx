import type { ReactNode } from 'react';
import AuthenticatedClientShell from '@/quickstart/ui/layouts/AuthenticatedClientShell';

export default function AuthenticatedLayout({
  children,
}: {
  children: ReactNode;
}) {
  return <AuthenticatedClientShell>{children}</AuthenticatedClientShell>;
}
