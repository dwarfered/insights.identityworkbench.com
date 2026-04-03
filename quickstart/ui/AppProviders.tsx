'use client';

import { useEffect, useState } from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { ErrorDialogProvider } from '@/quickstart/ui/common/error-dialog/ErrorDialog';
import MsalClientProvider from '@/quickstart/providers/msal/MsalClientProvider';
import { LoadingScreen } from '@/quickstart/ui/common/LoadingScreen';

export default function AppProviders({
  children,
}: {
  children: React.ReactNode;
}) {
  const [isHydrated, setIsHydrated] = useState(false);

  useEffect(() => {
    const frame = requestAnimationFrame(() => setIsHydrated(true));
    return () => cancelAnimationFrame(frame);
  }, []);

  if (!isHydrated) {
    return <LoadingScreen label='Preparing Fluent UI theme...' />;
  }

  return (
    <MsalClientProvider>
      <FluentProvider theme={webLightTheme}>
        <ErrorDialogProvider>{children}</ErrorDialogProvider>
      </FluentProvider>
    </MsalClientProvider>
  );
}
