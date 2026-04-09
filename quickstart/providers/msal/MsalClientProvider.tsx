'use client';

import {
  ReactNode,
  createContext,
  useContext,
  useEffect,
  useMemo,
  useState,
} from 'react';
import { MsalProvider } from '@azure/msal-react';
import {
  ensureMsalInitialized,
  loginRequest,
  msalInstance,
} from './msalAuthConfig';
import { clearInteractiveTokenRequestFlag } from '../msgraph/msGraphUtilities';

async function initializeMsalClient() {
  await ensureMsalInitialized();

  let result = null;
  try {
    result = await msalInstance.handleRedirectPromise();
  } finally {
    clearInteractiveTokenRequestFlag();
  }

  if (result?.account) {
    msalInstance.setActiveAccount(result.account);
    return;
  }

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
  }

  if (!msalInstance.getActiveAccount() && accounts.length > 0) {
    try {
      const ssoResponse = await msalInstance.ssoSilent({
        ...loginRequest,
        loginHint: accounts[0].username,
      });
      msalInstance.setActiveAccount(ssoResponse.account);
    } catch (ssoError) {
      console.warn('Silent SSO failed:', ssoError);
    }
  }
}

type MsalReadinessContextValue = {
  isReady: boolean;
  error: Error | null;
};

const MsalReadinessContext = createContext<MsalReadinessContextValue>({
  isReady: false,
  error: null,
});

export function useMsalReadiness() {
  return useContext(MsalReadinessContext);
}

export default function MsalClientProvider({
  children,
}: {
  children: ReactNode;
}) {
  const [status, setStatus] = useState<'initializing' | 'ready' | 'error'>(
    'initializing',
  );
  const [initError, setInitError] = useState<Error | null>(null);

  useEffect(() => {
    let cancelled = false;

    initializeMsalClient()
      .then(() => {
        if (!cancelled) {
          setStatus('ready');
          setInitError(null);
        }
      })
      .catch((err) => {
        console.error('MSAL initialization failed', err);
        if (!cancelled) {
          setStatus('error');
          setInitError(err instanceof Error ? err : new Error(String(err)));
        }
      });

    return () => {
      cancelled = true;
    };
  }, []);

  const contextValue = useMemo<MsalReadinessContextValue>(
    () => ({
      isReady: status === 'ready',
      error: initError,
    }),
    [status, initError],
  );

  return (
    <MsalReadinessContext.Provider value={contextValue}>
      <MsalProvider instance={msalInstance}>{children}</MsalProvider>
    </MsalReadinessContext.Provider>
  );
}
