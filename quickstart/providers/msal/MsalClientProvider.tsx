'use client';

import { ReactNode, useEffect } from 'react';
import { MsalProvider } from '@azure/msal-react';
import {
  ensureMsalInitialized,
  loginRequest,
  msalInstance,
} from './msalAuthConfig';

async function initializeMsalClient() {
  await ensureMsalInitialized();

  const result = await msalInstance.handleRedirectPromise();

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

export default function MsalClientProvider({
  children,
}: {
  children: ReactNode;
}) {
  useEffect(() => {
    initializeMsalClient().catch((err) => {
      console.error('MSAL initialization failed', err);
    });
  }, []);

  return <MsalProvider instance={msalInstance}>{children}</MsalProvider>;
}
