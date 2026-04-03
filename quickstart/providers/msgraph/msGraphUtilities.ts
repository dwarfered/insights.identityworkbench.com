import { InteractionRequiredAuthError } from '@azure/msal-browser';
import {
  ensureMsalInitialized,
  loginRequest,
  msalInstance,
} from '../msal/msalAuthConfig';

export class NoActiveAccountError extends Error {
  constructor() {
    super('No active account found.');
    this.name = 'NoActiveAccountError';
  }
}

export async function getMsGraphAccessToken(): Promise<string> {
  // Ensure MSAL is fully initialized before using any API on the instance
  await ensureMsalInitialized();

  let account = msalInstance.getActiveAccount();

  if (!account) {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      account = accounts[0];
      msalInstance.setActiveAccount(account);
    } else {
      throw new NoActiveAccountError();
    }
  }

  try {
    const response = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account,
    });
    return response.accessToken;
  } catch (error: unknown) {
    if (error instanceof InteractionRequiredAuthError) {
      console.warn('MSAL: Interaction required, redirecting to sign in');
      msalInstance.loginRedirect(loginRequest);
    } else {
      if (process.env.NODE_ENV === 'development') {
        console.error('MSAL: Silent token acquisition failed', error);
      }
    }
    throw error;
  }
}
