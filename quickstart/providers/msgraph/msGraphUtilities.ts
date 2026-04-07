import {
  BrowserAuthError,
  InteractionRequiredAuthError,
} from '@azure/msal-browser';
import {
  ensureMsalInitialized,
  loginRequest,
  msalInstance,
} from '../msal/msalAuthConfig';

const globalForMsGraph = globalThis as typeof globalThis & {
  __msalInteractiveTokenPending?: boolean;
};

function isInteractiveTokenPending() {
  return typeof window !== 'undefined'
    ? Boolean(globalForMsGraph.__msalInteractiveTokenPending)
    : false;
}

function setInteractiveTokenPending(value: boolean) {
  if (typeof window === 'undefined') {
    return;
  }
  globalForMsGraph.__msalInteractiveTokenPending = value;
}

export function clearInteractiveTokenRequestFlag() {
  setInteractiveTokenPending(false);
}

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
    const errorCode =
      typeof error === 'object' && error && 'errorCode' in error
        ? String((error as { errorCode?: string }).errorCode ?? '')
        : '';

    const requiresInteractive =
      error instanceof InteractionRequiredAuthError ||
      (error instanceof BrowserAuthError &&
        ['block_iframe_reload', 'token_renewal_timeout', 'monitor_window_timeout', 'timed_out'].includes(
          errorCode,
        ));

    if (requiresInteractive) {
      if (isInteractiveTokenPending()) {
        return new Promise<never>(() => {}) as Promise<string>;
      }
      setInteractiveTokenPending(true);
      console.warn(
        'MSAL: Silent token acquisition failed, requesting interactive token via redirect',
        error,
      );
      msalInstance
        .acquireTokenRedirect(
          account
            ? { ...loginRequest, account }
            : { ...loginRequest },
        )
        .catch((redirectError) => {
          setInteractiveTokenPending(false);
          throw redirectError;
        });
      return new Promise<never>(() => {}) as Promise<string>;
    } else if (process.env.NODE_ENV === 'development') {
      console.error('MSAL: Silent token acquisition failed', error);
    }
    throw error;
  }
}
