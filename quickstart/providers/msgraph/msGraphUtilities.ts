import {
  BrowserAuthError,
  InteractionRequiredAuthError,
  type AuthenticationResult,
  type SilentRequest,
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

type TokenAcquisitionOptions = {
  suppressInteractive?: boolean;
};

export async function acquireTokenSilentWithFallback(
  requestOverrides: Partial<SilentRequest> = {},
  options: TokenAcquisitionOptions = {},
): Promise<AuthenticationResult> {
  await ensureMsalInitialized();

  let account = requestOverrides.account ?? msalInstance.getActiveAccount();

  if (!account) {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      account = accounts[0];
      msalInstance.setActiveAccount(account);
    } else {
      throw new NoActiveAccountError();
    }
  }

  const silentRequest: SilentRequest = {
    ...loginRequest,
    ...requestOverrides,
    account,
  };

  try {
    return await msalInstance.acquireTokenSilent(silentRequest);
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
      if (options.suppressInteractive) {
        throw error instanceof Error ? error : new Error(String(error));
      }
      if (isInteractiveTokenPending()) {
        return new Promise<never>(() => {}) as Promise<AuthenticationResult>;
      }
      setInteractiveTokenPending(true);
      console.warn(
        'MSAL: Silent token acquisition failed, requesting interactive token via redirect',
        error,
      );
      try {
        await msalInstance.acquireTokenRedirect(
          account
            ? { ...loginRequest, ...requestOverrides, account }
            : { ...loginRequest, ...requestOverrides },
        );
      } catch (redirectError) {
        setInteractiveTokenPending(false);
        throw redirectError;
      }
      return new Promise<never>(() => {}) as Promise<AuthenticationResult>;
    } else if (process.env.NODE_ENV === 'development') {
      console.error('MSAL: Silent token acquisition failed', error);
    }
    throw error;
  }
}

export async function getMsGraphAccessToken(): Promise<string> {
  const response = await acquireTokenSilentWithFallback();
  return response.accessToken;
}
