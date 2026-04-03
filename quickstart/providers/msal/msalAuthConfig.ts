import { appConfig } from '@/config/appConfig';
import {
  Configuration,
  LogLevel,
  PublicClientApplication,
  RedirectRequest,
} from '@azure/msal-browser';

export const msalAuthConfig: Configuration = {
  auth: {
    clientId: appConfig.CLIENT_ID,
    authority: appConfig.AUTHORITY,
    redirectUri: appConfig.REDIRECT_URI,
    postLogoutRedirectUri: appConfig.REDIRECT_URI,
  },

  system: {
    allowPlatformBroker: false,

    loggerOptions: {
      logLevel: LogLevel.Warning,

      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }

        switch (level) {
          case LogLevel.Error:
            console.error(message);
            break;

          case LogLevel.Warning:
            // Suppress common dev-only warning
            if (
              message.includes(
                'There is already an instance of MSAL.js in the window with the same client id',
              )
            ) {
              return;
            }
            console.warn(message);
            break;

          case LogLevel.Info:
            console.info(message);
            break;

          case LogLevel.Verbose:
            console.debug(message);
            break;
        }
      },
    },
  },

  cache: {
    cacheLocation: 'localStorage',
  },
};

export const loginRequest: RedirectRequest = {
  scopes: appConfig.MSGRAPH_DELEGATED_SCOPES,
};

const isBrowser = typeof window !== 'undefined';

const globalForMsal = globalThis as typeof globalThis & {
  __msalInstance?: PublicClientApplication;
};

let msalInstance: PublicClientApplication;
let ensureMsalInitializedImpl: () => Promise<void>;

if (isBrowser) {
  if (!globalForMsal.__msalInstance) {
    if (process.env.NODE_ENV === 'development') {
      console.log('[MSAL] Creating new PublicClientApplication instance');
    }

    globalForMsal.__msalInstance = new PublicClientApplication(msalAuthConfig);
  } else if (process.env.NODE_ENV === 'development') {
    console.log('[MSAL] Reusing existing PublicClientApplication instance');
  }

  msalInstance = globalForMsal.__msalInstance!;

  let msalInitializationPromise: Promise<void> | null = null;
  const originalInitialize = msalInstance.initialize.bind(msalInstance);

  function getMsalInitializationPromise() {
    if (!msalInitializationPromise) {
      msalInitializationPromise = originalInitialize().catch((error) => {
        msalInitializationPromise = null;
        throw error;
      });
    }

    return msalInitializationPromise;
  }

  msalInstance.initialize = getMsalInitializationPromise;
  ensureMsalInitializedImpl = () => getMsalInitializationPromise();
} else {
  const errorMessage =
    'MSAL is only available in the browser environment. Ensure this code runs client-side.';
  msalInstance = new Proxy({} as PublicClientApplication, {
    get() {
      throw new Error(errorMessage);
    },
  });
  ensureMsalInitializedImpl = async () => {
    throw new Error(errorMessage);
  };
}

export { msalInstance };

export function ensureMsalInitialized() {
  return ensureMsalInitializedImpl();
}
