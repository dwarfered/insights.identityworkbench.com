import {
  AccountInfo,
  AuthenticationResult,
  NavigationClient,
} from '@azure/msal-browser';
import { AppRouterInstance } from 'next/dist/shared/lib/app-router-context.shared-runtime';
import { loginRequest, msalInstance } from './msalAuthConfig';
import { acquireTokenSilentWithFallback } from '../msgraph/msGraphUtilities';

export class CustomNavigationClient extends NavigationClient {
  private router: AppRouterInstance;

  constructor(router: AppRouterInstance) {
    super();
    this.router = router;
  }

  async navigateInternal(url: string, options: { noHistory?: boolean }) {
    const relativePath = url.replace(window.location.origin, '');
    if (options.noHistory) {
      this.router.replace(relativePath);
    } else {
      this.router.push(relativePath);
    }
    return false;
  }
}

export function handleSignIn() {
  msalInstance.loginRedirect(loginRequest).catch((e: unknown) => {
    console.error(`loginRedirect failed: ${e}`);
  });
}

export function handleSignOut() {
  msalInstance.logoutRedirect(loginRequest).catch((e: unknown) => {
    console.error(`loginRedirect failed: ${e}`);
  });
}

// Checks if the current users' access token contains a required scope.
// Used for incremental scope checks (does user have Sites.Read.All etc) or another perm a page may require.
export const hasScopes = async (
  requiredScopes: string[],
  account: AccountInfo | null,
): Promise<boolean> => {
  if (!account) {
    return false;
  }

  try {
    const response = await acquireTokenSilentWithFallback(
      {
        account,
        scopes: requiredScopes,
      },
      { suppressInteractive: true },
    );

    const tokenScopes = response.scopes || [];
    return requiredScopes.every((scope) => tokenScopes.includes(scope));
  } catch {
    return false;
  }
};
