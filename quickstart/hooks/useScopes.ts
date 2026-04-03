'use client';

import { useEffect, useMemo, useState } from 'react';
import { useMsal } from '@azure/msal-react';
import {
  InteractionStatus,
  InteractionRequiredAuthError,
  RedirectRequest,
} from '@azure/msal-browser';

export function useScopes(requiredScopes: string[]) {
  const { instance, accounts, inProgress } = useMsal();

  const [hasScopes, setHasScopes] = useState<boolean | null>(null);
  const [isChecking, setIsChecking] = useState<boolean>(true);
  const [authMessage, setAuthMessage] = useState<string | null>(null);

  const scopeKey = useMemo(
    () => [...requiredScopes].sort().join(' '),
    [requiredScopes],
  );

  useEffect(() => {
    let cancelled = false;

    const checkScopes = async () => {
      if (inProgress !== InteractionStatus.None) {
        return;
      }

      const account = instance.getActiveAccount() ?? accounts[0];

      if (!account) {
        if (!cancelled) {
          setHasScopes(false);
          setIsChecking(false);
          setAuthMessage('No signed-in account.');
        }
        return;
      }

      if (!cancelled) {
        setIsChecking(true);
        setAuthMessage(null);
      }

      try {
        await instance.acquireTokenSilent({
          account,
          scopes: requiredScopes,
        });

        if (!cancelled) {
          setHasScopes(true);
          setAuthMessage(null);
        }
      } catch (error: unknown) {
        if (!cancelled) {
          setHasScopes(false);
        }

        if (error instanceof InteractionRequiredAuthError) {
          if (!cancelled) {
            setAuthMessage(
              'Additional user interaction is required. This may be consent, sign-in again, or MFA depending on tenant session and policy.',
            );
          }
          return;
        }

        if (!cancelled) {
          setAuthMessage(
            error instanceof Error
              ? error.message
              : 'Token acquisition failed.',
          );
        }
      } finally {
        if (!cancelled) {
          setIsChecking(false);
        }
      }
    };

    void checkScopes();

    return () => {
      cancelled = true;
    };
  }, [instance, accounts, inProgress, scopeKey, requiredScopes]);

  const requestConsent = () => {
    const account = instance.getActiveAccount() ?? accounts[0];

    const request: RedirectRequest = {
      scopes: requiredScopes,
      account: account ?? undefined,
      // no prompt by default
    };

    void instance.acquireTokenRedirect(request);
  };

  const forceConsent = () => {
    const account = instance.getActiveAccount() ?? accounts[0];

    const request: RedirectRequest = {
      scopes: requiredScopes,
      account: account ?? undefined,
      prompt: 'consent',
    };

    void instance.acquireTokenRedirect(request);
  };

  return {
    isChecking,
    hasScopes,
    authMessage,
    requestConsent,
    forceConsent,
  };
}
