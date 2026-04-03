'use client';

import { ReactNode } from 'react';
import { InteractionType } from '@azure/msal-browser';
import {
  MsalAuthenticationResult,
  MsalAuthenticationTemplate,
  type IMsalContext,
} from '@azure/msal-react';
import {
  Body1,
  Button,
  Card,
  CardHeader,
  Title3,
} from '@fluentui/react-components';
import ClientShell from './ClientShell';
import { LoadingScreen } from '@/quickstart/ui/common/LoadingScreen';
import { loginRequest } from '@/quickstart/providers/msal/msalAuthConfig';

const AuthLoading = (context: IMsalContext) => {
  void context;
  return <LoadingScreen label='Checking your Microsoft 365 session...' />;
};

const AuthError = ({ error, login }: MsalAuthenticationResult) => (
  <ClientShell>
    <div
      style={{
        minHeight: '60vh',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      <Card appearance='outline' style={{ maxWidth: 520 }}>
        <CardHeader
          header={<Title3>We could not complete the sign-in</Title3>}
          description='Your session may have expired or the app registration settings need attention.'
        />
        <Body1 style={{ marginBottom: 16 }}>
          {error?.message ?? 'Unknown authentication error.'}
        </Body1>
        <Button
          appearance='primary'
          onClick={() => login(InteractionType.Redirect, loginRequest)}
        >
          Try again
        </Button>
      </Card>
    </div>
  </ClientShell>
);

export default function AuthenticatedClientShell({
  children,
}: {
  children: ReactNode;
}) {
  return (
    <MsalAuthenticationTemplate
      interactionType={InteractionType.Redirect}
      authenticationRequest={loginRequest}
      loadingComponent={AuthLoading}
      errorComponent={AuthError}
    >
      <ClientShell>{children}</ClientShell>
    </MsalAuthenticationTemplate>
  );
}
