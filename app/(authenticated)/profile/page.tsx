'use client';

import {
  Body1,
  Body1Strong,
  Button,
  Card,
  CardHeader,
  Spinner,
  Title3,
} from '@fluentui/react-components';
import { appConfig } from '@/config/appConfig';
import { useScopes } from '@/quickstart/hooks/useScopes';
import { useGraphData } from '@/quickstart/hooks/useGraphData';
import { msGraphEndpoints } from '@/quickstart/providers/msgraph/msGraphEndpoints';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { handleSignIn } from '@/quickstart/providers/msal/msalUtilities';

type GraphMeResponse = {
  displayName?: string;
  givenName?: string;
  surname?: string;
  jobTitle?: string;
  mail?: string;
  userPrincipalName?: string;
  businessPhones?: string[];
  officeLocation?: string;
};

export default function ProfilePage() {
  const { hasScopes, isChecking, requestConsent } = useScopes(
    appConfig.MSGRAPH_DELEGATED_SCOPES,
  );

  const { data, error, isLoading } = useGraphData<GraphMeResponse>(
    msGraphEndpoints.graphMe,
  );

  const awaitingData =
    !error && !data && hasScopes !== false && !isChecking && !isLoading;

  if (isChecking || isLoading || hasScopes === null || awaitingData) {
    return (
      <div style={{ display: 'flex', justifyContent: 'center' }}>
        <Spinner label='Loading your profile...' />
      </div>
    );
  }

  const cardStyle = { maxWidth: '720px' };

  if (hasScopes === false) {
    return (
      <Card appearance='outline' style={cardStyle}>
        <CardHeader
          header={<Title3>Grant profile access</Title3>}
          description='Consent to the scopes configured for this quickstart to view your profile.'
        />
        <Button appearance='primary' onClick={() => requestConsent()}>
          Review permissions
        </Button>
      </Card>
    );
  }

  if (error) {
    return (
      <Card appearance='outline' style={cardStyle}>
        <CardHeader
          header={
            error instanceof InteractionRequiredAuthError
              ? 'Session expired'
              : 'Unable to load profile'
          }
          description={
            error instanceof InteractionRequiredAuthError
              ? 'We need you to sign in again so we can acquire a fresh Microsoft Graph token.'
              : 'Check the Graph scopes or try signing out and back in.'
          }
        />
        <Body1 style={{ marginBottom: 16 }}>
          {error instanceof Error ? error.message : String(error)}
        </Body1>
        {error instanceof InteractionRequiredAuthError ? (
          <Button appearance='primary' onClick={() => handleSignIn()}>
            Sign in again
          </Button>
        ) : null}
      </Card>
    );
  }

  if (!data) {
    return null;
  }

  const businessPhone =
    data.businessPhones && data.businessPhones.length > 0
      ? data.businessPhones[0]
      : 'Not provided';

  return (
    <Card style={cardStyle}>
      <CardHeader
        header={<Title3>{data.displayName ?? 'Unknown user'}</Title3>}
        description={data.jobTitle ?? 'No job title on file'}
      />
      <Body1Strong>Given name</Body1Strong>
      <Body1>{data.givenName ?? 'Not provided'}</Body1>

      <Body1Strong style={{ marginTop: '12px' }}>Surname</Body1Strong>
      <Body1>{data.surname ?? 'Not provided'}</Body1>

      <Body1Strong style={{ marginTop: '12px' }}>Email</Body1Strong>
      <Body1>{data.mail ?? data.userPrincipalName ?? 'Not provided'}</Body1>

      <Body1Strong style={{ marginTop: '12px' }}>
        User Principal Name
      </Body1Strong>
      <Body1>{data.userPrincipalName}</Body1>

      <Body1Strong style={{ marginTop: '12px' }}>Business phone</Body1Strong>
      <Body1>{businessPhone}</Body1>

      <Body1Strong style={{ marginTop: '12px' }}>Office location</Body1Strong>
      <Body1>{data.officeLocation ?? 'Not provided'}</Body1>
    </Card>
  );
}
