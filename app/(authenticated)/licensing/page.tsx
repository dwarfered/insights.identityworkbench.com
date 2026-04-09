'use client';

import * as React from 'react';
import {
  Button,
  Card,
  CardHeader,
  Spinner,
  Text,
  Title3,
} from '@fluentui/react-components';
import { useScopes } from '@/quickstart/hooks/useScopes';
import { useSubscribedSkusQuery } from './components/useSubscribedSkusQuery';
import type { SubscribedSku } from './components/subscribedSkus.types';
import { resolveSkuDisplayName } from './components/skuCatalog.generated';
import {
  LicensingDashboard,
  type SkuUsageModel,
} from './components/LicensingDashboard';
import { useMsalReadiness } from '@/quickstart/providers/msal/MsalClientProvider';

const licenseScopes = [
  'LicenseAssignment.Read.All',
  'Group.Read.All',
  'GroupMember.Read.All',
  'User.Read.All',
];

function mapSubscribedSkuToUsageModel(sku: SubscribedSku): SkuUsageModel {
  const { enabled, warning, suspended, lockedOut } = sku.prepaidUnits;
  const displayName = resolveSkuDisplayName({
    skuPartNumber: sku.skuPartNumber,
    skuId: sku.skuId,
  });

  return {
    skuId: sku.skuId,
    skuPartNumber: sku.skuPartNumber,
    displayName,
    capabilityStatus: sku.capabilityStatus,
    appliesTo: sku.appliesTo,
    enabled,
    consumed: sku.consumedUnits,
    warning,
    suspended,
    lockedOut,
  };
}

export default function LicensingPage() {
  const { isReady: isMsalReady, error: msalError } = useMsalReadiness();
  const { hasScopes, isChecking, requestConsent } = useScopes(licenseScopes);
  const shouldFetchSkus = isMsalReady && hasScopes === true;
  const { skus, error, isLoading } = useSubscribedSkusQuery(shouldFetchSkus);

  const isAwaitingScopesDecision = hasScopes === null;
  const isBusy =
    !isMsalReady || isChecking || isLoading || isAwaitingScopesDecision;

  if (msalError) {
    return (
      <Card>
        <CardHeader
          header={<Title3>Microsoft identity unavailable</Title3>}
          description='MSAL failed to initialize. Refresh the page or sign out and back in.'
        />
        <Text>
          {msalError instanceof Error ? msalError.message : String(msalError)}
        </Text>
      </Card>
    );
  }

  if (isBusy) {
    return (
      <div style={{ display: 'flex', justifyContent: 'center' }}>
        <Spinner label='Loading licensing data...' />
      </div>
    );
  }

  if (hasScopes === false) {
    return (
      <Card>
        <CardHeader
          header={<Title3>Grant license access</Title3>}
          description='Consent to the configured Microsoft Graph scopes to view licensing insights.'
        />
        <Button appearance='primary' onClick={() => requestConsent()}>
          Review permissions
        </Button>
      </Card>
    );
  }

  if (error) {
    return (
      <Card>
        <CardHeader
          header={<Title3>Unable to load licensing data</Title3>}
          description='Check your Graph scopes or try signing out and back in.'
        />
        <Text>{error instanceof Error ? error.message : String(error)}</Text>
      </Card>
    );
  }

  const skuUsageModels = skus.map((sku) => mapSubscribedSkuToUsageModel(sku));

  if (!skuUsageModels.length) {
    return (
      <Card>
        <CardHeader
          header={<Title3>No subscriptions found</Title3>}
          description='Your tenant did not return any subscribed SKUs.'
        />
      </Card>
    );
  }

  return <LicensingDashboard skus={skuUsageModels} />;
}
