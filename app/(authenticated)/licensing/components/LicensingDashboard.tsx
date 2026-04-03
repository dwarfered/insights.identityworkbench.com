'use client';

import * as React from 'react';
import {
  Badge,
  Button,
  Card,
  CardHeader,
  Dropdown,
  Option,
  ProgressBar,
  Spinner,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { useSkuLicenseGroupsQuery } from './useSkuLicenseGroupsQuery';
import { useSkuEmployeeTypeBreakdown } from './useSkuEmployeeTypeBreakdown';

export type SkuUsageModel = {
  skuId: string;
  skuPartNumber: string;
  displayName?: string;
  capabilityStatus: 'Enabled' | 'Suspended' | 'Warning' | string;
  appliesTo: 'User' | 'Company' | string;
  enabled: number;
  consumed: number;
  warning: number;
  suspended: number;
  lockedOut: number;
};

const useStyles = makeStyles({
  root: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(340px, 1fr))',
    gap: tokens.spacingHorizontalXL,
    alignItems: 'start',
  },
  dashboard: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
  },
  pickerRow: {
    display: 'flex',
    flexWrap: 'wrap',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
  },
  card: {
    borderRadius: tokens.borderRadiusXLarge,
    padding: tokens.spacingHorizontalL,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  headerRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalS,
  },
  statBlock: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
  },
  bigStat: {
    fontSize: '32px',
    lineHeight: 1,
  },
  subText: {
    color: tokens.colorNeutralForeground3,
  },
  metaGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    gap: tokens.spacingHorizontalM,
    rowGap: tokens.spacingVerticalS,
  },
  metric: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
  },
  groupList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  groupRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalS}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
  },
  groupNameBlock: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
  },
  groupMeta: {
    color: tokens.colorNeutralForeground3,
  },
  breakdownList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  insightsRow: {
    display: 'grid',
    gap: tokens.spacingHorizontalXL,
    gridTemplateColumns: 'minmax(0, 1fr)',
    '@media (min-width: 1024px)': {
      gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    },
  },
});

function getAvailable(sku: SkuUsageModel) {
  return Math.max(
    sku.enabled - sku.consumed - sku.warning - sku.suspended - sku.lockedOut,
    0,
  );
}

function getPercentConsumed(sku: SkuUsageModel) {
  return sku.enabled > 0 ? sku.consumed / sku.enabled : 0;
}

function getStatusAppearance(
  status: string,
): 'filled' | 'tint' | 'outline' | 'ghost' {
  switch (status) {
    case 'Enabled':
      return 'filled';
    case 'Warning':
      return 'tint';
    default:
      return 'outline';
  }
}

export function LicensingDashboard({ skus }: { skus: SkuUsageModel[] }) {
  const styles = useStyles();
  const [selectedSkuId, setSelectedSkuId] = React.useState<string | null>(null);

  React.useEffect(() => {
    if (!skus.length) {
      setSelectedSkuId(null);
      return;
    }

    setSelectedSkuId((prev) => {
      if (prev && skus.some((sku) => sku.skuId === prev)) {
        return prev;
      }

      return null;
    });
  }, [skus]);

  const selectedSku = selectedSkuId
    ? skus.find((sku) => sku.skuId === selectedSkuId)
    : undefined;

  return (
    <div className={styles.dashboard}>
      <div className={styles.pickerRow}>
        <Text weight='semibold'>Select a license SKU</Text>
        <Dropdown
          placeholder='Select a license SKU'
          disabled={!skus.length}
          selectedOptions={selectedSkuId ? [selectedSkuId] : []}
          onOptionSelect={(_, data) => {
            const optionValue = data.optionValue;
            if (typeof optionValue === 'string') {
              setSelectedSkuId(optionValue);
            }
          }}
        >
          {skus.map((sku) => (
            <Option key={sku.skuId} value={sku.skuId}>
              {sku.displayName
                ? `${sku.displayName} (${sku.skuPartNumber})`
                : sku.skuPartNumber}
            </Option>
          ))}
        </Dropdown>
      </div>

      {selectedSku ? (
        <>
          <div className={styles.root}>
            <SkuUsageCard sku={selectedSku} />
          </div>
          <div className={styles.insightsRow}>
            <SkuLicenseGroupsCard skuId={selectedSku.skuId} />
            <SkuEmployeeTypeBreakdownCard skuId={selectedSku.skuId} />
          </div>
        </>
      ) : (
        <Card>
          <CardHeader
            header={
              <Text weight='semibold'>Choose a license to view details</Text>
            }
            description='Pick from the dropdown to see assignment stats.'
          />
        </Card>
      )}
    </div>
  );
}

function SkuUsageCard({ sku }: { sku: SkuUsageModel }) {
  const styles = useStyles();
  const available = getAvailable(sku);
  const percentConsumed = getPercentConsumed(sku);

  return (
    <Card className={styles.card}>
      <div className={styles.headerRow}>
        <CardHeader
          header={
            <Text weight='semibold'>
              {sku.displayName ?? sku.skuPartNumber}
            </Text>
          }
          description={
            <Text size={200} className={styles.subText}>
              {sku.appliesTo}
            </Text>
          }
        />
        <Badge appearance={getStatusAppearance(sku.capabilityStatus)}>
          {sku.capabilityStatus}
        </Badge>
      </div>

      <div className={styles.statBlock}>
        <Text className={styles.bigStat}>
          {sku.consumed} / {sku.enabled}
        </Text>
        <Text size={300} className={styles.subText}>
          Assigned licenses
        </Text>
      </div>

      <ProgressBar value={percentConsumed} />

      <div className={styles.metaGrid}>
        <div className={styles.metric}>
          <Text size={200} className={styles.subText}>
            Available
          </Text>
          <Text weight='semibold'>{available}</Text>
        </div>

        <div className={styles.metric}>
          <Text size={200} className={styles.subText}>
            Warning
          </Text>
          <Text weight='semibold'>{sku.warning}</Text>
        </div>

        <div className={styles.metric}>
          <Text size={200} className={styles.subText}>
            Suspended
          </Text>
          <Text weight='semibold'>{sku.suspended}</Text>
        </div>

        <div className={styles.metric}>
          <Text size={200} className={styles.subText}>
            Locked out
          </Text>
          <Text weight='semibold'>{sku.lockedOut}</Text>
        </div>
      </div>
    </Card>
  );
}

function SkuLicenseGroupsCard({ skuId }: { skuId: string }) {
  const styles = useStyles();
  const { groups, isLoading, error } = useSkuLicenseGroupsQuery(skuId);

  if (isLoading) {
    return (
      <Card>
        <CardHeader
          header={<Text weight='semibold'>Groups assigning this license</Text>}
          description='Loading group assignments from Microsoft Entra...'
        />
        <Spinner label='Loading groups...' />
      </Card>
    );
  }

  if (error) {
    return (
      <Card>
        <CardHeader
          header={<Text weight='semibold'>Unable to load license groups</Text>}
          description='Check whether your account has permission to read group assignments.'
        />
        <Text>{error instanceof Error ? error.message : String(error)}</Text>
      </Card>
    );
  }

  if (!groups.length) {
    return (
      <Card>
        <CardHeader
          header={<Text weight='semibold'>Groups assigning this license</Text>}
          description='No Entra groups currently assign this SKU.'
        />
      </Card>
    );
  }

  return (
    <Card>
      <CardHeader
        header={<Text weight='semibold'>Groups assigning this license</Text>}
        description={`${groups.length} group${groups.length === 1 ? '' : 's'} apply this SKU`}
      />
      <div className={styles.groupList}>
        {groups.map((group) => (
          <div key={group.id} className={styles.groupRow}>
            <div className={styles.groupNameBlock}>
              <Text weight='semibold'>
                {group.displayName ?? 'Unnamed group'}
              </Text>
              {group.description ? (
                <Text size={200} className={styles.groupMeta}>
                  {group.description}
                </Text>
              ) : null}
            </div>
            <Text weight='semibold'>
              {group.memberCount !== null
                ? `${group.memberCount} members`
                : 'Unknown count'}
            </Text>
          </div>
        ))}
      </div>
    </Card>
  );
}

function SkuEmployeeTypeBreakdownCard({ skuId }: { skuId: string }) {
  const styles = useStyles();
  const [requestId, setRequestId] = React.useState(0);
  React.useEffect(() => {
    setRequestId(0);
  }, [skuId]);

  const { breakdown, isLoading, error, isCancelled, cancel, progress } =
    useSkuEmployeeTypeBreakdown(skuId, requestId > 0, requestId);

  return (
    <Card>
      <CardHeader
        header={<Text weight='semibold'>Load employee type breakdown</Text>}
        description='User counts by employee type for this license.'
      />
      {requestId === 0 ? (
        <Button
          appearance='primary'
          onClick={() => setRequestId((id) => id + 1)}
        >
          Request breakdown
        </Button>
      ) : isLoading ? (
        <div className={styles.breakdownList}>
          <Spinner label='Calculating employee type breakdown… May take time for large tenants.' />
          <Text size={200} className={styles.groupMeta}>
            Processed {progress.processed.toLocaleString()} user
            {progress.processed === 1 ? '' : 's'}
            {progress.total !== null
              ? ` of ${progress.total.toLocaleString()} assigned`
              : ''}
          </Text>
          <Button appearance='secondary' onClick={() => cancel()}>
            Cancel request
          </Button>
        </div>
      ) : error ? (
        <>
          <Text>
            {error instanceof Error
              ? error.message
              : 'Unable to load breakdown'}
          </Text>
          <Button onClick={() => setRequestId((id) => id + 1)}>
            Try again
          </Button>
        </>
      ) : isCancelled ? (
        <>
          <Text size={200} className={styles.groupMeta}>
            Cancelled after processing {progress.processed.toLocaleString()}{' '}
            user
            {progress.processed === 1 ? '' : 's'}.
          </Text>
          <Button
            appearance='secondary'
            onClick={() => setRequestId((id) => id + 1)}
          >
            Restart breakdown
          </Button>
        </>
      ) : breakdown ? (
        <div className={styles.breakdownList}>
          {/* <Text size={200} className={styles.groupMeta}>
            Total licensed users: {breakdown.totalAssigned}
          </Text> */}
          {breakdown.buckets.map((bucket) => (
            <div key={bucket.label} className={styles.groupRow}>
              <Text weight='semibold'>{bucket.label}</Text>
              <Text weight='semibold'>{bucket.count}</Text>
            </div>
          ))}
          <Button
            appearance='secondary'
            onClick={() => setRequestId((id) => id + 1)}
          >
            Refresh breakdown
          </Button>
        </div>
      ) : null}
    </Card>
  );
}
