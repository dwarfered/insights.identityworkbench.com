'use client';

import * as React from 'react';
import {
  Badge,
  Button,
  Card,
  CardHeader,
  Dropdown,
  InfoLabel,
  Link,
  Option,
  ProgressBar,
  Spinner,
  Text,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  Bar,
  BarChart,
  CartesianGrid,
  LabelList,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
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
    width: '100%',
    minWidth: 0,
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
  metricInfoLabel: {
    display: 'inline-flex',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalXXS,
  },
  statusInfoLabel: {
    display: 'inline-flex',
    alignItems: 'center',
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
  groupLink: {
    fontWeight: 600,
  },
  breakdownList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    width: '100%',
    minWidth: 0,
  },
  breakdownContent: {
    marginTop: tokens.spacingVerticalL,
  },
  breakdownChartWrapper: {
    width: '100%',
    minWidth: 0,
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

const ENTRA_GROUP_URL_BASE =
  'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/';

function getEntraGroupUrl(groupId: string) {
  return `${ENTRA_GROUP_URL_BASE}${encodeURIComponent(groupId)}`;
}

const skuMetricDescriptions = {
  warning:
    'Subscription expired but still in the Microsoft grace window. Units stay active until the customer renews or the state changes to suspended.',
  suspended:
    'Subscription was canceled; units can no longer be assigned but can still be reactivated before Microsoft permanently deletes them.',
  lockedOut:
    'Subscription was canceled and fully locked. Units can’t be reactivated or assigned unless the customer purchases the SKU again.',
};

const capabilityStatusDescriptions: Record<string, string> = {
  Warning:
    'Microsoft flagged this SKU because consumption exceeded your purchased amount. Users stay active only during the grace period.',
  Suspended:
    'Microsoft disabled the SKU for this tenant until billing or compliance issues are resolved, so users cannot access it.',
};

export function LicensingDashboard({ skus }: { skus: SkuUsageModel[] }) {
  const styles = useStyles();
  const [selectedSkuId, setSelectedSkuId] = React.useState<string | null>(null);
  const sortedSkus = React.useMemo(() => {
    return [...skus].sort((a, b) => {
      const labelA = (a.displayName ?? a.skuPartNumber).toLowerCase();
      const labelB = (b.displayName ?? b.skuPartNumber).toLowerCase();
      return labelA.localeCompare(labelB);
    });
  }, [skus]);

  React.useEffect(() => {
    if (!sortedSkus.length) {
      setSelectedSkuId(null);
      return;
    }

    setSelectedSkuId((prev) => {
      if (prev && sortedSkus.some((sku) => sku.skuId === prev)) {
        return prev;
      }

      return null;
    });
  }, [sortedSkus]);

  const selectedSku = selectedSkuId
    ? sortedSkus.find((sku) => sku.skuId === selectedSkuId)
    : undefined;

  return (
    <div className={styles.dashboard}>
      <div className={styles.pickerRow}>
        <Text weight='semibold'>Select a license SKU</Text>
        <Dropdown
          placeholder='Select a license SKU'
          disabled={!sortedSkus.length}
          selectedOptions={selectedSkuId ? [selectedSkuId] : []}
          onOptionSelect={(_, data) => {
            const optionValue = data.optionValue;
            if (typeof optionValue === 'string') {
              setSelectedSkuId(optionValue);
            }
          }}
        >
          {sortedSkus.map((sku) => (
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
  const statusInfo = capabilityStatusDescriptions[sku.capabilityStatus];
  const badge = (
    <Badge appearance={getStatusAppearance(sku.capabilityStatus)}>
      {sku.capabilityStatus}
    </Badge>
  );
  const statusNode = statusInfo ? (
    <InfoLabel info={statusInfo} className={styles.statusInfoLabel}>
      {badge}
    </InfoLabel>
  ) : (
    badge
  );

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
        {statusNode}
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
          <InfoLabel
            info={skuMetricDescriptions.warning}
            className={styles.metricInfoLabel}
          >
            <Text size={200} className={styles.subText}>
              Warning
            </Text>
          </InfoLabel>
          <Text weight='semibold'>{sku.warning}</Text>
        </div>

        <div className={styles.metric}>
          <InfoLabel
            info={skuMetricDescriptions.suspended}
            className={styles.metricInfoLabel}
          >
            <Text size={200} className={styles.subText}>
              Suspended
            </Text>
          </InfoLabel>
          <Text weight='semibold'>{sku.suspended}</Text>
        </div>

        <div className={styles.metric}>
          <InfoLabel
            info={skuMetricDescriptions.lockedOut}
            className={styles.metricInfoLabel}
          >
            <Text size={200} className={styles.subText}>
              Locked out
            </Text>
          </InfoLabel>
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
        description={`${groups.length} group${groups.length === 1 ? '' : 's'} applies this license`}
      />
      <div className={styles.groupList}>
        {groups.map((group) => (
          <div key={group.id} className={styles.groupRow}>
            <div className={styles.groupNameBlock}>
              <Link
                href={getEntraGroupUrl(group.id)}
                target='_blank'
                rel='noreferrer'
                className={styles.groupLink}
              >
                {group.displayName ?? 'Unnamed group'}
              </Link>
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
  const chartData = React.useMemo(
    () =>
      breakdown?.buckets.map((bucket) => ({
        label: bucket.label,
        count: bucket.count,
      })) ?? [],
    [breakdown],
  );

  return (
    <Card>
      <CardHeader
        header={<Text weight='semibold'>Employee type breakdown</Text>}
        description='User counts by employee type for this license.'
      />
      <div className={styles.breakdownContent}>
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
            {chartData.length ? (
              <div className={styles.breakdownChartWrapper}>
                <ResponsiveContainer width='100%' height={320}>
                  <BarChart
                    data={chartData}
                    margin={{ top: 12, right: 16, left: 4, bottom: 32 }}
                  >
                    <CartesianGrid
                      strokeDasharray='3 3'
                      stroke={tokens.colorNeutralStroke2}
                    />
                    <XAxis
                      dataKey='label'
                      interval={0}
                      angle={-20}
                      textAnchor='end'
                      height={60}
                      tick={{ fill: tokens.colorNeutralForeground2 }}
                    />
                    <YAxis
                      allowDecimals={false}
                      tick={{ fill: tokens.colorNeutralForeground2 }}
                    />
                    <Tooltip
                      cursor={{ fill: 'transparent' }}
                      labelStyle={{ color: tokens.colorNeutralForeground1 }}
                      contentStyle={{
                        borderRadius: tokens.borderRadiusMedium,
                        border: `1px solid ${tokens.colorNeutralStroke2}`,
                      }}
                    />
                    <Bar
                      dataKey='count'
                      fill={tokens.colorPaletteBlueBackground2}
                      radius={[6, 6, 0, 0]}
                    >
                      <LabelList
                        dataKey='count'
                        position='top'
                        style={{
                          fill: tokens.colorNeutralForeground1,
                          fontWeight: 600,
                        }}
                      />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            ) : null}
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
      </div>
    </Card>
  );
}
