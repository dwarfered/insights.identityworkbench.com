'use client';

import * as React from 'react';
import {
  Badge,
  Button,
  Card,
  CardHeader,
  Combobox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  InfoLabel,
  Link,
  Option,
  ProgressBar,
  Spinner,
  Text,
  makeStyles,
  mergeClasses,
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
import {
  useSkuLicenseGroupsQuery,
  type GroupWithMemberCount,
} from './useSkuLicenseGroupsQuery';
import { useSkuEmployeeTypeBreakdown } from './useSkuEmployeeTypeBreakdown';
import { useGroupLicenseErrorMembersQuery } from './useGroupLicenseErrorMembersQuery';
import { useUserLicenseAssignmentStatesQuery } from './useUserLicenseAssignmentStatesQuery';

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
  licensePicker: {
    minWidth: '360px',
    flex: '1 1 520px',
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
    alignItems: 'center',
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalS}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
  },
  groupRowActions: {
    display: 'flex',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalS,
    rowGap: tokens.spacingVerticalXXS,
    flexWrap: 'wrap',
    justifyContent: 'flex-start',
  },
  errorBadge: {
    whiteSpace: 'nowrap',
  },
  groupNameBlock: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
  },
  groupCount: {
    color: tokens.colorNeutralForeground2,
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
  errorDialogContent: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  errorDialogSplit: {
    display: 'grid',
    gap: tokens.spacingHorizontalL,
    gridTemplateColumns: 'minmax(0, 1fr)',
    '@media (min-width: 768px)': {
      gridTemplateColumns: 'minmax(0, 240px) minmax(0, 1fr)',
    },
  },
  errorList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
    maxHeight: '360px',
    overflowY: 'auto',
    paddingRight: tokens.spacingHorizontalS,
  },
  errorUserButton: {
    textAlign: 'left',
    width: '100%',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    padding: tokens.spacingHorizontalS,
    cursor: 'pointer',
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
    transitionProperty: 'background, border, box-shadow',
    transitionDuration: tokens.durationUltraFast,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      border: `1px solid ${tokens.colorNeutralStrokeAccessible}`,
    },
  },
  errorUserButtonActive: {
    border: `1px solid ${tokens.colorPaletteRedBorder2}`,
    boxShadow: `0 0 0 1px ${tokens.colorPaletteRedBorder2}`,
    backgroundColor: tokens.colorPaletteRedBackground1,
  },
  errorUserMeta: {
    color: tokens.colorNeutralForeground3,
  },
  errorDetailsPanel: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  assignmentStateList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  assignmentStateRow: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalS,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
  },
  errorEmptyState: {
    color: tokens.colorNeutralForeground3,
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
const ENTRA_USER_URL_BASE =
  'https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/';

function getEntraGroupUrl(groupId: string) {
  return `${ENTRA_GROUP_URL_BASE}${encodeURIComponent(groupId)}`;
}

function getEntraUserUrl(userId: string) {
  return `${ENTRA_USER_URL_BASE}${encodeURIComponent(userId)}`;
}

function formatSkuLabel(sku: SkuUsageModel) {
  return sku.displayName
    ? `${sku.displayName} (${sku.skuPartNumber})`
    : sku.skuPartNumber;
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

const licenseAssignmentErrorDescriptions: Record<string, string> = {
  CountViolation:
    'The tenant has assigned more units than it purchased. Remove a seat, increase the license count, or reprocess the user assignment. In some cases, the assignment may remain in error until it is manually reprocessed even if capacity becomes available.',

  DependencyViolation:
    'This license requires another license or service plan that the user does not have.',

  CoexistenceViolation:
    'This license conflicts with another assigned license. Remove the conflicting license to continue.',

  ProhibitedInRegion:
    'This license is not available in the users location. Update the usage location or remove the assignment.',

  ConsumerSubscription:
    'This license is for personal use and cannot be assigned in this tenant.',

  ServicePlanConflict:
    'A service within this license conflicts with another assigned license.',

  PendingInput: 'License assignment is still processing. Check again shortly.',

  UniquenessViolation:
    'A required attribute (such as UPN or email) conflicts with another account.',
};

export function LicensingDashboard({ skus }: { skus: SkuUsageModel[] }) {
  const styles = useStyles();
  const [selectedSkuId, setSelectedSkuId] = React.useState<string | null>(null);
  const [filterText, setFilterText] = React.useState('');
  const [hasTypedFilter, setHasTypedFilter] = React.useState(false);
  const sortedSkus = React.useMemo(() => {
    return [...skus].sort((a, b) => {
      const labelA = (a.displayName ?? a.skuPartNumber).toLowerCase();
      const labelB = (b.displayName ?? b.skuPartNumber).toLowerCase();
      return labelA.localeCompare(labelB);
    });
  }, [skus]);
  const filteredSkus = React.useMemo(() => {
    const query = filterText.trim().toLowerCase();
    if (!query) {
      return sortedSkus;
    }
    return sortedSkus.filter((sku) => {
      const label = formatSkuLabel(sku).toLowerCase();
      return (
        label.includes(query) || sku.skuPartNumber.toLowerCase().includes(query)
      );
    });
  }, [sortedSkus, filterText]);

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

  const displaySkus = hasTypedFilter ? filteredSkus : sortedSkus;
  const selectedSku = selectedSkuId
    ? sortedSkus.find((sku) => sku.skuId === selectedSkuId)
    : undefined;
  React.useEffect(() => {
    if (selectedSku) {
      setFilterText(formatSkuLabel(selectedSku));
      setHasTypedFilter(false);
    }
  }, [selectedSku]);
  const handleComboboxFocus = React.useCallback(() => {
    setFilterText('');
    setHasTypedFilter(false);
  }, []);
  const handleComboboxPointerDown = React.useCallback(
    (event: React.PointerEvent<HTMLDivElement>) => {
      if (
        event.target instanceof HTMLInputElement &&
        document.activeElement === event.target
      ) {
        setFilterText('');
        setHasTypedFilter(false);
      }
    },
    [],
  );

  return (
    <div className={styles.dashboard}>
      <div className={styles.pickerRow}>
        <Text weight='semibold'>Select a license SKU</Text>
        <Combobox
          placeholder='Search licenses (e.g., F3)'
          disabled={!sortedSkus.length}
          selectedOptions={selectedSkuId ? [selectedSkuId] : []}
          value={filterText}
          className={styles.licensePicker}
          onFocus={handleComboboxFocus}
          onPointerDown={handleComboboxPointerDown}
          onInput={(event, data?: { value?: string }) => {
            const nextValue =
              data?.value ??
              (event?.currentTarget instanceof HTMLInputElement
                ? event.currentTarget.value
                : '');
            setFilterText(nextValue);
            setHasTypedFilter(nextValue.trim().length > 0);
          }}
          onOpenChange={(_, data?: { open: boolean }) => {
            if (data?.open) {
              setHasTypedFilter(false);
            }
          }}
          onOptionSelect={(_, data) => {
            const optionValue = data.optionValue;
            if (typeof optionValue === 'string') {
              setSelectedSkuId(optionValue);
              const match = sortedSkus.find((sku) => sku.skuId === optionValue);
              if (match) {
                setFilterText(formatSkuLabel(match));
              }
              setHasTypedFilter(false);
            }
          }}
        >
          {displaySkus.map((sku) => (
            <Option key={sku.skuId} value={sku.skuId}>
              {formatSkuLabel(sku)}
            </Option>
          ))}
        </Combobox>
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
  const totalGroups = groups.length;
  const groupNameLookup = React.useMemo(() => {
    return groups.reduce<Record<string, string>>((acc, group) => {
      if (group.id) {
        acc[group.id.toLowerCase()] = group.displayName ?? '';
      }
      return acc;
    }, {});
  }, [groups]);

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
        description={`${totalGroups} group${totalGroups === 1 ? '' : 's'} applies this license`}
      />
      <div className={styles.groupList}>
        {groups.map((group) => (
          <LicenseGroupRow
            key={group.id}
            group={group}
            skuId={skuId}
            groupNameLookup={groupNameLookup}
          />
        ))}
      </div>
    </Card>
  );
}

function LicenseGroupRow({
  group,
  skuId,
  groupNameLookup,
}: {
  group: GroupWithMemberCount;
  skuId: string;
  groupNameLookup: Record<string, string>;
}) {
  const styles = useStyles();
  return (
    <div className={styles.groupRow}>
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
        <Text size={200} className={styles.groupCount}>
          {group.memberCount !== null
            ? `${group.memberCount.toLocaleString()} member${group.memberCount === 1 ? '' : 's'}`
            : 'Member count unavailable'}
        </Text>
        <GroupLicenseErrorInsights
          groupId={group.id}
          groupName={group.displayName}
          skuId={skuId}
          groupNameLookup={groupNameLookup}
        />
      </div>
    </div>
  );
}

function GroupLicenseErrorInsights({
  groupId,
  groupName,
  skuId,
  groupNameLookup,
}: {
  groupId: string;
  groupName?: string;
  skuId: string;
  groupNameLookup: Record<string, string>;
}) {
  const styles = useStyles();
  const [dialogOpen, setDialogOpen] = React.useState(false);
  const [selectedUserId, setSelectedUserId] = React.useState<string | null>(
    null,
  );
  const { members, total, error, isLoading } =
    useGroupLicenseErrorMembersQuery(groupId);
  const sortedMembers = React.useMemo(() => {
    return [...members].sort((a, b) => {
      const nameA = (a.displayName ?? a.userPrincipalName ?? '').toLowerCase();
      const nameB = (b.displayName ?? b.userPrincipalName ?? '').toLowerCase();
      return nameA.localeCompare(nameB);
    });
  }, [members]);

  const issueCount = total ?? sortedMembers.length;
  const hasIssues = issueCount > 0;

  React.useEffect(() => {
    if (!dialogOpen) {
      setSelectedUserId(null);
      return;
    }
    if (!sortedMembers.length) {
      setSelectedUserId(null);
      return;
    }
    setSelectedUserId((current) => {
      if (current && sortedMembers.some((member) => member.id === current)) {
        return current;
      }
      return sortedMembers[0]?.id ?? null;
    });
  }, [dialogOpen, sortedMembers]);

  const closeDialog = () => {
    setDialogOpen(false);
    setSelectedUserId(null);
  };

  const selectedUser = sortedMembers.find((m) => m.id === selectedUserId);

  return (
    <>
      <div className={styles.groupRowActions}>
        {isLoading ? (
          <Spinner
            aria-label='Checking for license errors'
            size='extra-small'
          />
        ) : (
          <>
            <Badge
              appearance={hasIssues ? 'filled' : 'outline'}
              color={hasIssues ? 'danger' : 'brand'}
              className={styles.errorBadge}
            >
              {hasIssues
                ? `${issueCount.toLocaleString()} issue${issueCount === 1 ? '' : 's'}`
                : 'No license errors'}
            </Badge>
            {hasIssues ? (
              <Button
                appearance='secondary'
                size='small'
                onClick={() => setDialogOpen(true)}
              >
                View affected users
              </Button>
            ) : null}
          </>
        )}
      </div>
      <Dialog
        open={dialogOpen}
        onOpenChange={(_, data) => {
          if (!data.open) {
            closeDialog();
          }
        }}
      >
        <DialogSurface>
          <DialogBody>
            <DialogTitle>
              License issues in {groupName ?? 'this group'}
            </DialogTitle>
            <DialogContent>
              <div className={styles.errorDialogContent}>
                {isLoading ? (
                  <Spinner label='Loading issue details…' />
                ) : error ? (
                  <Text role='alert'>
                    {error instanceof Error ? error.message : String(error)}
                  </Text>
                ) : !sortedMembers.length ? (
                  <Text className={styles.errorEmptyState}>
                    Microsoft Graph did not return any users with assignment
                    errors for this group.
                  </Text>
                ) : (
                  <>
                    <Text weight='semibold'>
                      {issueCount.toLocaleString()} user
                      {issueCount === 1 ? '' : 's'} with assignment issues
                    </Text>
                    <div className={styles.errorDialogSplit}>
                      <div className={styles.errorList}>
                        {sortedMembers.map((member) => (
                          <button
                            key={member.id}
                            type='button'
                            className={mergeClasses(
                              styles.errorUserButton,
                              member.id === selectedUserId
                                ? styles.errorUserButtonActive
                                : undefined,
                            )}
                            onClick={() => setSelectedUserId(member.id)}
                          >
                            <Text weight='semibold'>
                              {member.displayName ??
                                member.userPrincipalName ??
                                'Unknown user'}
                            </Text>
                            {member.userPrincipalName ? (
                              <Text size={200} className={styles.errorUserMeta}>
                                {member.userPrincipalName}
                              </Text>
                            ) : null}
                          </button>
                        ))}
                      </div>
                      <div className={styles.errorDetailsPanel}>
                        {selectedUser ? (
                          <>
                            <Link
                              href={getEntraUserUrl(selectedUser.id)}
                              target='_blank'
                              rel='noreferrer'
                              className={styles.groupLink}
                            >
                              {selectedUser.displayName ??
                                selectedUser.userPrincipalName ??
                                'Selected user'}
                            </Link>
                            <UserLicenseAssignmentStatePanel
                              userId={selectedUser.id}
                              skuId={skuId}
                              groupNameLookup={groupNameLookup}
                            />
                          </>
                        ) : (
                          <Text className={styles.errorEmptyState}>
                            Select a user to view their assignment details.
                          </Text>
                        )}
                      </div>
                    </div>
                  </>
                )}
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance='secondary' onClick={closeDialog}>
                Close
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
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

function UserLicenseAssignmentStatePanel({
  userId,
  skuId,
  groupNameLookup,
}: {
  userId: string;
  skuId: string;
  groupNameLookup: Record<string, string>;
}) {
  const styles = useStyles();
  const { assignmentStates, error, isLoading } =
    useUserLicenseAssignmentStatesQuery(userId);
  const normalizedSkuId = skuId.toLowerCase();
  const filteredStates = React.useMemo(
    () =>
      assignmentStates.filter(
        (state) =>
          typeof state.skuId === 'string' &&
          state.skuId.toLowerCase() === normalizedSkuId,
      ),
    [assignmentStates, normalizedSkuId],
  );

  if (isLoading) {
    return <Spinner label='Loading assignment states…' />;
  }

  if (error) {
    return (
      <Text role='alert'>
        {error instanceof Error ? error.message : String(error)}
      </Text>
    );
  }

  if (!filteredStates.length) {
    return (
      <Text className={styles.errorEmptyState}>
        Microsoft Graph did not return licenseAssignmentStates for this SKU.
      </Text>
    );
  }

  return (
    <div className={styles.assignmentStateList}>
      {filteredStates.map((state, index) => (
        <div
          key={`${state.skuId}-${state.assignedByGroup ?? 'direct'}-${index}`}
          className={styles.assignmentStateRow}
        >
          <Text weight='semibold'>State: {state.state ?? 'Unknown state'}</Text>
          {state.error ? (
            <>
              <Text size={200} className={styles.groupMeta}>
                Error: {state.error}
              </Text>
              {getLicenseErrorDescription(state.error) ? (
                <Text size={200} className={styles.errorEmptyState}>
                  {getLicenseErrorDescription(state.error)}
                </Text>
              ) : null}
            </>
          ) : null}
          <Text size={200} className={styles.groupMeta}>
            Assigned by:{' '}
            {formatAssignedBy(state.assignedByGroup, groupNameLookup)}
          </Text>
          {state.lastUpdatedDateTime ? (
            <Text size={200} className={styles.groupMeta}>
              Updated {formatDate(state.lastUpdatedDateTime)}
            </Text>
          ) : null}
        </div>
      ))}
    </div>
  );
}

function formatDate(value: string) {
  const parsed = Date.parse(value);
  if (Number.isNaN(parsed)) {
    return value;
  }
  return new Date(parsed).toLocaleString();
}

function getLicenseErrorDescription(code?: string) {
  if (!code) {
    return undefined;
  }
  return licenseAssignmentErrorDescriptions[code] ?? undefined;
}

function formatAssignedBy(
  groupId: string | undefined,
  groupNameLookup: Record<string, string>,
) {
  if (!groupId) {
    return 'Direct or unknown assignment';
  }
  const lookupKey = groupId.toLowerCase();
  const groupName = groupNameLookup[lookupKey];
  if (groupName && groupName.length) {
    return `${groupName} (${groupId})`;
  }
  return groupId;
}
