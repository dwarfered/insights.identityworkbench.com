import { msGraphEndpoints } from '@/quickstart/providers/msgraph/msGraphEndpoints';
import { msGraphFetcher } from '@/quickstart/providers/msgraph/msGraphFetcher';
import useSWR from 'swr';

export type LicenseAssignmentState = {
  skuId: string;
  assignedByGroup?: string;
  state?: string;
  error?: string;
  lastUpdatedDateTime?: string;
};

export type UserLicenseAssignmentStates = {
  id: string;
  displayName?: string;
  userPrincipalName?: string;
  licenseAssignmentStates: LicenseAssignmentState[];
};

async function fetchUserLicenseAssignmentStates(
  userId: string,
): Promise<UserLicenseAssignmentStates> {
  const selectFields = encodeURIComponent(
    'id,displayName,userPrincipalName,licenseAssignmentStates',
  );
  const url = `${msGraphEndpoints.graphUsers}/${userId}?$select=${selectFields}`;

  const data = await msGraphFetcher(url);

  if (typeof data !== 'object' || data === null) {
    throw new Error('Unexpected Graph response for license states');
  }

  const record = data as Record<string, unknown>;
  const licenseAssignmentStates = Array.isArray(
    record.licenseAssignmentStates,
  )
    ? (record.licenseAssignmentStates.filter((state): state is Record<string, unknown> =>
        typeof state === 'object' && state !== null,
      ) as Record<string, unknown>[])
    : [];

  return {
    id: typeof record.id === 'string' ? record.id : userId,
    displayName:
      typeof record.displayName === 'string' ? record.displayName : undefined,
    userPrincipalName:
      typeof record.userPrincipalName === 'string'
        ? record.userPrincipalName
        : undefined,
    licenseAssignmentStates: licenseAssignmentStates.map((state) => ({
      skuId: typeof state.skuId === 'string' ? state.skuId : '',
      assignedByGroup:
        typeof state.assignedByGroup === 'string'
          ? state.assignedByGroup
          : undefined,
      state: typeof state.state === 'string' ? state.state : undefined,
      error: typeof state.error === 'string' ? state.error : undefined,
      lastUpdatedDateTime:
        typeof state.lastUpdatedDateTime === 'string'
          ? state.lastUpdatedDateTime
          : undefined,
    })),
  } satisfies UserLicenseAssignmentStates;
}

export function useUserLicenseAssignmentStatesQuery(
  userId?: string | null,
) {
  const shouldFetch = Boolean(userId);
  const { data, error } = useSWR(
    shouldFetch ? ['user-license-assignment-states', userId] : null,
    async ([, activeUserId]) =>
      fetchUserLicenseAssignmentStates(activeUserId as string),
  );

  return {
    assignmentStates: data?.licenseAssignmentStates ?? [],
    user: data,
    error,
    isLoading: shouldFetch && !data && !error,
  };
}
