import { msGraphEndpoints } from '@/quickstart/providers/msgraph/msGraphEndpoints';
import { msGraphFetcher } from '@/quickstart/providers/msgraph/msGraphFetcher';
import useSWR from 'swr';

type MembersWithLicenseErrorsResponse = {
  value?: unknown;
  '@odata.count'?: unknown;
};

export type LicenseErrorMember = {
  id: string;
  displayName?: string;
  userPrincipalName?: string;
  mail?: string;
};

async function fetchMembersWithLicenseErrors(
  groupId: string,
): Promise<{ members: LicenseErrorMember[]; total: number | null }> {
  const selectFields = encodeURIComponent(
    'id,displayName,userPrincipalName,mail',
  );
  const url = `${msGraphEndpoints.groups}/${groupId}/membersWithLicenseErrors?$top=999&$select=${selectFields}&$count=true`;

  const data = (await msGraphFetcher(url, undefined, {
    consistencyLevel: 'eventual',
  })) as MembersWithLicenseErrorsResponse;

  const rawMembers: unknown[] = Array.isArray(data.value) ? data.value : [];

  const members = rawMembers
    .filter((member): member is Record<string, unknown> =>
      typeof member === 'object' && member !== null,
    )
    .map((member) => {
      const maybeRecord = member as Record<string, unknown>;
      return {
        id: String(maybeRecord.id ?? ''),
        displayName:
          typeof maybeRecord.displayName === 'string'
            ? maybeRecord.displayName
            : undefined,
        userPrincipalName:
          typeof maybeRecord.userPrincipalName === 'string'
            ? maybeRecord.userPrincipalName
            : undefined,
        mail:
          typeof maybeRecord.mail === 'string' ? maybeRecord.mail : undefined,
      } satisfies LicenseErrorMember;
    })
    .filter((member) => member.id.length > 0);

  const total =
    typeof data['@odata.count'] === 'number'
      ? data['@odata.count']
      : members.length;

  return { members, total };
}

export function useGroupLicenseErrorMembersQuery(groupId?: string | null) {
  const shouldFetch = Boolean(groupId);
  const { data, error } = useSWR(
    shouldFetch ? ['group-license-error-members', groupId] : null,
    async ([, activeGroupId]) =>
      fetchMembersWithLicenseErrors(activeGroupId as string),
  );

  return {
    members: data?.members ?? [],
    total: data?.total ?? null,
    error,
    isLoading: shouldFetch && !data && !error,
  };
}
