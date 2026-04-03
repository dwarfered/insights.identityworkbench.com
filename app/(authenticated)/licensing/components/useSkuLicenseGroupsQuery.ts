import { msGraphEndpoints } from '@/quickstart/providers/msgraph/msGraphEndpoints';
import { msGraphFetcher } from '@/quickstart/providers/msgraph/msGraphFetcher';
import useSWR from 'swr';

export type GroupWithMemberCount = {
  id: string;
  displayName?: string;
  description?: string;
  memberCount: number | null;
};

type GraphGroup = {
  id: string;
  displayName?: string | null;
  description?: string | null;
};

async function fetchGroupsForSku(
  skuId: string,
): Promise<GroupWithMemberCount[]> {
  const encodedFilter = encodeURIComponent(
    `assignedLicenses/any(x:x/skuId eq ${skuId})`,
  );
  const selectFields = encodeURIComponent('id,displayName,description');
  const url = `${msGraphEndpoints.groups}?$filter=${encodedFilter}&$top=999&$select=${selectFields}`;

  const data = await msGraphFetcher(url, undefined, {
    consistencyLevel: 'eventual',
  });

  const rawGroups: unknown[] = Array.isArray(data.value) ? data.value : [];
  const groups = rawGroups.filter((group): group is GraphGroup => {
    if (typeof group !== 'object' || group === null) {
      return false;
    }
    const maybeGroup = group as Record<string, unknown>;
    return (
      typeof maybeGroup.id === 'string' &&
      (maybeGroup.displayName === undefined ||
        maybeGroup.displayName === null ||
        typeof maybeGroup.displayName === 'string') &&
      (maybeGroup.description === undefined ||
        maybeGroup.description === null ||
        typeof maybeGroup.description === 'string')
    );
  });

  const groupsWithCounts = await Promise.all(
    groups.map(async (group) => {
      try {
        const countResp = await msGraphFetcher(
          `${msGraphEndpoints.groups}/${group.id}/members/$count`,
          undefined,
          { consistencyLevel: 'eventual', paginate: false },
        );
        const parsedCount =
          typeof countResp === 'string'
            ? Number.parseInt(countResp, 10)
            : typeof countResp === 'number'
              ? countResp
              : Number.NaN;
        const memberCount = Number.isNaN(parsedCount) ? null : parsedCount;
        return {
          id: group.id as string,
          displayName: group.displayName as string | undefined,
          description: group.description as string | undefined,
          memberCount,
        } satisfies GroupWithMemberCount;
      } catch (error) {
        console.warn('Unable to load member count for group', group.id, error);
        return {
          id: group.id as string,
          displayName: group.displayName as string | undefined,
          description: group.description as string | undefined,
          memberCount: null,
        } satisfies GroupWithMemberCount;
      }
    }),
  );

  return groupsWithCounts.sort((a, b) => {
    const nameA = (a.displayName ?? '').toLowerCase();
    const nameB = (b.displayName ?? '').toLowerCase();
    if (nameA && nameB) {
      return nameA.localeCompare(nameB);
    }
    if (nameA) {
      return -1;
    }
    if (nameB) {
      return 1;
    }
    return 0;
  });
}

export function useSkuLicenseGroupsQuery(skuId?: string | null) {
  const shouldFetch = Boolean(skuId);
  const { data, error } = useSWR(
    shouldFetch ? ['sku-license-groups', skuId] : null,
    async ([, activeSkuId]) => fetchGroupsForSku(activeSkuId as string),
  );

  return {
    groups: data ?? [],
    error,
    isLoading: shouldFetch && !data && !error,
  };
}
