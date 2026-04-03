'use client';

import useSWR from 'swr';
import { msGraphFetcher } from '../providers/msgraph/msGraphFetcher';

export function useGraphData<T>(
  resource: string,
  swrConfig?: Parameters<typeof useSWR<T>>[2],
) {
  return useSWR<T>(
    resource,
    (url: string) => msGraphFetcher(url) as Promise<T>,
    {
      // SWR defaults (can overwridden by an instance of fetcher)
      revalidateOnFocus: true,
      revalidateOnReconnect: true,
      refreshInterval: 0,
      ...swrConfig,
    },
  );
}
