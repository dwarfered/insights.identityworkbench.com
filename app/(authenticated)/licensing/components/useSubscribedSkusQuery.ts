import { msGraphEndpoints } from '@/quickstart/providers/msgraph/msGraphEndpoints';
import { msGraphFetcher } from '@/quickstart/providers/msgraph/msGraphFetcher';
import useSWR from 'swr';
import { SubscribedSkusResponse } from './subscribedSkus.types';

export function useSubscribedSkusQuery(enabled = true) {
  const swrKey = enabled ? msGraphEndpoints.subscribedSkus : null;

  const { data, error } = useSWR<SubscribedSkusResponse>(swrKey, msGraphFetcher);

  return {
    skus: data?.value ?? [],
    error,
    isLoading: enabled && !error && !data,
  };
}
