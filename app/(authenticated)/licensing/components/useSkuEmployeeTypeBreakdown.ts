import {
  startTransition,
  useCallback,
  useEffect,
  useRef,
  useState,
} from 'react';
import { msGraphEndpoints } from '@/quickstart/providers/msgraph/msGraphEndpoints';
import { msGraphFetcher } from '@/quickstart/providers/msgraph/msGraphFetcher';

// This is an expensive query, not using SWR for auto fetching.

export type EmployeeTypeBreakdown = {
  totalAssigned: number;
  buckets: Array<{ label: string; count: number }>;
};

type ProgressInfo = {
  processed: number;
  total: number | null;
};

type FetchOptions = {
  signal?: AbortSignal;
  onProgress?: (info: ProgressInfo) => void;
};

async function fetchEmployeeTypeBreakdown(
  skuId: string,
  { signal, onProgress }: FetchOptions = {},
): Promise<EmployeeTypeBreakdown> {
  const filter = encodeURIComponent(
    `assignedLicenses/any(x:x/skuId eq ${skuId})`,
  );
  const select = encodeURIComponent('employeeType');
  let url = `${msGraphEndpoints.graphUsers}?$filter=${filter}&$select=${select}&$top=999&$count=true`;

  const counts = new Map<string, number>();
  let processed = 0;
  let totalAssigned: number | null = null;

  while (url) {
    const data = await msGraphFetcher(
      url,
      { signal },
      { consistencyLevel: 'eventual', paginate: false },
    );

    const users = Array.isArray(data.value) ? data.value : [];
    processed += users.length;

    if (totalAssigned === null && typeof data['@odata.count'] === 'number') {
      totalAssigned = data['@odata.count'];
    }

    for (const user of users) {
      const raw =
        typeof user.employeeType === 'string' ? user.employeeType.trim() : '';
      const key = raw.length ? raw : 'Unassigned';
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }

    onProgress?.({ processed, total: totalAssigned });

    const nextLink =
      typeof data['@odata.nextLink'] === 'string'
        ? data['@odata.nextLink']
        : undefined;
    url = nextLink ?? '';
  }

  if (!counts.size) {
    counts.set('Unassigned', 0);
  }

  const buckets = Array.from(counts.entries())
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => {
      if (a.label === 'Unassigned') return -1;
      if (b.label === 'Unassigned') return 1;
      return a.label.localeCompare(b.label);
    });

  return {
    totalAssigned: totalAssigned ?? processed,
    buckets,
  };
}

export function useSkuEmployeeTypeBreakdown(
  skuId?: string,
  enabled: boolean = false,
  requestToken = 0,
) {
  const abortRef = useRef<AbortController | null>(null);
  const [breakdown, setBreakdown] = useState<EmployeeTypeBreakdown | null>(
    null,
  );
  const [error, setError] = useState<Error | null>(null);
  const [status, setStatus] = useState<
    'idle' | 'loading' | 'success' | 'error' | 'cancelled'
  >('idle');
  const [progress, setProgress] = useState<ProgressInfo>({
    processed: 0,
    total: null,
  });
  const shouldRun = Boolean(skuId && enabled);

  const cancel = useCallback(() => {
    abortRef.current?.abort();
  }, []);

  useEffect(() => {
    if (shouldRun) {
      return;
    }

    abortRef.current?.abort();
    startTransition(() => {
      setStatus('idle');
      setBreakdown(null);
      setError(null);
      setProgress({ processed: 0, total: null });
    });
  }, [shouldRun]);

  useEffect(() => {
    if (!shouldRun || !skuId) {
      return;
    }

    abortRef.current?.abort();
    const controller = new AbortController();
    abortRef.current = controller;
    startTransition(() => {
      setStatus('loading');
      setBreakdown(null);
      setError(null);
      setProgress({ processed: 0, total: null });
    });

    fetchEmployeeTypeBreakdown(skuId, {
      signal: controller.signal,
      onProgress: (info) => setProgress(info),
    })
      .then((result) => {
        if (controller.signal.aborted) {
          return;
        }
        setBreakdown(result);
        setStatus('success');
      })
      .catch((err) => {
        if (controller.signal.aborted) {
          setStatus('cancelled');
          return;
        }
        setError(err instanceof Error ? err : new Error(String(err)));
        setStatus('error');
      });

    return () => {
      controller.abort();
    };
  }, [shouldRun, skuId, requestToken]);

  return {
    breakdown,
    error,
    isLoading: status === 'loading',
    isCancelled: status === 'cancelled',
    cancel,
    progress,
  };
}
