import { getMsGraphAccessToken } from './msGraphUtilities';

export type MsGraphFetcherOptions = {
  consistencyLevel?: 'eventual';
  /** Whether to follow @odata.nextLink for pagination. Defaults to true. */
  paginate?: boolean;
  /** Safety limit on number of pages when paginate=true. Defaults to Infinity. */
  maxPages?: number;
  /** How many times to retry 429/5xx responses before failing. Defaults to 3. */
  maxRetries?: number;
  /** Initial backoff delay in ms (exponential). Defaults to 500ms. */
  retryDelayMs?: number;
};

export async function msGraphFetcher(
  resource: RequestInfo,
  init?: RequestInit,
  optionsOverride?: MsGraphFetcherOptions,
) {
  let accessToken: string;

  try {
    accessToken = await getMsGraphAccessToken();
  } catch (error) {
    throw error;
  }

  const headers = new Headers(init?.headers);
  headers.set('Authorization', `Bearer ${accessToken}`);
  if (!headers.has('Accept')) {
    headers.set('Accept', 'application/json, text/plain');
  }

  // --- Simulated lag for testing (uncomment to enable) ---
  // await new Promise((res) => setTimeout(res, 2000));
  // --------------------------------------------------------

  if (optionsOverride?.consistencyLevel === 'eventual') {
    headers.set('ConsistencyLevel', 'eventual');
  }

  const options: RequestInit = {
    ...init,
    method: init?.method ?? 'GET',
    headers,
  };

  const maxRetries = optionsOverride?.maxRetries ?? 3;
  const retryDelayMs = optionsOverride?.retryDelayMs ?? 500;

  // First request
  const response = await fetchWithRetry(
    resource,
    options,
    maxRetries,
    retryDelayMs,
  );

  if (!response.ok) {
    const message = await response.text();
    const hint =
      response.status === 429
        ? formatThrottleHint(response.headers.get('Retry-After'))
        : response.status >= 500
          ? ' (Microsoft Graph returned a transient error. Try again shortly or reduce the payload size.)'
          : '';
    throw new Error(`Graph API Error: ${response.status} - ${message}${hint}`);
  }

  const contentType = response.headers.get('Content-Type');

  if (contentType?.includes('application/json')) {
    const data = await response.json();

    const paginate =
      optionsOverride?.paginate !== undefined ? optionsOverride.paginate : true;
    const maxPages =
      optionsOverride?.maxPages !== undefined
        ? optionsOverride.maxPages
        : Infinity;

    // If not paginating or no nextLink, return immediately
    if (!paginate || !data['@odata.nextLink']) {
      return data;
    }

    // Ensure data.value is an array to aggregate into
    if (!Array.isArray(data.value)) {
      data.value = data.value ? [data.value] : [];
    }

    // Follow @odata.nextLink iteratively
    let nextLink: string | undefined = data['@odata.nextLink'];
    let pageCount = 1;

    while (paginate && nextLink && pageCount < maxPages) {
      const pageResp = await fetchWithRetry(
        nextLink,
        options,
        maxRetries,
        retryDelayMs,
      );
      if (!pageResp.ok) {
        const msg = await pageResp.text();
        throw new Error(
          `Graph API Error (pagination): ${pageResp.status} - ${msg}`,
        );
      }

      const pageContentType = pageResp.headers.get('Content-Type');
      if (!pageContentType?.includes('application/json')) {
        throw new Error(
          `Unsupported content type during pagination: ${pageContentType}`,
        );
      }

      const pageData = await pageResp.json();

      // Merge page values
      const pageValues = Array.isArray(pageData.value)
        ? pageData.value
        : pageData.value
          ? [pageData.value]
          : [];

      data.value = [...data.value, ...pageValues];

      // Advance
      nextLink =
        typeof pageData['@odata.nextLink'] === 'string'
          ? pageData['@odata.nextLink']
          : undefined;

      pageCount += 1;
    }

    return data;
  } else if (contentType?.includes('blob') || contentType?.includes('image')) {
    return response.blob();
  } else if (contentType?.startsWith('text/')) {
    return response.text();
  }

  throw new Error(`Unsupported content type: ${contentType}`);
}

function shouldRetryStatus(status: number) {
  return status === 408 || status === 429 || (status >= 500 && status < 600);
}

function parseRetryAfterMs(header: string | null): number | null {
  if (!header) {
    return null;
  }

  const seconds = Number(header);
  if (!Number.isNaN(seconds)) {
    return Math.max(0, seconds) * 1000;
  }

  const dateTarget = Date.parse(header);
  if (!Number.isNaN(dateTarget)) {
    const delta = dateTarget - Date.now();
    return delta > 0 ? delta : 0;
  }

  return null;
}

function formatThrottleHint(header: string | null) {
  const retryAfterMs = parseRetryAfterMs(header);
  if (retryAfterMs === null) {
    return ' (Microsoft Graph throttled the request. Wait before retrying or lower the request frequency.)';
  }

  const seconds = Math.max(1, Math.ceil(retryAfterMs / 1000));
  return ` (Microsoft Graph throttled the request. Wait about ${seconds} seconds before retrying.)`;
}

async function fetchWithRetry(
  resource: RequestInfo,
  options: RequestInit,
  maxRetries: number,
  baseDelayMs: number,
) {
  let attempt = 0;

  while (true) {
    const response = await fetch(resource, options);

    if (response.ok || !shouldRetryStatus(response.status)) {
      return response;
    }

    if (attempt >= maxRetries) {
      return response;
    }

    const retryAfterMs = parseRetryAfterMs(response.headers.get('Retry-After'));
    const delay =
      retryAfterMs !== null ? retryAfterMs : baseDelayMs * 2 ** attempt;
    await wait(delay);
    attempt += 1;
  }
}

function wait(durationMs: number) {
  return new Promise((resolve) => setTimeout(resolve, durationMs));
}
