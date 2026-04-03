const DEFAULT_APP_NAME =
  'Next.js MSAL React SPA Quickstart (Microsoft Entra ID + Microsoft Graph)';
const DEFAULT_REDIRECT_URI = 'http://localhost:3000';

const requireEnv = (
  value: string | undefined,
  name:
    | 'NEXT_PUBLIC_TENANT_ID'
    | 'NEXT_PUBLIC_CLIENT_ID'
    | 'NEXT_PUBLIC_MSAL_SCOPES',
) => {
  const trimmed = value?.trim();

  if (!trimmed) {
    throw new Error(
      `[appConfig] Missing required environment variable ${name}. See README.md for setup details.`,
    );
  }

  return trimmed;
};

const TENANT_ID = requireEnv(
  process.env.NEXT_PUBLIC_TENANT_ID,
  'NEXT_PUBLIC_TENANT_ID',
);
const CLIENT_ID = requireEnv(
  process.env.NEXT_PUBLIC_CLIENT_ID,
  'NEXT_PUBLIC_CLIENT_ID',
);

const parsedScopes = requireEnv(
  process.env.NEXT_PUBLIC_MSAL_SCOPES,
  'NEXT_PUBLIC_MSAL_SCOPES',
)
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

if (parsedScopes.length === 0) {
  throw new Error(
    '[appConfig] NEXT_PUBLIC_MSAL_SCOPES must include at least one Microsoft Graph delegated permission.',
  );
}

export const appConfig = {
  APP_NAME: process.env.NEXT_PUBLIC_APP_NAME?.trim() || DEFAULT_APP_NAME,

  CLIENT_ID,

  TENANT_ID,

  AUTHORITY: `https://login.microsoftonline.com/${TENANT_ID}`,

  REDIRECT_URI:
    process.env.NEXT_PUBLIC_MSAL_REDIRECT_URI?.trim() || DEFAULT_REDIRECT_URI,

  MSGRAPH_DELEGATED_SCOPES: parsedScopes,
};
