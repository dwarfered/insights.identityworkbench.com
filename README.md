# Next.js MSAL React SPA Template (Microsoft Entra ID + Microsoft Graph)

> A production-ready Next.js React SPA template for Microsoft Entra ID authentication using MSAL and Microsoft Graph.

**Live demo:** https://spa.identityworkbench.com/  
**Built for:** real-world identity-enabled applications, not just samples

It provides a clean, working foundation to sign users in, acquire delegated tokens, and immediately start rendering Microsoft Graph data using modern React patterns (hooks, SWR, Fluent UI).

Built to reflect how identity-enabled applications are actually implemented, not just how they are documented.

[![Next.js](https://img.shields.io/badge/Next.js-App%20Router-blue)](https://github.com/vercel/next.js)
[![Auth](https://img.shields.io/badge/Auth-MS%20Entra%20ID%20(MSAL)-blueviolet)](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-react)
[![Microsoft Graph](https://img.shields.io/badge/API-Microsoft%20Graph-blue)](https://github.com/microsoftgraph/msgraph-sdk-javascript)
[![Fluent UI](https://img.shields.io/badge/UI-Fluent%20UI-ff69b4)](https://github.com/microsoft/fluentui)
[![License](https://img.shields.io/github/license/dwarfered/nextjs-msal-react-graph-spa-template)](...)
[![Last Commit](https://img.shields.io/github/last-commit/dwarfered/nextjs-msal-react-graph-spa-template)](...)

## ⭐ Support

If this project saves you time or helps you get started faster, consider giving it a star.

[![GitHub stars](https://img.shields.io/github/stars/dwarfered/nextjs-msal-react-graph-spa-template.svg?style=social)](https://github.com/dwarfered/nextjs-msal-react-graph-spa-template/stargazers)

## Live Demo

[https://spa.identityworkbench.com/](https://spa.identityworkbench.com/)

## Author

[www.linkedin.com/in/chris-dymond](https://www.linkedin.com/in/chris-dymond)

## Author’s Note

This template brings together patterns, lessons, and real-world experience working with MSAL, React, and Microsoft Graph into a single, practical starting point.

Instead of piecing together multiple samples and docs, the goal is to provide a clean foundation that just works.

If it saves you time or helps you avoid common pitfalls, it’s done its job.

## ⚠️ Security Notice

This is a browser-based SPA. Anything shipped to the client should be treated as public.

Do not include:
- Secrets (client secrets, API keys, certificates)
- Sensitive logic or decision-making
- Server-side enforcement of access controls

This template uses MSAL with Microsoft Entra ID (Authorization Code Flow with PKCE), which is appropriate for public clients.

If your application requires sensitive or security-critical operations, these should be handled by a backend or API layer where access control and enforcement can be trusted.

## Who this is for

- Engineers building Microsoft Entra ID-integrated SPAs
- Teams needing a quickstart for MSAL React + Microsoft Graph
- Anyone tired of stitching together multiple identity samples

## Table of contents

1. [Why this repo](#why-this-repo)
2. [Feature highlights](#feature-highlights)
3. [Prerequisites](#prerequisites)
4. [Setup](#setup)
5. [Available scripts](#available-scripts)
6. [Project layout](#project-layout)
7. [Authentication & Graph architecture](#authentication--graph-architecture)
8. [Customization playbook](#customization-playbook)
9. [Troubleshooting](#troubleshooting)
10. [Helpful links](#helpful-links)

## Why this repo

- **Pre-configured identity plumbing** – MSAL is pre-configured with redirect handling, active-account management, and silent SSO so you can focus on product work.
- **Graph-ready data layer with typed hooks and resilient fetching** – Typed hooks and a resilient fetcher abstract access tokens, pagination, and transient failures.
- **Fluent UI-based application shell** – Fluent UI navigation, responsive layout, loading states, and reusable error dialogs already wired up.

## Feature highlights

- ✅ **Next.js 16 + React 19** App Router with client components where needed and a global provider stack in `app/layout.tsx`.
- ✅ **MSAL React 5** wrapper (`quickstart/providers/msal/MsalClientProvider.tsx`) that initializes once, handles redirect promises, and keeps the active account in sync.
- ✅ **Microsoft Graph helpers**
  - `quickstart/providers/msgraph/msGraphFetcher.ts` injects access tokens, retries pagination via `@odata.nextLink`, and now adds exponential backoff + friendly throttling hints for 429/5xx responses.
  - `quickstart/hooks/useGraphData.ts` wraps SWR for caching, revalidation, and typed responses.
  - `quickstart/hooks/useScopes.ts` detects missing consent and triggers interactive flows.
- ✅ **Fluent UI 9 design system** with a sticky app shell (`quickstart/ui/layouts/ClientShell.tsx`), adaptive nav (`quickstart/ui/navigation/NavMenu.tsx`), and sign-in/out controls fed by MSAL account data.
- ✅ **Route protection baked in** – the `app/(authenticated)/layout.tsx` group wraps every protected page with `AuthenticatedClientShell`, which uses `MsalAuthenticationTemplate` to require sign-in (and display helpful loading/error states) even on deep links.
- ✅ **Sample Graph pages**
  - `app/(authenticated)/profile/page.tsx` renders `/me` data with incremental consent prompts.
  - `app/(authenticated)/organization/page.tsx` calls `/organization` to show tenant metadata (requires `Organization.Read.All`).
- ✅ **Utility components** such as the global `ErrorDialogProvider` and `LoadingScreen` keep UX consistent during auth flows.

## Prerequisites

- Node.js 18.18+ or 20+ (matches Next.js 16 requirements) and npm 9+.
- A Microsoft Entra ID (Azure AD) tenant with permissions to create app registrations.
- An account with permission to grant delegated Microsoft Graph scopes (User.Read, Organization.Read.All, etc.).

## Setup

### 1. Install dependencies

```bash
npm install
```

### 2. Register an Entra ID application

1. Visit [Azure Portal › Entra ID › App registrations](https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps).
2. **New registration** ➜ name it (e.g., `Next.js MSAL Graph SPA`).
3. Supported account types: choose _Single tenant_ (recommended for demos) or _Multi-tenant_ depending on your scenario.
4. Redirect URI (type _Single-page application_): `http://localhost:3000`.
5. After creating the app:
   - Copy the **Application (client) ID**.
   - Under **Authentication**, add your production redirect URIs (e.g., `http://localhost:3000` `https://yourdomain.com`).
   - Under **API permissions**, add delegated Microsoft Graph scopes:
     | Permission | Reason |
     |------------|--------|
     | `User.Read` | Load profile data via `/me` |
     | Any extra scopes your solution needs | Add them now or later |

### 3. Configure environment variables

Duplicate `.env.example` as `.env.local` (or any `.env*.local`) and adjust values:

```ini
NEXT_PUBLIC_APP_NAME="Next.js MSAL React SPA Quickstart"
NEXT_PUBLIC_TENANT_ID="<directory (tenant) ID or 'common' if a multi-tenant app>"
NEXT_PUBLIC_CLIENT_ID="<application (client) ID>"
NEXT_PUBLIC_MSAL_REDIRECT_URI="http://localhost:3000"
NEXT_PUBLIC_MSAL_SCOPES="User.Read"
```

Notes:

- The app now fails fast if `NEXT_PUBLIC_TENANT_ID`, `NEXT_PUBLIC_CLIENT_ID`, or `NEXT_PUBLIC_MSAL_SCOPES` are missing—restart `npm run dev` whenever you edit `.env.local` so Next.js reloads them.
- All variables are read in `config/appConfig.ts` and consumed across MSAL + Graph helpers.
- `NEXT_PUBLIC_MSAL_SCOPES` is a comma-separated list; align it with the permissions you granted. The profile page checks these scopes via `useScopes` before rendering data.
- For multi-environment deployments, create `.env.development.local` / `.env.production.local` or workspace secrets.

### 4. Run the development server

```bash
npm run dev
```

Visit [http://localhost:3000](http://localhost:3000) and select **Sign in**. Once consented, explore **Profile** and **Organization** from the sidebar.

## Available scripts

| Script          | Description                                                                    |
| --------------- | ------------------------------------------------------------------------------ |
| `npm run dev`   | Starts Next.js with hot reload and client-side MSAL initialization logs.       |
| `npm run build` | Production build (generates `.next/`). Useful for CI/CD validation.            |
| `npm run start` | Serves the production build; ensure env vars are set in your hosting platform. |
| `npm run lint`  | Runs ESLint (Next.js config) to catch common mistakes.                         |

## Project layout

```
app/
├─ layout.tsx                 # Global providers + styles
├─ page.tsx                   # Quickstart landing page
└─ (authenticated)/           # Routes that require MSAL auth
   ├─ profile/page.tsx        # /me sample
   └─ organization/page.tsx   # /organization sample
config/
├─ appConfig.ts               # Reads NEXT_PUBLIC_* env vars
└─ NavItems.tsx               # Sidebar + mobile nav definitions
quickstart/
├─ providers/msal/            # MSAL config, helpers, navigation client
├─ providers/msgraph/         # Graph endpoints + fetcher utilities
├─ hooks/                     # Reusable React hooks (Graph data, scopes)
├─ ui/layouts/ClientShell.tsx # Fluent UI shell with navbar + sidebar
├─ ui/navigation/             # Nav components & sign-in/out menu
└─ ui/common/                 # Error dialog + loading spinner
```

## Authentication & Graph architecture

1. **AppProviders** wraps the tree with `MsalClientProvider`, Fluent UI theme, and the `ErrorDialogProvider` so every route can surface failures.
2. **MsalClientProvider** ensures `PublicClientApplication` initializes once per browser, handles `handleRedirectPromise`, sets the active account, and tries silent SSO if possible.
3. **ClientShell** renders a toolbar (`NavBar`) + sidebar (`NavMenu`). `NavMenu` filters routes that require auth using `useIsAuthenticated` so anonymous users only see what they can access.
4. **Sign-in/out buttons** reuse `handleSignIn` / `handleSignOut` (redirect flows) and display profile photos fetched from `/me/photos/48x48/$value` via SWR.
5. **Graph calls**
   - `useGraphData(resource)` → `msGraphFetcher` → fetch with `Authorization: Bearer <token>`.
   - `msGraphFetcher` injects tokens, optional `ConsistencyLevel`, paginates automatically, and throws descriptive errors surfaced by cards/dialogs.
6. **Consent flow**: `useScopes` checks delegated scopes silently, shows descriptive CTA cards, and can force consent when needed.

## Customization playbook

- **Rename the app**: change `NEXT_PUBLIC_APP_NAME` or override the fallback string in `config/appConfig.ts`.
- **Add pages**: duplicate the pattern inside `app/(authenticated)/` and register new routes in `config/NavItems.tsx` with `requiresAuth` flags.
- **Call additional Graph endpoints**: extend `msGraphEndpoints` and reuse `useGraphData` for SWR-powered UI updates.
- **Handle errors centrally**: import `useErrorDialog()` anywhere inside the provider tree to show blocking issues (token failures, Graph throttling, etc.).
- **Style it**: adjust Fluent UI tokens in `quickstart/styling/fluentui/styles.ts` or wrap `FluentProvider` with a custom theme.
- **Charts & data viz**: Recharts 3 is already installed—perfect for rendering usage reports with Graph analytics data.

## Troubleshooting

- **"No active account found"** – Sign in first; `msalInstance` cannot fetch tokens without an account. This is thrown by `getMsGraphAccessToken`.
- **`interaction_required` errors** – The user must consent to new scopes. The Profile page surfaces a "Grant profile access" card that calls `requestConsent()`.
- **Organization screen empty** – Ensure an admin granted `Organization.Read.All`. Without it, Graph returns HTTP 403.
- **Redirect URI mismatch** – Update the SPA redirect in Azure Portal to exactly match `NEXT_PUBLIC_MSAL_REDIRECT_URI` (including protocol + port).

## Helpful links

- [MSAL React documentation](https://learn.microsoft.com/azure/active-directory/develop/tutorial-v2-react)
- [Microsoft Graph delegated permissions](https://learn.microsoft.com/graph/permissions-reference)
- [Fluent UI React Components](https://react.fluentui.dev/)
- [Next.js deployment guide](https://nextjs.org/docs/app/building-your-application/deploying)

---

Need more? File an issue or open a discussion—feedback on the template is welcome!
