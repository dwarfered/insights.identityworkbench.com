'use client';

import { useRouter } from 'next/navigation';
import {
  Body1,
  Button,
  Link,
  Title1,
  Subtitle2,
  makeStyles,
  shorthands,
  tokens,
} from '@fluentui/react-components';
import ClientShell from '@/quickstart/ui/layouts/ClientShell';

const useStyles = makeStyles({
  hero: {
    ...shorthands.padding('24px'),
    backgroundColor: '#f7f7fb',
    borderRadius: tokens.borderRadiusXLarge,
    marginBottom: '32px',
  },
  heroActions: {
    marginTop: tokens.spacingVerticalM,
    display: 'flex',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
  },
  section: {
    marginBottom: '32px',
  },
  featureGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))',
    gap: tokens.spacingHorizontalL,
  },
  resourceList: {
    listStyle: 'disc',
    paddingLeft: '20px',
    margin: 0,
  },
});

const resourceLinks = [
  {
    label: 'README quickstart guide',
    href: 'https://github.com/dwarfered/nextjs-msal-react-graph-spa-template',
  },
  {
    label: 'Microsoft identity platform docs',
    href: 'https://learn.microsoft.com/azure/active-directory/develop/',
  },
  {
    label: 'Microsoft Graph permission reference',
    href: 'https://learn.microsoft.com/graph/permissions-reference',
  },
];

export default function Home() {
  const styles = useStyles();
  const router = useRouter();

  return (
    <ClientShell>
      <section className={styles.hero}>
        <Title1>Next.js + MSAL React + Microsoft Graph quickstart</Title1>
        <br />
        <Body1 style={{ marginTop: tokens.spacingVerticalS }}>
          Easily build a single or multi-tenant Next.js React SPA that
          authenticates against Microsoft Entra ID, calls Microsoft Graph with
          resilient SWR hooks, and showcases real user + organization data out
          of the box.
        </Body1>
        <div className={styles.heroActions}>
          <Button appearance='primary' onClick={() => router.push('/profile')}>
            View Graph profile sample
          </Button>
          <Button
            as='a'
            href='https://github.com/dwarfered/nextjs-msal-react-graph-spa-template'
            target='_blank'
            rel='noreferrer'
          >
            Open README
          </Button>
        </div>
      </section>

      <section className={styles.section}>
        <Subtitle2>Deep-dive resources</Subtitle2>
        <ul className={styles.resourceList}>
          {resourceLinks.map((link) => (
            <li key={link.href}>
              <Link href={link.href} target='_blank' rel='noreferrer'>
                {link.label}
              </Link>
            </li>
          ))}
        </ul>
      </section>
    </ClientShell>
  );
}
