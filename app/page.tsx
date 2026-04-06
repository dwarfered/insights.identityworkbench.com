'use client';

import Image from 'next/image';
import { useRouter } from 'next/navigation';
import {
  Body1,
  Button,
  Link,
  Subtitle2,
  Text,
  Title1,
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
  sectionHeading: {
    display: 'block',
    marginBottom: tokens.spacingVerticalXL,
  },
  featureCard: {
    ...shorthands.padding(tokens.spacingHorizontalM),
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground3,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
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

const featureHighlights = [
  {
    title: 'Targeted licensing insights',
    description:
      'Efficient Microsoft Graph queries, surfacing SKU assignments and the current landscape.',
  },
  {
    title: 'Actionable breakdowns',
    description: 'Employee types, group-based licensing, and member counts.',
  },
  {
    title: 'Reusable Microsoft Graph patterns',
    description:
      'Each component is a copy-paste ready sample for your own automation or insights apps.',
  },
];

export default function Home() {
  const styles = useStyles();
  const router = useRouter();

  return (
    <ClientShell>
      <section className={styles.hero}>
        <Title1>Insights - Identity Workbench</Title1>
        <br />
        <Body1 style={{ marginTop: tokens.spacingVerticalS }}>
          This is my sandbox for trying Microsoft Graph ideas that make Entra
          management easier. The samples are practical, and easy to take for
          your own tenant.
        </Body1>
        <div className={styles.heroActions}>
          <Button
            appearance='primary'
            onClick={() => router.push('/licensing')}
          >
            Explore licensing insights
          </Button>
          <Button
            as='a'
            href='https://github.com/dwarfered/insights.identityworkbench.com'
            target='_blank'
            rel='noreferrer'
          >
            View Insights source
          </Button>
          <Button
            as='a'
            href='https://github.com/dwarfered/nextjs-msal-react-graph-spa-template'
            target='_blank'
            rel='noreferrer'
          >
            View template source
          </Button>
        </div>
      </section>

      <section className={styles.section}>
        <Subtitle2 className={styles.sectionHeading}>
          What you'll find so far
        </Subtitle2>
        <div className={styles.featureGrid}>
          {featureHighlights.map((feature) => (
            <div key={feature.title} className={styles.featureCard}>
              <Text weight='semibold'>{feature.title}</Text>
              <Body1>{feature.description}</Body1>
            </div>
          ))}
        </div>
      </section>

      <section className={styles.section}>
        <Subtitle2 className={styles.sectionHeading}>
          Example signed-in session
        </Subtitle2>
        <Image
          src='/docs/insights-hero.png'
          alt='Insights licensing dashboard'
          width={2266}
          height={1746}
          style={{
            width: '100%',
            maxWidth: 960,
            height: 'auto',
            borderRadius: tokens.borderRadiusXLarge,
          }}
          priority
        />
      </section>

      <section className={styles.section}>
        <Subtitle2 className={styles.sectionHeading}>
          Deep-dive resources
        </Subtitle2>
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
