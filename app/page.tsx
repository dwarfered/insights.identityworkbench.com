'use client';

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
      'Run focused Graph queries to surface SKU assignment, group targeting, and actions to make license reviews faster.',
  },
  {
    title: 'Actionable breakdowns',
    description:
      'Drill into employee-type splits, group ownership, and member counts so you can align license decisions.',
  },
  {
    title: 'Reusable Graph patterns',
    description:
      'Every hook doubles as sample code-clone, adapt, or extend them in your own Entra management automations.',
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
          A living showcase of Microsoft Graph queries and actions that assist
          Entra administrators with day-to-day tenant management.
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
        <Subtitle2>What this Insights project highlights</Subtitle2>
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
