import React from 'react';
import { Button, Body1, makeStyles, tokens } from '@fluentui/react-components';
import { useRouter, usePathname } from 'next/navigation';

interface NavButtonProps {
  label: string;
  route: string;
  showActiveIcon?: boolean;
  icon?: React.ReactNode;
}

const useStyles = makeStyles({
  button: {
    width: '100%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'flex-start',
    columnGap: tokens.spacingHorizontalS,
    paddingLeft: '0',
    paddingRight: tokens.spacingHorizontalM,
    minHeight: '40px',
  },
  indicatorSlot: {
    display: 'flex',
    alignItems: 'center',
    alignSelf: 'stretch',
    width: '4px',
    marginLeft: tokens.spacingHorizontalXXS,
  },
  indicator: {
    width: '4px',
    backgroundColor: tokens.colorCompoundBrandForeground1,
    display: 'block',
    height: '24px',
  },
  navIcon: {
    display: 'flex',
    alignItems: 'center',
    fontSize: '18px',
  },
  label: {
    color: tokens.colorNeutralForeground1,
  },
});

export function NavButton({
  label,
  route,
  icon,
  showActiveIcon = false,
}: NavButtonProps) {
  const router = useRouter();
  const pathname = usePathname();
  const styles = useStyles();

  const isActive = pathname === route;
  const buttonStyle = {
    paddingLeft: showActiveIcon ? '0' : tokens.spacingHorizontalM,
  };

  return (
    <Button
      shape='square'
      appearance='subtle'
      onClick={() => router.push(route)}
      className={styles.button}
      style={buttonStyle}
    >
      {showActiveIcon && (
        <span className={styles.indicatorSlot}>
          {isActive ? <span className={styles.indicator} /> : null}
        </span>
      )}
      {icon ? <span className={styles.navIcon}>{icon}</span> : null}
      <Body1 className={styles.label}>{label}</Body1>
    </Button>
  );
}
