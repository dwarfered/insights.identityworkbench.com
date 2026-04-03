import React from 'react';
// import {
//   AuthenticatedTemplate,
//   UnauthenticatedTemplate,
// } from "@azure/msal-react";
import {
  Button,
  makeStyles,
  shorthands,
  Toolbar,
  ToolbarGroup,
  DrawerProps,
} from '@fluentui/react-components';
import { NavOverlay } from './NavOverlay';
import {
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
} from '@azure/msal-react';
import { useQuickstartStyles } from '@/quickstart/styling/fluentui/styles';
import { appConfig } from '@/config/appConfig';
import SignInButton from './sign-in-sign-out/SignInButton';
import SignOutButton from './sign-in-sign-out/SignOutButton';

const useStyles = makeStyles({
  toolbar: {
    justifyContent: 'space-between',
    backgroundColor: '#1c1c1c',
    ...shorthands.padding(0),
  },
  navMobileOverlay: {
    '@media (min-width: 768px)': {
      display: 'none',
    },
  },
  navStandard: {
    '@media (max-width: 768px)': {
      display: 'none',
    },
  },
});

type DrawerType = Required<DrawerProps>['type'];

export const FarGroup = () => {
  const standardStyles = useQuickstartStyles();
  const styles = useStyles();
  const [type] = React.useState<DrawerType>('overlay');

  return (
    <Toolbar aria-label='with Separeted Groups' className={styles.toolbar}>
      <ToolbarGroup role='presentation'>
        <div className={styles.navMobileOverlay}>
          <NavOverlay />
        </div>
        <div className={styles.navStandard}>
          <Button
            shape='square'
            size='large'
            className={standardStyles.toolbarNavButton}
            appearance='primary'
          >
            {type === 'inline' ? 'Toggle' : appConfig.APP_NAME}
          </Button>
        </div>
      </ToolbarGroup>
      <ToolbarGroup role='presentation'>
        <UnauthenticatedTemplate>
          <SignInButton />
        </UnauthenticatedTemplate>
        <AuthenticatedTemplate>
          <SignOutButton />
        </AuthenticatedTemplate>
      </ToolbarGroup>
    </Toolbar>
  );
};

const NavBar = () => {
  return <FarGroup />;
};

export default NavBar;
