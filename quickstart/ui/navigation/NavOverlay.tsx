import React, { useEffect } from 'react';
import {
  Button,
  Drawer,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  DrawerProps,
} from '@fluentui/react-components';
import { Dismiss24Regular, NavigationRegular } from '@fluentui/react-icons';
import { usePathname } from 'next/navigation';
import NavMenu from './NavMenu';
import { useQuickstartStyles } from '@/quickstart/styling/fluentui/styles';
import { appConfig } from '@/config/appConfig';

type DrawerType = Required<DrawerProps>['type'];

export const NavOverlay = () => {
  const styles = useQuickstartStyles();
  const [isOpen, setIsOpen] = React.useState(false);
  const [type] = React.useState<DrawerType>('overlay');
  const pathname = usePathname();

  useEffect(() => {
    const frame = requestAnimationFrame(() => setIsOpen(false));
    return () => cancelAnimationFrame(frame);
  }, [pathname]);

  return (
    <>
      <Button
        shape='square'
        size='large'
        className={styles.toolbarNavButton}
        appearance='primary'
        onClick={() => setIsOpen(!isOpen)}
        icon={<NavigationRegular />}
      >
        {type === 'inline' ? 'Toggle' : appConfig.APP_NAME}
      </Button>
      <Drawer
        type={type}
        separator
        open={isOpen}
        onOpenChange={(_, { open }) => setIsOpen(open)}
      >
        <DrawerHeader>
          <DrawerHeaderTitle
            action={
              <Button
                appearance='subtle'
                aria-label='Close'
                icon={<Dismiss24Regular />}
                onClick={() => setIsOpen(false)}
              />
            }
          >
            {appConfig.APP_NAME}
          </DrawerHeaderTitle>
        </DrawerHeader>

        <DrawerBody>
          <NavMenu />
        </DrawerBody>
      </Drawer>
    </>
  );
};
