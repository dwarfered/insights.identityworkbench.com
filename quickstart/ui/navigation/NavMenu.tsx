import {
  Toolbar,
  ToolbarGroup,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Divider,
  Body1Strong,
  makeStyles,
} from '@fluentui/react-components';

import { NavButton } from './NavButton';
import { navConfig } from '@/config/NavItems';
import { useIsAuthenticated } from '@azure/msal-react';

const useStyles = makeStyles({
  toolbar: {
    width: '100%',
    alignItems: 'stretch',
    padding: 0,
    gap: 0,
  },
  group: {
    width: '100%',
    alignItems: 'stretch',
    gap: 0,
  },
  navItemWrapper: {
    width: '100%',
  },
  accordion: {
    width: '100%',
  },
  accordionPanel: {
    paddingLeft: 0,
    paddingRight: 0,
  },
});

export default function NavMenu() {
  const styles = useStyles();
  const isAuthenticated = useIsAuthenticated();

  const filteredItems = navConfig
    .map((item) => {
      if (item.children && item.children.length > 0) {
        const filteredChildren = item.children.filter(
          (child) => !child.requiresAuth || isAuthenticated,
        );
        return {
          ...item,
          children: filteredChildren,
        };
      }
      return item;
    })
    .filter((item) => !item.requiresAuth || isAuthenticated)
    .filter((item) => {
      if (!item.children) {
        return true;
      }
      return item.children.length > 0;
    });

  return (
    <Toolbar aria-label='Navigation Menu' vertical className={styles.toolbar}>
      <ToolbarGroup role='presentation' className={styles.group}>
        {filteredItems.map((item, index) => {
          if (item.children && item.children.length > 0) {
            return (
              <Accordion
                key={index}
                defaultOpenItems='1'
                className={styles.accordion}
              >
                <AccordionItem value='1'>
                  <AccordionHeader
                    icon={{
                      as: 'div',
                      children: item.icon,
                    }}
                    expandIconPosition='end'
                  >
                    <Body1Strong>{item.label}</Body1Strong>
                  </AccordionHeader>
                  <AccordionPanel className={styles.accordionPanel}>
                    {item.children.map((child, childIndex) => (
                      <NavButton
                        key={childIndex}
                        label={child.label}
                        route={child.route}
                        showActiveIcon={child.showActiveIcon}
                      />
                    ))}
                  </AccordionPanel>
                </AccordionItem>
              </Accordion>
            );
          }

          return (
            <div key={index} className={styles.navItemWrapper}>
              <NavButton
                label={item.label}
                route={item.route}
                icon={item.icon}
                showActiveIcon={item.showActiveIcon}
              />
              {index < navConfig.length - 1 && <Divider />}
            </div>
          );
        })}
      </ToolbarGroup>
    </Toolbar>
  );
}
