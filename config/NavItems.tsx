import { HomeColor, OrgColor, PersonColor } from '@fluentui/react-icons';

interface NavItem {
  label: string;
  route: string;
  icon?: React.ReactNode; // Icon for the header or button
  showActiveIcon?: boolean; // Whether to show the active divider icon
  children?: NavItem[]; // Sub-items if this item has an accordion
  requiresAuth?: boolean;
}

export const navConfig: NavItem[] = [
  {
    label: 'Home',
    route: '/',
    icon: <HomeColor />,
    showActiveIcon: true,
  },
  {
    label: 'Organization',
    route: '/organization',
    icon: <OrgColor />,
    showActiveIcon: true,
    requiresAuth: true,
  },
  {
    label: 'Profile',
    route: '/profile',
    icon: <PersonColor />,
    showActiveIcon: true,
    requiresAuth: true,
  },
];
