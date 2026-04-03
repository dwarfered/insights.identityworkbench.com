'use client';

import { msalInstance } from '@/quickstart/providers/msal/msalAuthConfig';
import { handleSignOut } from '@/quickstart/providers/msal/msalUtilities';
import { msGraphEndpoints } from '@/quickstart/providers/msgraph/msGraphEndpoints';
import { msGraphFetcher } from '@/quickstart/providers/msgraph/msGraphFetcher';
import {
  Avatar,
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { useEffect, useMemo } from 'react';
import useSWR from 'swr';

const useStyles = makeStyles({
  buttonContent: {
    paddingRight: '8px',
    display: 'flex',
    alignItems: 'center',
    maxWidth: '250px',
    backgroundColor: '#1c1c1c',
    color: tokens.colorBrandBackgroundInverted,
    ':hover:active': {
      backgroundColor: '#333',
      color: '#fff',
    },
    ':hover': {
      backgroundColor: '#333',
      color: '#fff',
    },
  },
  textContainer: {
    textAlign: 'left',
    whiteSpace: 'nowrap',
    textOverflow: 'ellipsis',
    paddingRight: '8px',
    paddingLeft: '8px',
  },
  text: {
    whiteSpace: 'nowrap',
    textOverflow: 'ellipsis',
    textAlign: 'right',
  },
});

type OrganizationResponse = {
  value?: Array<{
    displayName?: string;
  }>;
};

export default function SignOutButton() {
  const styles = useStyles();
  const account = msalInstance.getActiveAccount();

  const { data: photoBlob } = useSWR<Blob>(
    msGraphEndpoints.graphMePhoto,
    msGraphFetcher,
    {
      revalidateOnFocus: false,
      revalidateOnReconnect: false,
      refreshInterval: 0,
    },
  );

  const { data: orgData } = useSWR<OrganizationResponse>(
    msGraphEndpoints.organization,
    (url: string) => msGraphFetcher(url),
    {
      revalidateOnFocus: false,
      revalidateOnReconnect: false,
      refreshInterval: 0,
    },
  );

  const imageSrc = useMemo(() => {
    if (!photoBlob || !(photoBlob instanceof Blob)) {
      return undefined;
    }
    return URL.createObjectURL(photoBlob);
  }, [photoBlob]);

  useEffect(() => {
    return () => {
      if (imageSrc) {
        URL.revokeObjectURL(imageSrc);
      }
    };
  }, [imageSrc]);

  const tenantName = useMemo(() => {
    const firstOrg = orgData?.value?.[0];
    return firstOrg?.displayName;
  }, [orgData]);

  if (!account) {
    return null;
  }

  return (
    <>
      <Menu positioning='below-end'>
        <MenuTrigger disableButtonEnhancement>
          <Button
            className={styles.buttonContent}
            size='small'
            shape='square'
            appearance='primary'
            style={{ minWidth: 0 }}
          >
            <div
              className={styles.textContainer}
              style={{ flex: 1, overflow: 'hidden' }}
            >
              <div className={styles.text} style={{ overflow: 'hidden' }}>
                {account?.username}
              </div>
              <div className={styles.text} style={{ overflow: 'hidden' }}>
                {tenantName ?? account?.name}
              </div>
            </div>
            <Avatar
              name={account?.name}
              image={imageSrc ? { src: imageSrc } : undefined}
            />
          </Button>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem
              onClick={() => handleSignOut()}
              key='logoutRedirect'
              style={{ textAlign: 'left' }}
            >
              Sign out
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>
    </>
  );
}
