import { Avatar, Button } from '@fluentui/react-components';
import { useQuickstartStyles } from '@/quickstart/styling/fluentui/styles';
import { handleSignIn } from '@/quickstart/providers/msal/msalUtilities';

export default function SignInButton() {
  const styles = useQuickstartStyles();

  return (
    <Button
      className={styles.toolbarNavButton}
      size='small'
      shape='square'
      appearance='primary'
      style={{ minWidth: 0, columnGap: '8px' }}
      onClick={() => handleSignIn()}
    >
      Sign In
      <Avatar name={undefined} image={undefined} />
    </Button>
  );
}
