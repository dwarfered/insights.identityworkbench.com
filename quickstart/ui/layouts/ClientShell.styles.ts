import { makeStyles } from '@fluentui/react-components';

export const useClientShellStyles = makeStyles({
  toolbar: {
    position: 'fixed',
    top: 0,
    zIndex: 1000,
    width: '100%',
    backgroundColor: '#fff',
  },

  // Sidebar + content
  mainContainer: {
    display: 'flex',
    height: 'calc(100dvh - var(--toolbar-h, 0px))',
    overflow: 'hidden',
    marginTop: 'var(--toolbar-h, 0px)',
  },

  // Sidebar shell
  sidePanelShell: {
    position: 'sticky',
    top: 0,
    alignSelf: 'stretch',
    height: '100%',
  },

  // Sidebar column
  sidePanel: {
    width: '280px',
    flexShrink: 0,
    height: '100%',
    overflowY: 'auto',
    backgroundColor: '#e9e9e9',
    overflowAnchor: 'none',
    scrollbarGutter: 'stable both-edges',

    '@media (max-width: 768px)': {
      display: 'none',
    },
  },

  // Scroll shell around main content column
  scrollShell: {
    margin: 0,
    overflow: 'hidden',
    backgroundColor: 'transparent',
    height: '100%',
    display: 'flex',
    flexDirection: 'column',
    flexGrow: 1,
    flex: 1,
    minWidth: 0,
    minHeight: 0,
  },

  // Main content column
  container: {
    flexGrow: 1,
    flex: 1,
    height: '100%',
    minHeight: 0,
    overflowY: 'auto',
    margin: 0,
    overflowAnchor: 'none',
    scrollbarGutter: 'stable both-edges',
  },

  content: {
    margin: '16px',
  },
});
