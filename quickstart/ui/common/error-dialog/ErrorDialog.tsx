import * as React from 'react';
import {
  Dialog,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogActions,
  Button,
} from '@fluentui/react-components';

export interface ErrorDialogProps {
  open: boolean;
  title?: string;
  message: string | null;
  onClose: () => void;
}

export function ErrorDialog({
  open,
  title = 'Error',
  message,
  onClose,
}: ErrorDialogProps) {
  return (
    <Dialog open={open} onOpenChange={(_e, data) => !data.open && onClose()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{title}</DialogTitle>
          <DialogContent>
            <pre
              style={{
                whiteSpace: 'pre-wrap',
                wordBreak: 'break-word',
                fontFamily: 'inherit',
              }}
            >
              {message}
            </pre>
          </DialogContent>
          <DialogActions>
            <Button appearance='primary' onClick={onClose}>
              Close
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

interface ErrorState {
  open: boolean;
  title: string;
  message: string | null;
}

interface ErrorDialogContextValue {
  showError: (message: string | Error, options?: { title?: string }) => void;
  hideError: () => void;
}

const ErrorDialogContext = React.createContext<
  ErrorDialogContextValue | undefined
>(undefined);

export function useErrorDialog(): ErrorDialogContextValue {
  const context = React.useContext(ErrorDialogContext);
  if (!context) {
    throw new Error(
      'useErrorDialog must be used within an ErrorDialogProvider',
    );
  }
  return context;
}

export function ErrorDialogProvider({
  children,
}: {
  children: React.ReactNode;
}): React.ReactElement {
  const [state, setState] = React.useState<ErrorState>({
    open: false,
    title: 'Error',
    message: null,
  });

  const showError = React.useCallback<ErrorDialogContextValue['showError']>(
    (message, options) => {
      const resolvedMessage =
        typeof message === 'string'
          ? message
          : (message.message ?? 'An error occurred.');

      setState({
        open: true,
        title: options?.title ?? 'Error',
        message: resolvedMessage,
      });
    },
    [],
  );

  const hideError = React.useCallback<
    ErrorDialogContextValue['hideError']
  >(() => {
    setState((prev) => ({ ...prev, open: false }));
  }, []);

  const contextValue = React.useMemo(
    () => ({ showError, hideError }),
    [showError, hideError],
  );

  return (
    <ErrorDialogContext.Provider value={contextValue}>
      {children}
      <ErrorDialog
        open={state.open}
        title={state.title}
        message={state.message}
        onClose={hideError}
      />
    </ErrorDialogContext.Provider>
  );
}
