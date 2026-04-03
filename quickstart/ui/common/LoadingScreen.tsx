'use client';

import { Spinner } from '@fluentui/react-components';

export function LoadingScreen({ label }: { label?: string }) {
  return (
    <div
      style={{
        minHeight: '100vh',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      <Spinner label={label ?? 'Loading...'} size='extra-large' />
    </div>
  );
}
