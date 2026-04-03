'use client';

import * as React from 'react';
import { Button } from '@fluentui/react-components';
import { ArrowUp20Filled } from '@fluentui/react-icons';

export function GoToTop({ targetId }: { targetId?: string }) {
  const [visible, setVisible] = React.useState(false);

  React.useEffect(() => {
    const scrollEl = targetId ? document.getElementById(targetId) : window;

    if (!scrollEl) return;

    const onScroll = () => {
      const pos =
        scrollEl instanceof Window ? scrollEl.scrollY : scrollEl.scrollTop;
      setVisible(pos > 200); // show after 200px
    };

    scrollEl.addEventListener('scroll', onScroll);
    return () => scrollEl.removeEventListener('scroll', onScroll);
  }, [targetId]);

  const scrollToTop = () => {
    if (targetId) {
      document
        .getElementById(targetId)
        ?.scrollTo({ top: 0, behavior: 'smooth' });
    } else {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  };

  if (!visible) return null;

  return (
    <Button
      shape='circular'
      size='large'
      appearance='primary'
      icon={<ArrowUp20Filled />}
      onClick={scrollToTop}
      style={{
        position: 'fixed',
        bottom: '2rem',
        right: '4rem',
        backgroundColor: '#0078D4', // Azure blue
        color: 'white',
        width: '32px',
        height: '32px',
        minWidth: '32px', // keep it round
        boxShadow: '0 4px 12px rgba(0,0,0,0.3)',
        zIndex: 1000,
      }}
      aria-label='Go to top'
    />
  );
}
