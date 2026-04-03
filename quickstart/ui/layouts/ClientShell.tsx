'use client';

import { useEffect, useLayoutEffect, useRef } from 'react';
import { useClientShellStyles } from './ClientShell.styles';
import { usePathname } from 'next/navigation';
import NavBar from '@/quickstart/ui/navigation/NavBar';
import SideBar from '@/quickstart/ui/navigation/SideBar';
import { GoToTop } from '@/quickstart/ui/navigation/GoToTop';

export default function ClientShell({
  children,
}: {
  children: React.ReactNode;
}) {
  const styles = useClientShellStyles();
  const sidePanelRef = useRef<HTMLDivElement | null>(null);
  const contentRef = useRef<HTMLDivElement | null>(null);
  const toolbarRef = useRef<HTMLDivElement | null>(null);
  const pathname = usePathname();

  useLayoutEffect(() => {
    const el = toolbarRef.current;
    if (!el) return;
    const setToolbarHeightVar = () => {
      const h = el.getBoundingClientRect().height;
      document.documentElement.style.setProperty('--toolbar-h', `${h}px`);
    };

    setToolbarHeightVar();
    const ro = new ResizeObserver(() => setToolbarHeightVar());
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  useEffect(() => {
    if ('scrollRestoration' in window.history) {
      const prev = window.history.scrollRestoration as 'auto' | 'manual';
      window.history.scrollRestoration = 'manual';
      return () => {
        window.history.scrollRestoration = prev;
      };
    }
  }, []);

  useEffect(() => {
    if (typeof document === 'undefined') return;

    const sc = contentRef.current;
    const sp = sidePanelRef.current;
    const root = document.documentElement;

    const previousBehavior = root.style.scrollBehavior;
    root.style.scrollBehavior = 'auto';

    if (sc) sc.scrollTop = 0;
    if (sp) sp.scrollTop = 0;

    root.style.scrollBehavior = previousBehavior;
  }, [pathname]);

  return (
    <>
      <div className={styles.toolbar} ref={toolbarRef}>
        <NavBar />
      </div>
      <div className={styles.mainContainer}>
        <div className={styles.sidePanelShell}>
          <div className={styles.sidePanel} ref={sidePanelRef}>
            <SideBar />
          </div>
        </div>
        <div className={styles.scrollShell}>
          <div
            className={styles.container}
            id='scroll-container'
            ref={contentRef}
          >
            <div className={styles.content}>{children}</div>
            <GoToTop targetId='scroll-container' />
          </div>
        </div>
      </div>
    </>
  );
}
