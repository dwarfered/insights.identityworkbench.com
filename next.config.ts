import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  /* config options here */
  output: 'export', // this is saying static export for SPA build
  reactCompiler: true,
};

export default nextConfig;
