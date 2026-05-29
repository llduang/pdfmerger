import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // For local dev server, keep "standalone".
  // For Cloudflare Pages deployment, change to: output: "export"
  // and add: images: { unoptimized: true }
  output: "export",
  typescript: {
    ignoreBuildErrors: true,
  },
  reactStrictMode: false,
  images: {
    unoptimized: true, // Required for static export
  },
};

export default nextConfig;
