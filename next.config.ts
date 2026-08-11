/** @type {import('next').NextConfig} */
const nextConfig = {
  async rewrites() {
    return [
      {
        source: '/raileats-reports',
        destination: 'https://raileats-reports.vercel.app',
      },
      {
        source: '/raileats-reports/:path*',
        destination: 'https://raileats-reports.vercel.app/:path*',
      },
    ];
  },
};

module.exports = nextConfig;
