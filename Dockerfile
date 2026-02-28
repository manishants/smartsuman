# Use the official Node.js 20 image with Debian for LibreOffice support
FROM node:20-bookworm AS base

# Install dependencies only when needed
FROM base AS deps
WORKDIR /app

# Install dependencies based on the preferred package manager
COPY package.json package-lock.json* ./
# Speed up install and ensure native modules (sharp) can compile if needed
ENV PUPPETEER_SKIP_DOWNLOAD=1
ENV PYTHON=/usr/bin/python3
RUN set -eux; \
    apt-get update; \
    apt-get install -y --no-install-recommends \
      build-essential \
      python3 \
      pkg-config \
      libvips-dev; \
    rm -rf /var/lib/apt/lists/*; \
    npm ci --no-audit --no-fund

# Rebuild the source code only when needed
FROM base AS builder
WORKDIR /app
COPY --from=deps /app/node_modules ./node_modules
COPY . .
RUN rm -rf src/app/admin src/app/superadmin

# Next.js collects completely anonymous telemetry data about general usage.
# Learn more here: https://nextjs.org/telemetry
# Uncomment the following line in case you want to disable telemetry during the build.
ENV NEXT_TELEMETRY_DISABLED=1

# Build the application
RUN npm run build

# Production image, copy all the files and run next
FROM base AS runner
WORKDIR /app

ENV NODE_ENV=production
# Uncomment the following line in case you want to disable telemetry during runtime.
ENV NEXT_TELEMETRY_DISABLED=1
 # Explicit path for LibreOffice binary to ensure discovery in app code
 ENV LIBREOFFICE_PATH=/usr/bin/soffice
ENV TMPDIR=/tmp

# Install LibreOffice and minimal runtime libs in production image
ENV DEBIAN_FRONTEND=noninteractive
RUN set -eux; \
    apt-get update; \
    apt-get install -y --no-install-recommends \
      # LibreOffice headless components
      libreoffice-core \
      libreoffice-common \
      libreoffice-writer \
      libreoffice-calc \
      libreoffice-impress \
      fonts-liberation \
      fonts-dejavu \
      fonts-freefont-ttf \
      fonts-opensymbol \
      ca-certificates \
      libnss3 \
      libx11-6 \
      libxcomposite1 \
      libxdamage1 \
      libxext6 \
      libxfixes3 \
      libxrandr2 \
      libxkbcommon0 \
      libxshmfence1 \
      libxss1 \
      libgbm1 \
      libgtk-3-0 \
      libatk1.0-0 \
      libatk-bridge2.0-0 \
      libcups2 \
      libdrm2 \
      libasound2 \
      xdg-utils \
      # Runtime libs for sharp
      libvips \
      # Chromium for Puppeteer
      chromium; \
    apt-get clean; \
    rm -rf /var/lib/apt/lists/*; \
    chmod 1777 /tmp; \
    rm -rf /var/tmp && ln -s /tmp /var/tmp

RUN addgroup --system --gid 1001 nodejs
RUN adduser --system --uid 1001 nextjs

# Ensure writable home and XDG dirs for LibreOffice/dconf
ENV HOME=/home/nextjs
ENV XDG_CONFIG_HOME=/home/nextjs/.config
ENV XDG_CACHE_HOME=/home/nextjs/.cache
ENV XDG_RUNTIME_DIR=/tmp/xdg
RUN mkdir -p /home/nextjs/.config /home/nextjs/.cache /tmp/xdg \
    && chown -R nextjs:nodejs /home/nextjs /tmp/xdg

COPY --from=builder /app/public ./public

# Set the correct permission for prerender cache
RUN mkdir .next
RUN chown nextjs:nodejs .next

# Automatically leverage output traces to reduce image size
# https://nextjs.org/docs/advanced-features/output-file-tracing
COPY --from=builder --chown=nextjs:nodejs /app/.next/standalone ./
COPY --from=builder --chown=nextjs:nodejs /app/.next/static ./.next/static

USER nextjs

EXPOSE 3000

ENV PORT=3000
ENV HOSTNAME="0.0.0.0"
ENV PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium

# server.js is created by next build from the standalone output
# https://nextjs.org/docs/pages/api-reference/next-config-js/output
CMD ["node", "server.js"]
