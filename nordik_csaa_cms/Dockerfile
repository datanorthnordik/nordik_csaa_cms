# Build stage
FROM node:22-alpine AS builder
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY . .
RUN npm run build

# Production stage
FROM nginx:alpine
LABEL maintainer="athulnarayanan62@gmail.com"
WORKDIR /usr/share/nginx/html
RUN rm -rf ./*
COPY --from=builder /app/dist ./
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD wget --no-verbose --tries=1 --spider http://localhost/index.html || exit 1

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
