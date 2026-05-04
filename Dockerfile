FROM node:20-bookworm-slim

ENV NODE_ENV=production

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev && npm cache clean --force

COPY --chown=node:node public ./public
COPY --chown=node:node templates ./templates
COPY --chown=node:node server.js ./

USER node

EXPOSE 3000

CMD ["node", "server.js"]
