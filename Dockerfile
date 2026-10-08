FROM node:24-slim
WORKDIR /app
COPY package*.json ./
RUN npm ci --omit=dev
COPY . .
ENV NODE_ENV=production DATA_DIR=/data TRUST_PROXY=1 PORT=3000
EXPOSE 3000
CMD ["node", "--disable-warning=ExperimentalWarning", "server.js"]
