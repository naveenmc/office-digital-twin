FROM node:24-alpine AS assemble
WORKDIR /app
COPY office-asset-* office-files.json assemble-office.mjs restore-build.mjs build-parts.json ./
RUN node assemble-office.mjs
FROM node:24-alpine
WORKDIR /app
COPY --from=assemble --chown=node:node /app/Builds ./Builds
COPY --from=assemble --chown=node:node /app/Server ./Server
ENV NODE_ENV=production OFFICE_HOST=0.0.0.0 OFFICE_DATA_DIR=/tmp/office
USER node
EXPOSE 8787
CMD ["node", "Server/server.mjs"]
