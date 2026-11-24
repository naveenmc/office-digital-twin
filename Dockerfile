FROM node:24-alpine
WORKDIR /app
COPY --chown=node:node Builds/WebGL ./Builds/WebGL
COPY --chown=node:node Server/server.mjs ./Server/server.mjs
ENV NODE_ENV=production OFFICE_HOST=0.0.0.0 OFFICE_DATA_DIR=/tmp/office
USER node
EXPOSE 8787
CMD ["node", "Server/server.mjs"]
