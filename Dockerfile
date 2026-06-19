# deck-tools sidecar service
#
# Runs alongside your n8n container (NOT inside it). n8n's official image
# is distroless from v2.x onward, so apk/apt-get and LibreOffice cannot be
# added to it directly any more. This image is a small standalone Node
# service that n8n's HTTP Request node calls over the docker network.
#
# Set DECK_TOOLS_SECRET at runtime (e.g. `docker run -e DECK_TOOLS_SECRET=...`
# or in docker-compose `environment:`) if this service is reachable from the
# public internet rather than a private network n8n already shares. Requests
# must then include a matching x-deck-tools-secret header.

FROM node:22-slim

# libreoffice-impress is enough for pptx -> pdf; the smaller "core" + this
# component avoids pulling in the full Writer/Calc/Draw stack.
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      libreoffice-impress \
      poppler-utils \
      fonts-dejavu \
      fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm install --omit=dev

COPY server.js build_deck.js ./

ENV PORT=4000
EXPOSE 4000

CMD ["node", "server.js"]
