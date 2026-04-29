FROM debian:bookworm-slim

RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      libreoffice-writer \
      fonts-liberation \
      fonts-dejavu \
      fonts-noto-core \
    && rm -rf /var/lib/apt/lists/*

ENTRYPOINT ["/usr/bin/soffice"]
