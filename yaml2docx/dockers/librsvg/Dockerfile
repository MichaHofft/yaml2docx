# Use a minimal Debian base image
FROM debian:bookworm-slim

# Install rsvg-convert and clean up
RUN apt-get update && \
    apt-get install -y --no-install-recommends librsvg2-bin && \
    rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Default command (can be overridden)
ENTRYPOINT ["rsvg-convert"]