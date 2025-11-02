# syntax=docker/dockerfile:1
FROM debian:bookworm-slim

# Install build tools and dependencies
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        git \
        build-essential \
        bmake \
        cmake \
        flex \
        bison \
        graphviz \
        pkg-config \
        libgraphviz-dev \
        libexpat1-dev \
        libxml2-dev \
        ca-certificates \
        && rm -rf /var/lib/apt/lists/*

# Clone and build kgt
WORKDIR /opt
RUN git clone --recursive https://github.com/katef/kgt.git && \
    cd kgt && \
    bmake -r install


# Default workdir for user input
WORKDIR /data

# Example: mount your grammar.bnf and run
# docker run --rm -v $(pwd):/data kgt grammar.bnf -l bnf -e abnf -o diagram.svg
ENTRYPOINT ["kgt"]