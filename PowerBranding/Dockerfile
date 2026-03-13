FROM node:22-bookworm

RUN apt-get update && apt-get install -y \
    git \
    python3 \
    make \
    g++ \
    ca-certificates \
    openssl \
    && rm -rf /var/lib/apt/lists/*

RUN npm install -g yo @microsoft/generator-sharepoint @rushstack/heft

USER node
WORKDIR /workspace

ENV HOME=/home/node
ENV XDG_CONFIG_HOME=/home/node/.config

CMD ["bash"]