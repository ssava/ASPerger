FROM rust:1.79-slim-bookworm AS builder

WORKDIR /app
COPY Cargo.toml Cargo.lock ./
RUN mkdir src && echo "fn main() {}" > src/main.rs
RUN cargo build --release --bin asperger 2>&1 | tail -5

COPY src/ src/
COPY benches/ benches/
RUN touch src/main.rs
RUN cargo build --release --bin asperger

FROM gcr.io/distroless/cc-debian12:latest

COPY --from=builder /app/target/release/asperger /asperger

EXPOSE 8080

ENV ASPERGER_HOST=0.0.0.0
ENV ASPERGER_PORT=8080
ENV ASPERGER_FOLDER=/asp_files

VOLUME ["/asp_files"]

ENTRYPOINT ["/asperger"]
