use clap::Parser;
use tracing_subscriber::EnvFilter;

use asperger::asp::config::{AspServerConfig, Config};
use asperger::asp::server::AspServer;

fn init_logging(cli_log_level: Option<&str>, ini_log_level: &str) {
    let filter = if let Ok(filter) = EnvFilter::try_from_default_env() {
        filter
    } else {
        let level = cli_log_level.filter(|l| !l.is_empty()).unwrap_or(ini_log_level);
        EnvFilter::new(level)
    };

    tracing_subscriber::fmt()
        .with_env_filter(filter)
        .with_target(false)
        .init();
}

#[tokio::main]
async fn main() {
    // Parse command-line arguments.
    let cli = Config::parse();

    // If a positional argument is provided, use it as the served folder
    let folder = cli.program.as_deref().unwrap_or(&cli.folder);

    // Build runtime config: start with INI from folder, then apply CLI overrides
    let mut cfg = AspServerConfig::from_folder(folder);
    cfg.apply_overrides(
        Some(&cli.host),
        Some(cli.port),
        Some(folder),
        cli.default_documents.as_deref(),
        Some(cli.enable_directory_listing),
    );
    cfg.apply_log_level(cli.log_level.as_deref());

    // Initialize structured logging.
    // Priority: RUST_LOG env > CLI --log-level > asp.ini log_level > "info"
    init_logging(cli.log_level.as_deref(), &cfg.log_level);

    // Start the server with the specified configuration.
    let server = AspServer::new(cli);
    if let Err(e) = server.start_axum(&cfg).await {
        tracing::error!("Server error: {}", e);
    }
}
