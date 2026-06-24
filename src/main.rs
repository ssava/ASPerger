use clap::Parser;

use asperger::asp::config::{AspServerConfig, Config};
use asperger::asp::server::AspServer;

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

    // Start the server with the specified configuration.
    let server = AspServer::new(cli);
    if let Err(e) = server.start_axum(&cfg).await {
        eprintln!("Server error: {}", e);
    }
}
