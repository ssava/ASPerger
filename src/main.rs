use clap::Parser;

use asperger::asp::config::{AspServerConfig, Config};
use asperger::asp::server::AspServer;

#[tokio::main]
async fn main() {
    // Parse command-line arguments.
    let cli = Config::parse();

    // Build runtime config: start with INI from folder, then apply CLI overrides
    let mut cfg = AspServerConfig::from_folder(&cli.folder);
    cfg.apply_overrides(
        Some(&cli.host),
        Some(cli.port),
        Some(&cli.folder),
        None, // default_document not settable from CLI yet
    );

    // Start the server with the specified configuration.
    let server = AspServer::new(cli);
    if let Err(e) = server.start_with_config(&cfg).await {
        eprintln!("Server error: {}", e);
    }
}
