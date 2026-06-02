use clap::Parser;

use asperger::asp::config::Config;
use asperger::asp::server::AspServer;

#[tokio::main]
async fn main() {
    // Parse command-line arguments.
    let config = Config::parse();

    // Start the server with the specified configuration.
    let server = AspServer::new(config);
    if let Err(e) = server.start().await {
        eprintln!("Server error: {}", e);
    }
}