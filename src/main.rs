use clap::Parser;

use asperger::asp::config::Config;
use asperger::asp::server::AspServer;

#[tokio::main]
async fn main() {
    // Leggi i parametri di avvio
    let config = Config::parse();

    // Avvia il server con la configurazione specificata
    let server = AspServer::new(config);
    if let Err(e) = server.start().await {
        eprintln!("Server error: {}", e);
    }
}