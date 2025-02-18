mod asp;
mod vbscript;

use clap::Parser;

use crate::asp::config::Config;
use crate::asp::server::AspServer;

#[tokio::main]
async fn main() {
    // Leggi i parametri di avvio
    let config = Config::parse();

    // Avvia il server con la configurazione specificata
    let server = AspServer::new(config);
    if let Err(e) = server.start().await {
        eprintln!("Errore del server: {}", e);
    }
}