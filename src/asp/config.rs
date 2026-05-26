use clap::Parser;

/// Configurazione del server ASP
#[derive(Parser, Debug)]
#[clap(author, version, about, long_about = None)]
pub struct Config {
    /// Indirizzo su cui il server sarà in ascolto
    #[clap(long, default_value = "127.0.0.1")]
    pub host: String,

    /// Porta su cui il server sarà in ascolto
    #[clap(short, long, default_value = "8080")]
    pub port: u16,

    /// Cartella contenente i file ASP
    #[clap(short, long, default_value = "./")]
    pub folder: String,
}