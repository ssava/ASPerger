use clap::Parser;

/// ASP server configuration.
#[derive(Parser, Debug)]
#[clap(author, version, about, long_about = None)]
pub struct Config {
    /// Host address the server will listen on.
    #[clap(long, default_value = "127.0.0.1")]
    pub host: String,

    /// Port the server will listen on.
    #[clap(short, long, default_value = "8080")]
    pub port: u16,

    /// Directory containing ASP files to serve.
    #[clap(short, long, default_value = "./")]
    pub folder: String,
}