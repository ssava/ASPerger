mod asp;
mod vbscript;

#[tokio::main]
async fn main() {
    let server = asp::server::AspServer::new();
    if let Err(e) = server.start(8080).await {
        eprintln!("Errore del server: {}", e);
    }
}