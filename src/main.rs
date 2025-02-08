mod asp;
mod vbscript;

fn main() {
    let server = asp::server::AspServer::new();
    if let Err(e) = server.start(8080) {
        eprintln!("Errore del server: {}", e);
    }
}