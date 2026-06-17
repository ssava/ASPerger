use criterion::{criterion_group, criterion_main, Criterion};
use std::sync::Arc;
use std::time::Instant;

fn make_request(path: &str) -> String {
    format!("GET /{} HTTP/1.1\r\nHost: 127.0.0.1\r\nConnection: close\r\n\r\n", path)
}

/// Spawn a per-iteration server, send one request, return the response body.
fn bench_single_request(path: &str) -> String {
    let rt = tokio::runtime::Runtime::new().unwrap();
    rt.block_on(async {
        let config = asperger::asp::config::Config {
            host: "127.0.0.1".to_string(),
            port: 0,
            folder: "asp_files".to_string(),
            program: None,
            enable_directory_listing: false,
        };
        let server = asperger::asp::server::AspServer::new(config);
        let handler = Arc::clone(&server.handler_chain);
        let store = Arc::clone(&server.store);
        let listener = Arc::new(tokio::net::TcpListener::bind("127.0.0.1:0").await.unwrap());
        let addr = listener.local_addr().unwrap();
        let default_doc = "index.asp".to_string();

        let accept_handle = tokio::spawn(async move {
            let (mut stream, _) = listener.accept().await.unwrap();
            let _ = asperger::asp::server::AspServer::handle_connection(
                &handler, &mut stream, "asp_files", &default_doc, &store, false,
            ).await;
        });

        let body = send_and_receive(path, addr).await;
        accept_handle.await.unwrap();
        body
    })
}

/// Spawn N accept tasks, each listening on `listener` and handling one connection.
async fn spawn_accept_tasks(listener: Arc<tokio::net::TcpListener>, handler: Arc<dyn asperger::asp::handler::Handler + Send + Sync>, store: Arc<asperger::vbscript::store::Store>, default_doc: String, count: usize) -> Vec<tokio::task::JoinHandle<()>> {
    let mut handles = Vec::with_capacity(count);
    for _ in 0..count {
        let l = Arc::clone(&listener);
        let h = Arc::clone(&handler);
        let s = Arc::clone(&store);
        let dd = default_doc.clone();
        handles.push(tokio::spawn(async move {
            let (mut stream, _) = l.accept().await.unwrap();
            let _ = asperger::asp::server::AspServer::handle_connection(
                &h, &mut stream, "asp_files", &dd, &s, false,
            ).await;
        }));
    }
    handles
}

async fn send_and_receive(path: &str, addr: std::net::SocketAddr) -> String {
    let mut client = tokio::net::TcpStream::connect(addr).await.unwrap();
    let request = make_request(path);
    use tokio::io::AsyncWriteExt;
    client.write_all(request.as_bytes()).await.unwrap();
    client.shutdown().await.unwrap();

    use tokio::io::AsyncBufReadExt;
    let mut reader = tokio::io::BufReader::new(&mut client);
    let mut response = String::new();
    let mut line = String::new();
    let _ = reader.read_line(&mut line).await;
    response.push_str(&line);
    loop {
        line.clear();
        let n = reader.read_line(&mut line).await.unwrap();
        if n <= 2 {
            break;
        }
        response.push_str(&line);
    }
    let mut body = String::new();
    use tokio::io::AsyncReadExt;
    let _ = reader.read_to_string(&mut body).await;
    response.push_str(&body);
    response
}

/// Simulate V virtual users making R requests each concurrently.
/// Returns the total elapsed time for all requests to complete.
fn run_virtual_users(path: &str, users: usize, requests_per_user: usize) -> f64 {
    let rt = tokio::runtime::Runtime::new().unwrap();
    rt.block_on(async {
        let config = asperger::asp::config::Config {
            host: "127.0.0.1".to_string(),
            port: 0,
            folder: "asp_files".to_string(),
            program: None,
            enable_directory_listing: false,
        };
        let server = asperger::asp::server::AspServer::new(config);
        let handler = Arc::clone(&server.handler_chain);
        let store = Arc::clone(&server.store);
        let listener = Arc::new(tokio::net::TcpListener::bind("127.0.0.1:0").await.unwrap());
        let addr = listener.local_addr().unwrap();
        let default_doc = "index.asp".to_string();

        let total = users * requests_per_user;
        let _accept_handles = spawn_accept_tasks(
            Arc::clone(&listener),
            Arc::clone(&handler),
            Arc::clone(&store),
            default_doc.clone(),
            total,
        ).await;

        let start = Instant::now();

        let mut user_handles = Vec::with_capacity(users);
        for _ in 0..users {
            let path = path.to_string();
            user_handles.push(tokio::spawn(async move {
                for _ in 0..requests_per_user {
                    let _body = send_and_receive(&path, addr).await;
                }
            }));
        }

        for h in user_handles {
            h.await.unwrap();
        }

        let elapsed = start.elapsed().as_secs_f64();
        elapsed
    })
}

fn bench_empty_html(c: &mut Criterion) {
    c.bench_function("http_empty_html", |b| {
        b.iter(|| {
            let _resp = bench_single_request("bench_empty.asp");
        });
    });
}

fn bench_small_echo(c: &mut Criterion) {
    c.bench_function("http_small_echo", |b| {
        b.iter(|| {
            let resp = bench_single_request("bench_small_echo.asp");
            assert!(resp.contains("42"));
        });
    });
}

fn bench_heavy_math(c: &mut Criterion) {
    c.bench_function("http_heavy_math", |b| {
        b.iter(|| {
            let resp = bench_single_request("bench_heavy_math.asp");
            assert!(resp.contains("Result:"));
        });
    });
}

fn bench_string_concat(c: &mut Criterion) {
    c.bench_function("http_string_concat", |b| {
        b.iter(|| {
            let resp = bench_single_request("bench_string_concat.asp");
            assert!(resp.contains("Len: 5000"));
        });
    });
}

fn bench_session(c: &mut Criterion) {
    c.bench_function("http_session", |b| {
        b.iter(|| {
            let resp = bench_single_request("bench_session.asp");
            assert!(resp.contains("200 OK"));
        });
    });
}

// ── Virtual User Benchmarks ─────────────────────────────────────────────────
// Each benchmark simulates V virtual users making R requests each concurrently.
// Each iteration runs one complete batch; criterion estimates iterations from warm-up.
const REQUESTS_PER_USER: usize = 5;

macro_rules! vu_bench {
    ($name:ident, $path:expr, $users:expr) => {
        fn $name(c: &mut Criterion) {
            c.bench_function(concat!("vu/", stringify!($name)), |b| {
                b.iter(|| {
                    let _elapsed = criterion::black_box(
                        run_virtual_users($path, $users, REQUESTS_PER_USER),
                    );
                });
            });
        }
    };
}

vu_bench!(vu4_empty, "bench_empty.asp", 4);
vu_bench!(vu8_empty, "bench_empty.asp", 8);
vu_bench!(vu16_empty, "bench_empty.asp", 16);
vu_bench!(vu4_math, "bench_heavy_math.asp", 4);
vu_bench!(vu8_math, "bench_heavy_math.asp", 8);
vu_bench!(vu16_math, "bench_heavy_math.asp", 16);
vu_bench!(vu4_session, "bench_session.asp", 4);

criterion_group!(
    name = http;
    config = Criterion::default().sample_size(50).measurement_time(std::time::Duration::from_secs(10));
    targets =
        bench_empty_html,
        bench_small_echo,
        bench_heavy_math,
        bench_string_concat,
        bench_session,
);

criterion_group!(
    name = vu;
    config = Criterion::default().sample_size(20).measurement_time(std::time::Duration::from_secs(20));
    targets =
        vu4_empty,
        vu8_empty,
        vu16_empty,
        vu4_math,
        vu8_math,
        vu16_math,
        vu4_session,
);

criterion_main!(http, vu);
