fn main() {
    // Desktop OAuth secrets are public-client configuration: Google does not
    // treat them as confidential, but keep the local source file out of git.
    let secret_path = std::path::Path::new("gdrive-client-secret.txt");
    println!("cargo:rerun-if-changed={}", secret_path.display());
    if let Ok(secret) = std::fs::read_to_string(secret_path) {
        let secret = secret.trim();
        if !secret.is_empty() {
            println!("cargo:rustc-env=CHURCH_PRESENTER_GDRIVE_CLIENT_SECRET={secret}");
        }
    }
    tauri_build::build()
}
