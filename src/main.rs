#![windows_subsystem = "windows"]
use anyhow::Result;

slint::include_modules!();

mod attendance;

fn main() -> Result<()> {
    let app = attendance::App::new();
    app.run()
}
