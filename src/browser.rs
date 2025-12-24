use anyhow::{anyhow, Context, Result};
use chromiumoxide::browser::Browser;
use futures::StreamExt;
use serde::Deserialize;
use std::process::{Command, Stdio};

#[derive(Debug, Deserialize)]
struct BrowserVersion {
    #[serde(rename = "webSocketDebuggerUrl")]
    ws_url: String,
}

/// Browser executable paths to try in order of preference
const BROWSER_CANDIDATES: &[(&str, &[&str])] = &[
    ("Vivaldi", &[
        // Linux
        "/usr/bin/vivaldi",
        "/usr/bin/vivaldi-stable",
        "/opt/vivaldi/vivaldi",
        // macOS
        "/Applications/Vivaldi.app/Contents/MacOS/Vivaldi",
    ]),
    ("Chromium", &[
        // Linux
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/snap/bin/chromium",
        // macOS
        "/Applications/Chromium.app/Contents/MacOS/Chromium",
    ]),
    ("Chrome", &[
        // Linux
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        "/opt/google/chrome/google-chrome",
        // macOS
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    ]),
];

/// Find the first available browser executable
fn find_browser() -> Option<(&'static str, &'static str)> {
    for (name, paths) in BROWSER_CANDIDATES {
        for path in *paths {
            if std::path::Path::new(path).exists() {
                return Some((name, path));
            }
        }
    }
    None
}

/// Start a browser with remote debugging enabled
pub fn start_browser(port: u16) -> Result<()> {
    let (name, path) = find_browser().ok_or_else(|| {
        anyhow!(
            "No supported browser found. Install one of: Vivaldi, Chromium, or Chrome"
        )
    })?;

    eprintln!("Starting {} with remote debugging on port {}...", name, port);

    Command::new(path)
        .arg(format!("--remote-debugging-port={}", port))
        .arg("https://outlook.office.com/mail/")
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .context(format!("Failed to start {}", name))?;

    Ok(())
}

pub async fn get_browser_ws_url(port: u16) -> Result<String> {
    let url = format!("http://127.0.0.1:{}/json/version", port);
    let resp: BrowserVersion = reqwest::get(&url)
        .await
        .context(format!(
            "Failed to connect to browser on port {}",
            port
        ))?
        .json()
        .await?;
    Ok(resp.ws_url)
}

/// Try to connect to browser, starting one if needed
pub async fn connect_or_start_browser(port: u16) -> Result<Browser> {
    // First try to connect to existing browser
    if let Ok(browser) = connect_browser(port).await {
        return Ok(browser);
    }

    // No browser running, start one
    start_browser(port)?;

    // Wait for browser to start and retry connection
    for attempt in 1..=10 {
        tokio::time::sleep(tokio::time::Duration::from_millis(500)).await;
        if let Ok(browser) = connect_browser(port).await {
            return Ok(browser);
        }
        if attempt == 10 {
            return Err(anyhow!("Browser started but failed to connect after 5 seconds"));
        }
    }

    unreachable!()
}

pub async fn connect_browser(port: u16) -> Result<Browser> {
    let ws_url = get_browser_ws_url(port).await?;

    let (mut browser, mut handler) = Browser::connect(&ws_url)
        .await
        .context("Failed to connect to browser via WebSocket")?;

    tokio::spawn(async move {
        while handler.next().await.is_some() {}
    });

    // Fetch existing targets (pages that were open before we connected)
    browser.fetch_targets().await?;

    // Give pages a moment to be ready
    tokio::time::sleep(tokio::time::Duration::from_millis(100)).await;

    Ok(browser)
}

pub async fn find_outlook_page(browser: &Browser) -> Result<chromiumoxide::Page> {
    let pages = browser.pages().await?;
    let timeout = std::time::Duration::from_secs(2);

    for page in pages {
        let url_result = tokio::time::timeout(timeout, page.url()).await;
        if let Ok(Ok(Some(u))) = url_result {
            if u.contains("outlook.office.com")
                || u.contains("outlook.live.com")
                || u.contains("outlook.office365.com")
            {
                return Ok(page);
            }
        }
    }

    Err(anyhow!("No Outlook tab found. Open Outlook in the browser first."))
}

/// Navigate to inbox if not already there
pub async fn navigate_to_inbox(page: &chromiumoxide::Page) -> Result<()> {
    let script = r#"
        (() => {
            const url = window.location.href;
            if (url.includes('/inbox') || url.match(/\/mail\/\d+\/?($|id\/)/)) return 'already';
            const match = url.match(/(https:\/\/outlook\.[^\/]+\/mail\/\d+\/)/);
            if (match) {
                window.location.href = match[1] + 'inbox';
                return 'navigating';
            }
            return 'already';
        })()
    "#;

    let result = page.evaluate(script).await?;
    if result.into_value::<String>().unwrap_or_default() == "navigating" {
        tokio::time::sleep(tokio::time::Duration::from_secs(2)).await;
    }

    Ok(())
}

/// Navigate to junk/spam folder
pub async fn navigate_to_junk(page: &chromiumoxide::Page) -> Result<()> {
    let script = r#"
        (() => {
            const url = window.location.href;
            if (url.includes('/junkemail')) return 'already';
            const match = url.match(/(https:\/\/outlook\.[^\/]+\/mail\/\d+\/)/);
            if (match) {
                window.location.href = match[1] + 'junkemail';
                return 'navigating';
            }
            return 'failed';
        })()
    "#;

    let result = page.evaluate(script).await?;
    let status = result.into_value::<String>().unwrap_or_default();

    if status == "failed" {
        anyhow::bail!("Failed to parse Outlook URL for navigation");
    }
    if status == "navigating" {
        tokio::time::sleep(tokio::time::Duration::from_secs(2)).await;
    }

    Ok(())
}
