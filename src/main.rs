use anyhow::{anyhow, Context, Result};
use chromiumoxide::browser::Browser;
use clap::{Parser, Subcommand};
use futures::StreamExt;
use serde::{Deserialize, Serialize};

#[derive(Parser)]
#[command(name = "outlook-web")]
#[command(about = "CLI to access Outlook Web via browser automation")]
struct Cli {
    /// Output as JSON
    #[arg(long, global = true)]
    json: bool,

    /// Chrome debugging port (default: 9222)
    #[arg(long, global = true, default_value = "9222")]
    port: u16,

    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// List inbox messages
    List {
        /// Maximum number of messages
        #[arg(short = 'n', long, default_value = "20")]
        max: u32,
    },
    /// Read a specific message by ID
    Read {
        /// Message ID
        id: String,
    },
    /// Test connection to browser
    Test,
    /// Inspect DOM to find selectors
    Inspect,
}

#[derive(Debug, Serialize, Deserialize)]
struct Message {
    id: String,
    subject: Option<String>,
    from: Option<String>,
    preview: Option<String>,
    #[serde(default)]
    labels: Vec<String>,
    #[serde(rename = "isUnread", default)]
    is_unread: bool,
}

#[derive(Debug, Deserialize)]
struct BrowserVersion {
    #[serde(rename = "webSocketDebuggerUrl")]
    ws_url: String,
}

async fn get_browser_ws_url(port: u16) -> Result<String> {
    let url = format!("http://127.0.0.1:{}/json/version", port);
    let resp: BrowserVersion = reqwest::get(&url)
        .await
        .context(format!(
            "Failed to connect to browser on port {}. Start browser with:\n  \
            vivaldi --remote-debugging-port={}",
            port, port
        ))?
        .json()
        .await?;
    Ok(resp.ws_url)
}

async fn connect_browser(port: u16) -> Result<Browser> {
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

async fn find_outlook_page(browser: &Browser) -> Result<chromiumoxide::Page> {
    let pages = browser.pages().await?;

    for page in pages {
        if let Ok(url) = page.url().await {
            if let Some(u) = url {
                if u.contains("outlook.office.com")
                    || u.contains("outlook.live.com")
                    || u.contains("outlook.office365.com")
                {
                    return Ok(page);
                }
            }
        }
    }

    Err(anyhow!("No Outlook tab found. Open Outlook in the browser first."))
}

async fn list_messages(port: u16, max: u32, json_output: bool) -> Result<()> {
    let browser = connect_browser(port).await?;
    let page = find_outlook_page(&browser).await?;

    let script = r#"
        (() => {
            const messages = [];
            const knownLabels = ['Classified', 'Urgent', 'Security', 'Account', 'Important', 'Personal', 'Private', 'Work', 'Family'];
            const items = document.querySelectorAll('[data-convid]');
            items.forEach(item => {
                const id = item.getAttribute('data-convid');
                const ariaLabel = item.getAttribute('aria-label') || '';

                // Extract labels from list item
                let labels = [];
                item.querySelectorAll('button, span').forEach(el => {
                    const text = el.textContent?.trim();
                    if (knownLabels.includes(text) && !labels.includes(text)) {
                        labels.push(text);
                    }
                });

                // aria-label format: "Sender Subject Time Preview..."
                let from = '';
                let subject = '';
                let preview = '';

                // Try to find sender name - usually at start, before subject line
                const match = ariaLabel.match(/^(.*?)\s+(Re:|Fw:|New\s|Your\s|Action|Urgent|Important|Microsoft|Amazon|Google|Welcome)/i);
                if (match) {
                    from = match[1].trim();
                    const rest = ariaLabel.substring(match[1].length).trim();
                    const timeParts = rest.split(/\d{1,2}:\d{2}|\d{4}-\d{2}-\d{2}/);
                    subject = timeParts[0]?.trim() || '';
                    preview = timeParts.slice(1).join(' ').trim();
                } else {
                    // Fallback: split on double space or assume first part is sender
                    const parts = ariaLabel.split(/\s{2,}/);
                    from = parts[0] || '';
                    subject = parts[1] || '';
                    preview = parts.slice(2).join(' ');
                }

                // Clean subject - remove labels and Unread marker
                subject = subject.replace(/^Unread\s*/i, '');
                labels.forEach(label => { subject = subject.replace(label, ''); });
                subject = subject.trim();

                // Check for Unread marker
                const isUnread = ariaLabel.toLowerCase().includes('unread');

                if (id) {
                    messages.push({ id, subject, from, preview, labels, isUnread });
                }
            });
            return JSON.stringify(messages);
        })()
    "#;

    let result = page.evaluate(script).await?;
    let messages_str = result.into_value::<String>().unwrap_or_default();
    let parsed: Vec<Message> = serde_json::from_str(&messages_str).unwrap_or_default();

    if json_output {
        println!("{}", serde_json::to_string(&parsed)?);
    } else if parsed.is_empty() {
        let url = page.url().await?.unwrap_or_default();
        println!("No messages found. Make sure Outlook inbox is visible.");
        println!("Current tab: {}", url);
    } else {
        for msg in parsed.iter().take(max as usize) {
            let from = msg.from.as_deref().unwrap_or("Unknown");
            let subject = msg.subject.as_deref().unwrap_or("(no subject)");
            let unread = if msg.is_unread { "*" } else { " " };
            let labels = if msg.labels.is_empty() {
                String::new()
            } else {
                format!(" [{}]", msg.labels.join(", "))
            };
            println!("{}{} | {} | {}{}", unread, msg.id, from, subject, labels);
        }
    }

    Ok(())
}

async fn test_connection(port: u16) -> Result<()> {
    let browser = connect_browser(port).await?;
    let pages = browser.pages().await?;

    println!("Connected to browser successfully!");
    println!("Found {} pages:", pages.len());

    for page in &pages {
        if let Ok(Some(url)) = page.url().await {
            let is_outlook = url.contains("outlook");
            let marker = if is_outlook { " <-- Outlook" } else { "" };
            let title = page
                .evaluate("document.title")
                .await
                .ok()
                .and_then(|r| r.into_value::<String>().ok())
                .unwrap_or_default();
            println!("  {}{}", title, marker);
        }
    }

    match find_outlook_page(&browser).await {
        Ok(page) => {
            let url = page.url().await?.unwrap_or_default();
            println!("\nOutlook tab found: {}", url);
        }
        Err(_) => {
            println!("\nNo Outlook tab found. Open Outlook in the browser.");
        }
    }

    Ok(())
}

async fn inspect_dom(port: u16) -> Result<()> {
    let browser = connect_browser(port).await?;
    let page = find_outlook_page(&browser).await?;

    let script = r#"
        (() => {
            const info = {};

            // Find message reading pane
            const readingPane = document.querySelector('[aria-label*="Message body"], [class*="ReadingPane"], [class*="readingPane"], div[role="document"]');
            if (readingPane) {
                info.readingPane = {
                    tag: readingPane.tagName,
                    classes: readingPane.className,
                    ariaLabel: readingPane.getAttribute('aria-label'),
                    textPreview: readingPane.innerText?.substring(0, 200)
                };
            }

            // Find subject in reading pane
            const subjects = document.querySelectorAll('h1, h2, [role="heading"], [class*="subject"], [class*="Subject"]');
            info.subjects = Array.from(subjects).slice(0, 5).map(el => ({
                tag: el.tagName,
                classes: el.className,
                text: el.textContent?.trim()?.substring(0, 100)
            }));

            // Find sender info
            const senders = document.querySelectorAll('[class*="sender"], [class*="Sender"], [class*="From"], [class*="from"]');
            info.senders = Array.from(senders).slice(0, 5).map(el => ({
                tag: el.tagName,
                classes: el.className,
                text: el.textContent?.trim()?.substring(0, 100)
            }));

            // Find message body candidates
            const bodies = document.querySelectorAll('[class*="UniqueMessageBody"], [class*="messageBody"], [class*="ItemContent"], div[dir="ltr"]');
            info.bodies = Array.from(bodies).slice(0, 3).map(el => ({
                tag: el.tagName,
                classes: el.className,
                textPreview: el.innerText?.substring(0, 200)
            }));

            // Find message list items
            const listItems = document.querySelectorAll('[data-convid], [role="option"], [class*="listItem"]');
            info.listItemCount = listItems.length;
            if (listItems.length > 0) {
                const first = listItems[0];
                info.firstListItem = {
                    tag: first.tagName,
                    classes: first.className,
                    dataConvid: first.getAttribute('data-convid'),
                    html: first.outerHTML?.substring(0, 500)
                };
            }

            return JSON.stringify(info, null, 2);
        })()
    "#;

    let result = page.evaluate(script).await?;
    let info = result.into_value::<String>().unwrap_or_default();
    println!("{}", info);

    Ok(())
}

async fn read_message(port: u16, id: &str, json_output: bool) -> Result<()> {
    let browser = connect_browser(port).await?;
    let page = find_outlook_page(&browser).await?;

    // Click on message with given ID
    let click_script = format!(
        r#"
        (() => {{
            const item = document.querySelector('[data-convid="{}"], [data-item-index="{}"]');
            if (item) {{
                item.click();
                return true;
            }}
            return false;
        }})()
    "#,
        id, id
    );

    page.evaluate(click_script).await?;
    tokio::time::sleep(tokio::time::Duration::from_secs(2)).await;

    let read_script = r#"
        (() => {
            // Subject and labels - find the reading pane header area
            let subject = '';
            let labels = [];

            // Look for labels in the message header area (they're usually buttons or spans)
            const headerArea = document.querySelector('[class*="ReadingPane"], [class*="ItemHeader"]')
                || document.querySelector('.fui-FluentProviderr4k');
            if (headerArea) {
                // Labels are often in small buttons/spans after the subject
                headerArea.querySelectorAll('button, span').forEach(el => {
                    const text = el.textContent?.trim();
                    // Common Outlook category names
                    if (['Classified', 'Urgent', 'Security', 'Account', 'Important', 'Personal', 'Private', 'Work', 'Family'].includes(text)) {
                        if (!labels.includes(text)) labels.push(text);
                    }
                });
            }

            // Get subject - use the allowTextSelection div but strip known labels
            const subjectEl = document.querySelector('.allowTextSelection, [class*="SubjectLine"]');
            if (subjectEl) {
                subject = subjectEl.textContent?.trim() || '';
                // Remove concatenated labels from end of subject
                labels.forEach(label => {
                    subject = subject.replace(label, '');
                });
                subject = subject.trim();
            }

            // From - try multiple selectors
            let from = '';
            // Look for sender display name
            const senderEl = document.querySelector('[class*="Sender"], [class*="sender"], [class*="From"], button[class*="Persona"]');
            if (senderEl) {
                from = senderEl.textContent?.trim();
            }
            // Try aria-label on selected list item
            if (!from) {
                const selected = document.querySelector('[data-convid][aria-selected="true"]');
                if (selected) {
                    const label = selected.getAttribute('aria-label') || '';
                    // First part before subject keywords is usually sender
                    const match = label.match(/^([^<]+?)(?:\s+(?:Re:|Fw:|New\s|Your\s|Microsoft|Amazon))/i);
                    if (match) from = match[1].trim();
                }
            }

            // Body is in the reading pane
            const bodyEl = document.querySelector('div[role="document"]');
            const body = bodyEl?.innerText?.trim();

            return JSON.stringify({ subject, from, body, labels });
        })()
    "#;

    let result = page.evaluate(read_script).await?;
    let message_str = result.into_value::<String>().unwrap_or_default();

    if json_output {
        println!("{}", message_str);
    } else {
        let msg: serde_json::Value = serde_json::from_str(&message_str)?;
        println!("From: {}", msg["from"].as_str().unwrap_or("Unknown"));
        println!(
            "Subject: {}",
            msg["subject"].as_str().unwrap_or("(no subject)")
        );
        if let Some(labels) = msg["labels"].as_array() {
            if !labels.is_empty() {
                let label_strs: Vec<&str> = labels.iter().filter_map(|l| l.as_str()).collect();
                println!("Labels: {}", label_strs.join(", "));
            }
        }
        println!("---");
        println!("{}", msg["body"].as_str().unwrap_or(""));
    }

    Ok(())
}

#[tokio::main]
async fn main() -> Result<()> {
    let cli = Cli::parse();

    match cli.command {
        Commands::List { max } => {
            list_messages(cli.port, max, cli.json).await?;
        }
        Commands::Read { id } => {
            read_message(cli.port, &id, cli.json).await?;
        }
        Commands::Test => {
            test_connection(cli.port).await?;
        }
        Commands::Inspect => {
            inspect_dom(cli.port).await?;
        }
    }

    Ok(())
}
