use anyhow::Result;
use clap::{Parser, Subcommand};
use outlook_web::{api::Client, browser, config};

#[derive(Parser)]
#[command(name = "outlook-web")]
#[command(about = "CLI to access Outlook Web via browser automation")]
struct Cli {
    /// Output as JSON
    #[arg(long, global = true)]
    json: bool,

    /// Chrome debugging port (default: 9222)
    #[arg(long, global = true)]
    port: Option<u16>,

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
    /// List junk/spam folder messages
    ListSpam {
        /// Maximum number of messages
        #[arg(short = 'n', long, default_value = "20")]
        max: u32,
    },
    /// Read a specific message by ID
    Read {
        /// Message ID
        id: String,
    },
    /// Archive a message
    Archive {
        /// Message ID
        id: String,
    },
    /// Delete a message
    Delete {
        /// Message ID
        id: String,
    },
    /// Mark as spam
    Spam {
        /// Message ID
        id: String,
    },
    /// Add label/category to message
    Label {
        /// Message ID
        id: String,
        /// Label to add
        label: String,
    },
    /// Remove label/category from message
    Unlabel {
        /// Message ID
        id: String,
        /// Label to remove
        label: String,
    },
    /// List available labels/categories
    Labels,
    /// Move message from Junk to Inbox
    Unspam {
        /// Message ID
        id: String,
    },
    /// Mark message as read
    MarkRead {
        /// Message ID
        id: String,
    },
    /// Mark message as unread
    MarkUnread {
        /// Message ID
        id: String,
    },
    /// Remove all labels from message
    ClearLabels {
        /// Message ID
        id: String,
    },
    /// Unsubscribe from a mailing list (opens unsubscribe link)
    Unsubscribe {
        /// Message ID
        id: String,
    },
    /// Sync labels: show categories used on messages but not in master list
    SyncLabels,
    /// Test connection to browser
    Test,
    /// Inspect DOM to find selectors
    Inspect,
    /// Configure settings
    Config {
        /// Set default port
        #[arg(long)]
        port: Option<u16>,
    },
}

#[tokio::main]
async fn main() -> Result<()> {
    let cli = Cli::parse();
    let cfg = config::load_config()?;
    let port = cli.port.unwrap_or_else(|| cfg.port());

    match cli.command {
        Commands::Config { port: new_port } => {
            let mut cfg = config::load_config()?;
            if let Some(p) = new_port {
                cfg.port = Some(p);
                config::save_config(&cfg)?;
                println!("Port set to: {}", p);
            } else {
                println!("Current settings:");
                println!("  port: {}", cfg.port());
            }
        }
        Commands::List { max } => {
            let client = Client::new(port);
            let messages = client.list_messages(max).await?;

            if cli.json {
                println!("{}", serde_json::to_string(&messages)?);
            } else if messages.is_empty() {
                println!("No messages found. Make sure Outlook inbox is visible.");
            } else {
                for msg in &messages {
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
        }
        Commands::ListSpam { max } => {
            let client = Client::new(port);
            let messages = client.list_spam(max).await?;

            if cli.json {
                println!("{}", serde_json::to_string(&messages)?);
            } else if messages.is_empty() {
                println!("No spam messages found.");
            } else {
                for msg in &messages {
                    let from = msg.from.as_deref().unwrap_or("Unknown");
                    let subject = msg.subject.as_deref().unwrap_or("(no subject)");
                    println!("{} | {} | {}", msg.id, from, subject);
                }
            }
        }
        Commands::Read { id } => {
            let client = Client::new(port);
            let msg = client.get_message(&id).await?;

            if cli.json {
                println!("{}", serde_json::to_string(&msg)?);
            } else {
                println!("From: {}", msg.from.as_deref().unwrap_or("Unknown"));
                println!(
                    "Subject: {}",
                    msg.subject.as_deref().unwrap_or("(no subject)")
                );
                if !msg.labels.is_empty() {
                    println!("Labels: {}", msg.labels.join(", "));
                }
                println!("---");
                println!("{}", msg.body.as_deref().unwrap_or(""));
            }
        }
        Commands::Archive { id } => {
            let client = Client::new(port);
            client.archive(&id).await?;
            println!("Archived: {}", id);
        }
        Commands::Delete { id } => {
            let client = Client::new(port);
            client.trash(&id).await?;
            println!("Deleted: {}", id);
        }
        Commands::Spam { id } => {
            let client = Client::new(port);
            client.mark_spam(&id).await?;
            println!("Marked as spam: {}", id);
        }
        Commands::Label { id, label } => {
            let client = Client::new(port);
            client.add_label(&id, &label).await?;
            println!("Added label '{}' to: {}", label, id);
        }
        Commands::Unlabel { id, label } => {
            let client = Client::new(port);
            client.remove_label(&id, &label).await?;
            println!("Removed label '{}' from: {}", label, id);
        }
        Commands::Labels => {
            let client = Client::new(port);
            let labels = client.list_labels().await?;
            if cli.json {
                println!("{}", serde_json::to_string(&labels)?);
            } else {
                for label in &labels {
                    println!("{}", label);
                }
            }
        }
        Commands::Unspam { id } => {
            let client = Client::new(port);
            client.unspam(&id).await?;
            println!("Moved to inbox: {}", id);
        }
        Commands::MarkRead { id } => {
            let client = Client::new(port);
            client.mark_read(&id).await?;
            println!("Marked as read: {}", id);
        }
        Commands::MarkUnread { id } => {
            let client = Client::new(port);
            client.mark_unread(&id).await?;
            println!("Marked as unread: {}", id);
        }
        Commands::ClearLabels { id } => {
            let client = Client::new(port);
            client.clear_labels(&id).await?;
            println!("Cleared labels from: {}", id);
        }
        Commands::Unsubscribe { id } => {
            let client = Client::new(port);
            if let Some(url) = client.get_unsubscribe_url(&id).await? {
                println!("Opening unsubscribe link: {}", url);
                open::that(&url)?;
            } else {
                anyhow::bail!("No unsubscribe link found in message");
            }
        }
        Commands::SyncLabels => {
            let client = Client::new(port);
            let known_labels = client.list_labels().await?;
            let known_set: std::collections::HashSet<String> =
                known_labels.into_iter().map(|l| l.to_lowercase()).collect();

            // Scan messages for categories
            let messages = client.list_messages(100).await?;
            let mut found: std::collections::HashSet<String> = std::collections::HashSet::new();

            for msg in &messages {
                for label in &msg.labels {
                    if !known_set.contains(&label.to_lowercase()) {
                        found.insert(label.clone());
                    }
                }
            }

            if found.is_empty() {
                println!("All categories are in the known list.");
            } else {
                println!("Categories found on messages but not in known list:");
                for label in found {
                    println!("  - {}", label);
                }
                println!("\nAdd these to list_labels() in src/api.rs");
            }
        }
        Commands::Test => {
            test_connection(port).await?;
        }
        Commands::Inspect => {
            inspect_dom(port).await?;
        }
    }

    Ok(())
}

async fn test_connection(port: u16) -> Result<()> {
    let browser_instance = browser::connect_or_start_browser(port).await?;
    let pages = browser_instance.pages().await?;

    println!("Connected to browser successfully!");
    println!("Found {} pages:", pages.len());

    let timeout = std::time::Duration::from_secs(2);
    for page in &pages {
        let url_result = tokio::time::timeout(timeout, page.url()).await;
        let url = match url_result {
            Ok(Ok(Some(u))) => u,
            Ok(Ok(None)) => "(no url)".to_string(),
            Ok(Err(_)) => "(error getting url)".to_string(),
            Err(_) => "(timeout getting url)".to_string(),
        };
        let is_outlook = url.contains("outlook");
        let marker = if is_outlook { " <-- Outlook" } else { "" };
        let title = tokio::time::timeout(timeout, page.evaluate("document.title"))
            .await
            .ok()
            .and_then(|r| r.ok())
            .and_then(|r| r.into_value::<String>().ok())
            .unwrap_or_else(|| "(no title)".to_string());
        println!("  {} [{}]{}", title, url, marker);
    }

    match browser::find_outlook_page(&browser_instance).await {
        Ok(page) => {
            let url = tokio::time::timeout(timeout, page.url())
                .await
                .ok()
                .and_then(|r| r.ok())
                .flatten()
                .unwrap_or_default();
            println!("\nOutlook tab found: {}", url);
        }
        Err(_) => {
            println!("\nNo Outlook tab found. Open Outlook in the browser.");
        }
    }

    Ok(())
}

async fn inspect_dom(port: u16) -> Result<()> {
    let info = outlook_web::inspect::inspect_dom(port).await?;
    println!("{}", info);
    Ok(())
}
