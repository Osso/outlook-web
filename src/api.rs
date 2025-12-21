use crate::browser::{connect_or_start_browser, find_outlook_page};
use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Message {
    pub id: String,
    pub subject: Option<String>,
    pub from: Option<String>,
    pub body: Option<String>,
    pub preview: Option<String>,
    #[serde(default)]
    pub labels: Vec<String>,
    #[serde(rename = "isUnread", default)]
    pub is_unread: bool,
}

pub struct Client {
    port: u16,
}

impl Client {
    pub fn new(port: u16) -> Self {
        Self { port }
    }

    pub async fn list_messages(&self, max: u32) -> Result<Vec<Message>> {
        let browser = connect_or_start_browser(self.port).await?;
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

                    // aria-label format: "[Unread] Sender Subject Time Preview..."
                    let from = '';
                    let subject = '';
                    let preview = '';

                    // Remove "Unread" prefix if present
                    let label = ariaLabel.replace(/^Unread\s+/i, '').trim();

                    // Common subject starters to find where sender ends
                    const subjectStarters = /\s+(Re:|Fw:|FW:|RE:|New\s|Your\s|Action\s|Welcome|Microsoft|Amazon|Google|Apple|Thanks|Thank\s|Confirm|Verify|Update|Alert|Notice|Reminder|Invoice|Order|Shipping|Delivery)/i;
                    const match = label.match(subjectStarters);

                    if (match) {
                        from = label.substring(0, match.index).trim();
                        const rest = label.substring(match.index).trim();
                        // Split on time pattern to separate subject from preview
                        const timeParts = rest.split(/\s+\d{1,2}:\d{2}\s+|\s+\d{4}-\d{2}-\d{2}\s+/);
                        subject = timeParts[0]?.trim() || '';
                        preview = timeParts.slice(1).join(' ').trim();
                    } else {
                        // Fallback: look for time pattern to split
                        const timeSplit = label.split(/\s+\d{1,2}:\d{2}\s+|\s+\d{4}-\d{2}-\d{2}\s+/);
                        if (timeSplit.length > 1) {
                            // First part has sender + subject, need to split by common patterns
                            const firstPart = timeSplit[0];
                            // Try splitting on < which often separates display name from email
                            const emailMatch = firstPart.match(/^(.+?)<[^>]+>\s*(.*)/);
                            if (emailMatch) {
                                from = emailMatch[1].trim();
                                subject = emailMatch[2].trim();
                            } else {
                                // Assume first few words are sender
                                const words = firstPart.split(/\s+/);
                                from = words.slice(0, 3).join(' ');
                                subject = words.slice(3).join(' ');
                            }
                            preview = timeSplit.slice(1).join(' ').trim();
                        } else {
                            from = label;
                        }
                    }

                    // Clean subject - remove labels
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
        let mut parsed: Vec<Message> = serde_json::from_str(&messages_str).unwrap_or_default();
        parsed.truncate(max as usize);
        Ok(parsed)
    }

    pub async fn get_message(&self, id: &str) -> Result<Message> {
        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        // Click on message with given ID
        let click_script = format!(
            r#"
            (() => {{
                const item = document.querySelector('[data-convid="{}"]');
                if (item) {{
                    item.click();
                    return true;
                }}
                return false;
            }})()
        "#,
            id
        );

        let clicked = page.evaluate(click_script).await?;
        if !clicked.into_value::<bool>().unwrap_or(false) {
            anyhow::bail!("Message not found: {}", id);
        }

        tokio::time::sleep(tokio::time::Duration::from_secs(2)).await;

        let read_script = r#"
            (() => {
                const knownLabels = ['Classified', 'Urgent', 'Security', 'Account', 'Important', 'Personal', 'Private', 'Work', 'Family'];
                let subject = '';
                let labels = [];

                // Look for labels in the message header area
                const headerArea = document.querySelector('[class*="ReadingPane"], [class*="ItemHeader"]')
                    || document.querySelector('.fui-FluentProviderr4k');
                if (headerArea) {
                    headerArea.querySelectorAll('button, span').forEach(el => {
                        const text = el.textContent?.trim();
                        if (knownLabels.includes(text) && !labels.includes(text)) {
                            labels.push(text);
                        }
                    });
                }

                // Get subject
                const subjectEl = document.querySelector('.allowTextSelection, [class*="SubjectLine"]');
                if (subjectEl) {
                    subject = subjectEl.textContent?.trim() || '';
                    labels.forEach(label => {
                        subject = subject.replace(label, '');
                    });
                    subject = subject.trim();
                }

                // From
                let from = '';
                const senderEl = document.querySelector('[class*="Sender"], [class*="sender"], [class*="From"], button[class*="Persona"]');
                if (senderEl) {
                    from = senderEl.textContent?.trim();
                }
                if (!from) {
                    const selected = document.querySelector('[data-convid][aria-selected="true"]');
                    if (selected) {
                        const label = selected.getAttribute('aria-label') || '';
                        const match = label.match(/^([^<]+?)(?:\s+(?:Re:|Fw:|New\s|Your\s|Microsoft|Amazon))/i);
                        if (match) from = match[1].trim();
                    }
                }

                // Body
                const bodyEl = document.querySelector('div[role="document"]');
                const body = bodyEl?.innerText?.trim();

                // Get ID from selected item
                const selected = document.querySelector('[data-convid][aria-selected="true"]');
                const id = selected?.getAttribute('data-convid') || '';

                return JSON.stringify({ id, subject, from, body, labels, isUnread: false });
            })()
        "#;

        let result = page.evaluate(read_script).await?;
        let message_str = result.into_value::<String>().unwrap_or_default();
        let message: Message = serde_json::from_str(&message_str)
            .context("Failed to parse message")?;
        Ok(message)
    }

    pub async fn add_label(&self, id: &str, label: &str) -> Result<()> {
        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        // Right-click to open context menu, then categorize
        let script = format!(
            r#"
            (async () => {{
                const item = document.querySelector('[data-convid="{}"]');
                if (!item) return 'not_found';

                // Right-click on the item
                const event = new MouseEvent('contextmenu', {{
                    bubbles: true,
                    cancelable: true,
                    view: window,
                    button: 2
                }});
                item.dispatchEvent(event);

                await new Promise(r => setTimeout(r, 500));

                // Find and click "Categorize" menu item
                const menuItems = document.querySelectorAll('[role="menuitem"], [role="menuitemcheckbox"]');
                for (const mi of menuItems) {{
                    if (mi.textContent?.includes('Categorize')) {{
                        mi.click();
                        await new Promise(r => setTimeout(r, 500));
                        break;
                    }}
                }}

                // Find and click the specific category
                const categoryItems = document.querySelectorAll('[role="menuitemcheckbox"], [role="menuitem"]');
                for (const ci of categoryItems) {{
                    if (ci.textContent?.trim() === '{}') {{
                        ci.click();
                        return 'success';
                    }}
                }}

                // Category might need to be created - look for "New category" or similar
                return 'category_not_found';
            }})()
        "#,
            id, label
        );

        let result = page.evaluate(script).await?;
        let status = result.into_value::<String>().unwrap_or_default();

        match status.as_str() {
            "success" => Ok(()),
            "not_found" => anyhow::bail!("Message not found: {}", id),
            "category_not_found" => anyhow::bail!("Category not found: {}. Create it in Outlook first.", label),
            _ => anyhow::bail!("Failed to add label: {}", status),
        }
    }

    pub async fn archive(&self, id: &str) -> Result<()> {
        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        // Select message and press 'e' for archive (Outlook shortcut)
        let click_script = format!(
            r#"
            (() => {{
                const item = document.querySelector('[data-convid="{}"]');
                if (!item) return false;
                item.click();
                return true;
            }})()
        "#,
            id
        );

        let clicked = page.evaluate(click_script).await?;
        if !clicked.into_value::<bool>().unwrap_or(false) {
            anyhow::bail!("Message not found: {}", id);
        }

        tokio::time::sleep(tokio::time::Duration::from_millis(300)).await;

        // Press 'e' for archive
        page.evaluate("document.dispatchEvent(new KeyboardEvent('keydown', { key: 'e', code: 'KeyE', bubbles: true }))").await?;

        tokio::time::sleep(tokio::time::Duration::from_millis(500)).await;
        Ok(())
    }

    pub async fn trash(&self, id: &str) -> Result<()> {
        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        // Select message and press Delete
        let click_script = format!(
            r#"
            (() => {{
                const item = document.querySelector('[data-convid="{}"]');
                if (!item) return false;
                item.click();
                return true;
            }})()
        "#,
            id
        );

        let clicked = page.evaluate(click_script).await?;
        if !clicked.into_value::<bool>().unwrap_or(false) {
            anyhow::bail!("Message not found: {}", id);
        }

        tokio::time::sleep(tokio::time::Duration::from_millis(300)).await;

        // Press Delete key
        page.evaluate("document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', code: 'Delete', bubbles: true }))").await?;

        tokio::time::sleep(tokio::time::Duration::from_millis(500)).await;
        Ok(())
    }

    pub async fn mark_spam(&self, id: &str) -> Result<()> {
        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        // Right-click to open context menu, then mark as junk
        let script = format!(
            r#"
            (async () => {{
                const item = document.querySelector('[data-convid="{}"]');
                if (!item) return 'not_found';

                const event = new MouseEvent('contextmenu', {{
                    bubbles: true,
                    cancelable: true,
                    view: window,
                    button: 2
                }});
                item.dispatchEvent(event);

                await new Promise(r => setTimeout(r, 500));

                // Find and click "Report" or "Junk" menu item
                const menuItems = document.querySelectorAll('[role="menuitem"]');
                for (const mi of menuItems) {{
                    const text = mi.textContent?.toLowerCase() || '';
                    if (text.includes('junk') || text.includes('spam') || text.includes('report')) {{
                        mi.click();
                        await new Promise(r => setTimeout(r, 500));

                        // Look for "Junk" submenu option
                        const subItems = document.querySelectorAll('[role="menuitem"]');
                        for (const si of subItems) {{
                            if (si.textContent?.toLowerCase().includes('junk')) {{
                                si.click();
                                return 'success';
                            }}
                        }}
                        return 'success';
                    }}
                }}

                return 'menu_not_found';
            }})()
        "#,
            id
        );

        let result = page.evaluate(script).await?;
        let status = result.into_value::<String>().unwrap_or_default();

        match status.as_str() {
            "success" => Ok(()),
            "not_found" => anyhow::bail!("Message not found: {}", id),
            _ => anyhow::bail!("Failed to mark as spam: {}", status),
        }
    }

    pub async fn unspam(&self, _id: &str) -> Result<()> {
        // For unspam, user needs to navigate to Junk folder first
        anyhow::bail!("unspam requires navigating to Junk folder - not yet implemented")
    }

    pub async fn list_labels(&self) -> Result<Vec<String>> {
        // Return known categories - Outlook Web doesn't have an easy API for this
        Ok(vec![
            "Classified".to_string(),
            "Urgent".to_string(),
            "Security".to_string(),
            "Account".to_string(),
            "Important".to_string(),
            "Personal".to_string(),
            "Private".to_string(),
            "Work".to_string(),
            "Family".to_string(),
        ])
    }
}
