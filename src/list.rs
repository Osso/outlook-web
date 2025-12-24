use crate::api::Message;
use crate::browser::{connect_or_start_browser, find_outlook_page, navigate_to_inbox};
use anyhow::{Context, Result};

pub async fn list_messages(port: u16, max: u32) -> Result<Vec<Message>> {
    let browser = connect_or_start_browser(port).await?;
    let page = find_outlook_page(&browser).await?;

    navigate_to_inbox(&page).await?;

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

pub async fn list_spam(port: u16, max: u32) -> Result<Vec<Message>> {
    use crate::browser::navigate_to_junk;

    let browser = connect_or_start_browser(port).await?;
    let page = find_outlook_page(&browser).await?;

    navigate_to_junk(&page).await?;

    let script = r#"
        (() => {
            const messages = [];
            const items = document.querySelectorAll('[data-convid]');
            items.forEach(item => {
                const id = item.getAttribute('data-convid');
                const ariaLabel = item.getAttribute('aria-label') || '';

                let from = '';
                let subject = '';
                let preview = '';

                let label = ariaLabel.replace(/^Unread\s+/i, '').trim();

                const subjectStarters = /\s+(Re:|Fw:|FW:|RE:|New\s|Your\s|Action\s|Welcome|Microsoft|Amazon|Google|Apple|Thanks|Thank\s|Confirm|Verify|Update|Alert|Notice|Reminder|Invoice|Order|Shipping|Delivery)/i;
                const match = label.match(subjectStarters);

                if (match) {
                    from = label.substring(0, match.index).trim();
                    const rest = label.substring(match.index).trim();
                    const timeParts = rest.split(/\s+\d{1,2}:\d{2}\s+|\s+\d{4}-\d{2}-\d{2}\s+/);
                    subject = timeParts[0]?.trim() || '';
                    preview = timeParts.slice(1).join(' ').trim();
                } else {
                    const timeSplit = label.split(/\s+\d{1,2}:\d{2}\s+|\s+\d{4}-\d{2}-\d{2}\s+/);
                    if (timeSplit.length > 1) {
                        const firstPart = timeSplit[0];
                        const emailMatch = firstPart.match(/^(.+?)<[^>]+>\s*(.*)/);
                        if (emailMatch) {
                            from = emailMatch[1].trim();
                            subject = emailMatch[2].trim();
                        } else {
                            const words = firstPart.split(/\s+/);
                            from = words.slice(0, 3).join(' ');
                            subject = words.slice(3).join(' ');
                        }
                        preview = timeSplit.slice(1).join(' ').trim();
                    } else {
                        from = label;
                    }
                }

                const isUnread = ariaLabel.toLowerCase().includes('unread');

                if (id) {
                    messages.push({ id, subject, from, preview, labels: [], isUnread });
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

pub async fn get_message(port: u16, id: &str) -> Result<Message> {
    let browser = connect_or_start_browser(port).await?;
    let page = find_outlook_page(&browser).await?;

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
    let message: Message =
        serde_json::from_str(&message_str).context("Failed to parse message")?;
    Ok(message)
}
