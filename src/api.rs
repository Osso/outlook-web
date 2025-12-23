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

        // Ensure we're in the inbox (base URL /mail/0/ is also inbox)
        let nav_script = r#"
            (() => {
                const url = window.location.href;
                // Check if already in inbox: either /inbox or base /mail/N/ with optional message ID
                if (url.includes('/inbox') || url.match(/\/mail\/\d+\/?($|id\/)/)) return 'already_inbox';
                const match = url.match(/(https:\/\/outlook\.[^\/]+\/mail\/\d+\/)/);
                if (match) {
                    window.location.href = match[1] + 'inbox';
                    return 'navigating';
                }
                return 'already_inbox';
            })()
        "#;

        let nav_result = page.evaluate(nav_script).await?;
        let nav_status = nav_result.into_value::<String>().unwrap_or_default();

        if nav_status == "navigating" {
            tokio::time::sleep(tokio::time::Duration::from_secs(2)).await;
        }

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

    pub async fn list_spam(&self, max: u32) -> Result<Vec<Message>> {
        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        // Navigate to Junk folder via URL
        let nav_script = r#"
            (() => {
                const url = window.location.href;
                // Extract base URL (e.g., https://outlook.live.com/mail/0/)
                const match = url.match(/(https:\/\/outlook\.[^\/]+\/mail\/\d+\/)/);
                if (match) {
                    window.location.href = match[1] + 'junkemail';
                    return 'navigating';
                }
                return 'url_parse_failed';
            })()
        "#;

        let nav_result = page.evaluate(nav_script).await?;
        let nav_status = nav_result.into_value::<String>().unwrap_or_default();

        if nav_status == "url_parse_failed" {
            anyhow::bail!("Failed to parse Outlook URL for navigation");
        }

        // Wait for navigation and page load
        tokio::time::sleep(tokio::time::Duration::from_secs(2)).await;

        // Now list messages (same logic as list_messages)
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
        let message: Message =
            serde_json::from_str(&message_str).context("Failed to parse message")?;
        Ok(message)
    }

    pub async fn add_label(&self, id: &str, label: &str) -> Result<()> {
        use crate::menu;

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        // Wait for the message to be visible
        if !menu::wait_for_message(&page, id).await? {
            anyhow::bail!("Message not found or not visible: {}", id);
        }

        // Step 1: Check if context menu is open, if not right-click to open it
        if !menu::is_context_menu_open(&page).await? {
            let (x, y) = menu::get_message_position(&page, id)
                .await?
                .ok_or_else(|| anyhow::anyhow!("Message not found: {}", id))?;

            menu::right_click(&page, x, y).await?;
            tokio::time::sleep(tokio::time::Duration::from_millis(1000)).await;

            if !menu::is_context_menu_open(&page).await? {
                anyhow::bail!("Context menu didn't open");
            }
        }

        // Step 2: Check if category submenu is open, if not click Categorize
        if !menu::is_category_visible(&page, label).await? {
            if !menu::is_categorize_button_visible(&page).await? {
                anyhow::bail!("Categorize button not found");
            }

            if !menu::click_categorize(&page).await? {
                anyhow::bail!("Failed to click Categorize");
            }
            tokio::time::sleep(tokio::time::Duration::from_millis(800)).await;

            if !menu::is_category_visible(&page, label).await? {
                anyhow::bail!("Category submenu didn't open");
            }
        }

        // Step 3: Click on the category
        let result = menu::click_category(&page, label).await?;

        match result.status.as_str() {
            "success" => Ok(()),
            "category_not_found" => anyhow::bail!(
                "Category '{}' not found. Available: {:?}",
                label,
                result.categories
            ),
            _ => anyhow::bail!("Failed to add label: {}", result.status),
        }
    }

    pub async fn remove_label(&self, id: &str, label: &str) -> Result<()> {
        // Remove label is the same as add label - clicking toggles the category
        self.add_label(id, label).await
    }

    pub async fn get_unsubscribe_url(&self, id: &str) -> Result<Option<String>> {
        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        // Click on message to open it
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

        // Search for unsubscribe links in the message body
        let script = r#"
            (() => {
                const bodyEl = document.querySelector('div[role="document"]');
                if (!bodyEl) return null;

                // Look for links with "unsubscribe" in text or href
                const links = bodyEl.querySelectorAll('a[href]');
                for (const link of links) {
                    const href = link.href || '';
                    const text = link.textContent?.toLowerCase() || '';
                    if (text.includes('unsubscribe') ||
                        text.includes('opt out') ||
                        text.includes('opt-out') ||
                        href.toLowerCase().includes('unsubscribe') ||
                        href.toLowerCase().includes('optout')) {
                        return href;
                    }
                }

                // Also check for List-Unsubscribe in any visible headers
                const allText = bodyEl.innerText || '';
                const match = allText.match(/unsubscribe[:\s]*(https?:\/\/[^\s<>"]+)/i);
                if (match) {
                    return match[1];
                }

                return null;
            })()
        "#;

        let result = page.evaluate(script).await?;
        Ok(result.into_value::<Option<String>>().unwrap_or(None))
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

    pub async fn unspam(&self, id: &str) -> Result<()> {
        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        // Navigate to Junk folder via URL
        let nav_script = r#"
            (() => {
                const url = window.location.href;
                const match = url.match(/(https:\/\/outlook\.[^\/]+\/mail\/\d+\/)/);
                if (match) {
                    window.location.href = match[1] + 'junkemail';
                    return 'navigating';
                }
                return 'url_parse_failed';
            })()
        "#;

        let nav_result = page.evaluate(nav_script).await?;
        let nav_status = nav_result.into_value::<String>().unwrap_or_default();

        if nav_status == "url_parse_failed" {
            anyhow::bail!("Failed to parse Outlook URL for navigation");
        }

        tokio::time::sleep(tokio::time::Duration::from_secs(2)).await;

        // Right-click on message and select "Not junk" or "Move to Inbox"
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

                // Look for "Not junk", "Not spam", or "Move to" options
                const menuItems = document.querySelectorAll('[role="menuitem"]');
                for (const mi of menuItems) {{
                    const text = mi.textContent?.toLowerCase() || '';
                    if (text.includes('not junk') || text.includes('not spam')) {{
                        mi.click();
                        return 'success';
                    }}
                }}

                // Try "Move to" -> "Inbox" approach
                for (const mi of menuItems) {{
                    const text = mi.textContent?.toLowerCase() || '';
                    if (text.includes('move to') || text.includes('move')) {{
                        mi.click();
                        await new Promise(r => setTimeout(r, 500));

                        const subItems = document.querySelectorAll('[role="menuitem"]');
                        for (const si of subItems) {{
                            if (si.textContent?.toLowerCase().includes('inbox')) {{
                                si.click();
                                return 'success';
                            }}
                        }}
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
            "not_found" => anyhow::bail!("Message not found in Junk folder: {}", id),
            _ => anyhow::bail!("Failed to move to inbox: {}", status),
        }
    }

    pub async fn mark_read(&self, id: &str) -> Result<()> {
        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

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

                const menuItems = document.querySelectorAll('[role="menuitem"]');
                for (const mi of menuItems) {{
                    const text = mi.textContent?.toLowerCase() || '';
                    if (text.includes('mark as read') || text.includes('mark read')) {{
                        mi.click();
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
            _ => anyhow::bail!("Failed to mark as read: {}", status),
        }
    }

    pub async fn mark_unread(&self, id: &str) -> Result<()> {
        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

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

                const menuItems = document.querySelectorAll('[role="menuitem"]');
                for (const mi of menuItems) {{
                    const text = mi.textContent?.toLowerCase() || '';
                    if (text.includes('mark as unread') || text.includes('mark unread')) {{
                        mi.click();
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
            _ => anyhow::bail!("Failed to mark as unread: {}", status),
        }
    }

    pub async fn clear_labels(&self, id: &str) -> Result<()> {
        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

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

                // Find and click "Categorize" menu item
                const menuItems = document.querySelectorAll('[role="menuitem"], [role="menuitemcheckbox"]');
                for (const mi of menuItems) {{
                    if (mi.textContent?.includes('Categorize')) {{
                        mi.click();
                        await new Promise(r => setTimeout(r, 500));
                        break;
                    }}
                }}

                // Find and click "Clear categories" or similar
                const categoryItems = document.querySelectorAll('[role="menuitem"], [role="menuitemcheckbox"]');
                for (const ci of categoryItems) {{
                    const text = ci.textContent?.toLowerCase() || '';
                    if (text.includes('clear') && text.includes('categor')) {{
                        ci.click();
                        return 'success';
                    }}
                }}

                return 'clear_not_found';
            }})()
        "#,
            id
        );

        let result = page.evaluate(script).await?;
        let status = result.into_value::<String>().unwrap_or_default();

        match status.as_str() {
            "success" => Ok(()),
            "not_found" => anyhow::bail!("Message not found: {}", id),
            "clear_not_found" => anyhow::bail!("Clear categories option not found"),
            _ => anyhow::bail!("Failed to clear labels: {}", status),
        }
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
