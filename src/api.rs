use crate::browser::{connect_or_start_browser, find_outlook_page};
use anyhow::Result;
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
        crate::list::list_messages(self.port, max).await
    }

    pub async fn list_spam(&self, max: u32) -> Result<Vec<Message>> {
        crate::list::list_spam(self.port, max).await
    }

    pub async fn get_message(&self, id: &str) -> Result<Message> {
        crate::list::get_message(self.port, id).await
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
        use crate::browser::click_element;

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        let selector = format!("[data-convid=\"{}\"]", id);
        if !click_element(&page, &selector, Some(2000)).await? {
            anyhow::bail!("Message not found: {}", id);
        }

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
        use crate::browser::click_element;

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        let selector = format!("[data-convid=\"{}\"]", id);
        if !click_element(&page, &selector, None).await? {
            anyhow::bail!("Message not found: {}", id);
        }

        // Press 'e' for archive
        page.evaluate("document.dispatchEvent(new KeyboardEvent('keydown', { key: 'e', code: 'KeyE', bubbles: true }))").await?;

        tokio::time::sleep(tokio::time::Duration::from_millis(500)).await;
        Ok(())
    }

    pub async fn trash(&self, id: &str) -> Result<()> {
        use crate::browser::click_element;

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        let selector = format!("[data-convid=\"{}\"]", id);
        if !click_element(&page, &selector, None).await? {
            anyhow::bail!("Message not found: {}", id);
        }

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
        use crate::browser::navigate_to_inbox;
        use crate::menu;

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        navigate_to_inbox(&page).await?;

        // Find any message to right-click
        let first_msg_script = r#"
            (() => {
                const item = document.querySelector('[data-convid]');
                return item?.getAttribute('data-convid') || null;
            })()
        "#;

        let result = page.evaluate(first_msg_script).await?;
        let msg_id: Option<String> = result.into_value().ok();

        let msg_id =
            msg_id.ok_or_else(|| anyhow::anyhow!("No messages found to open category menu"))?;

        // Wait for message to be visible
        if !menu::wait_for_message(&page, &msg_id).await? {
            anyhow::bail!("Message not visible");
        }

        // Step 1: Right-click to open context menu
        if !menu::is_context_menu_open(&page).await? {
            let (x, y) = menu::get_message_position(&page, &msg_id)
                .await?
                .ok_or_else(|| anyhow::anyhow!("Message not found"))?;

            menu::right_click(&page, x, y).await?;
            tokio::time::sleep(tokio::time::Duration::from_millis(1000)).await;

            if !menu::is_context_menu_open(&page).await? {
                anyhow::bail!("Context menu didn't open");
            }
        }

        // Step 2: Click Categorize to open submenu
        if !menu::click_categorize(&page).await? {
            anyhow::bail!("Failed to click Categorize");
        }
        tokio::time::sleep(tokio::time::Duration::from_millis(800)).await;

        // Step 3: Click "Manage Categories" to open the full list
        if !menu::click_manage_categories(&page).await? {
            let items = menu::list_menu_items(&page).await?;
            anyhow::bail!("'Manage Categories' not found. Available items: {:?}", items);
        }

        tokio::time::sleep(tokio::time::Duration::from_secs(1)).await;

        // Step 4: Extract categories from the dialog
        let categories = menu::extract_categories_from_dialog(&page).await?;

        // Close dialog by pressing Escape
        page.evaluate("document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', bubbles: true }))").await?;
        tokio::time::sleep(tokio::time::Duration::from_millis(300)).await;

        // If we didn't find categories in the dialog, fall back to the submenu items
        if categories.is_empty() {
            let fallback_categories = menu::extract_categories_from_submenu(&page).await?;

            // Close menu by clicking elsewhere
            page.evaluate("document.body.click()").await?;

            return Ok(fallback_categories);
        }

        Ok(categories)
    }
}
