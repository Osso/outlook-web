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

        // Close any existing menus to ensure clean state
        menu::close_menus(&page).await?;

        // Additional delay to let the page settle after menu interaction
        tokio::time::sleep(tokio::time::Duration::from_millis(500)).await;

        // Wait for the message to be visible
        if !menu::wait_for_message(&page, id).await? {
            anyhow::bail!("Message not found or not visible: {}", id);
        }

        // Step 1: Check if context menu is open, if not right-click to open it
        if !menu::is_context_menu_open(&page).await? {
            let (x, y) = menu::get_message_position(&page, id)
                .await?
                .ok_or_else(|| anyhow::anyhow!("Message not found: {}", id))?;

            menu::open_context_menu_at(&page, x, y).await?;
        }

        // Step 2: Check if category submenu is open, if not click Categorize
        if !menu::is_category_visible(&page, label).await? {
            if !menu::is_categorize_button_visible(&page).await? {
                anyhow::bail!("Categorize button not found in context menu");
            }

            menu::click_categorize(&page, Some(300)).await?;

            // Retry loop: wait for submenu to appear (check for any category item)
            let mut submenu_opened = false;
            for _ in 0..10 {
                // Check if submenu is open by looking for common category items
                if menu::is_category_visible(&page, "category").await?
                    || menu::is_category_visible(&page, "Manage categories").await?
                {
                    submenu_opened = true;
                    break;
                }
                tokio::time::sleep(tokio::time::Duration::from_millis(200)).await;
            }

            if !submenu_opened {
                anyhow::bail!("Category submenu didn't open");
            }

            // Check if the specific label exists, create if not
            if !menu::is_category_visible(&page, label).await? {
                // Create the missing category
                menu::create_category(&page, label).await?;

                // Close everything and start fresh
                menu::close_menus(&page).await?;
                tokio::time::sleep(tokio::time::Duration::from_millis(500)).await;

                // Reopen context menu with retry
                let (x, y) = menu::get_message_position(&page, id)
                    .await?
                    .ok_or_else(|| anyhow::anyhow!("Message not found: {}", id))?;
                menu::right_click(&page, x, y, Some(500)).await?;

                // Wait for context menu
                let mut menu_opened = false;
                for _ in 0..10 {
                    if menu::is_context_menu_open(&page).await? {
                        menu_opened = true;
                        break;
                    }
                    tokio::time::sleep(tokio::time::Duration::from_millis(200)).await;
                }
                if !menu_opened {
                    anyhow::bail!("Context menu didn't reopen after category creation");
                }

                // Reopen category submenu
                menu::click_categorize(&page, Some(500)).await?;
            }
        }

        // Step 3: Click on the category
        menu::click_category(&page, label, Some(300)).await
    }

    pub async fn remove_label(&self, id: &str, label: &str) -> Result<()> {
        // Remove label is the same as add label - clicking toggles the category
        self.add_label(id, label).await
    }

    pub async fn get_unsubscribe_url(&self, id: &str) -> Result<Option<String>> {
        use crate::browser::click_element;

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        let selector = crate::browser::message_selector(id);
        click_element(&page, &selector, Some(2000)).await?;

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
        use crate::browser::{click_element, press_key};

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        let selector = crate::browser::message_selector(id);
        click_element(&page, &selector, None).await?;
        press_key(&page, "e", None, None).await?;
        Ok(())
    }

    pub async fn trash(&self, id: &str) -> Result<()> {
        use crate::browser::{click_element, press_key};

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        let selector = crate::browser::message_selector(id);
        click_element(&page, &selector, None).await?;
        press_key(&page, "Delete", None, None).await?;
        Ok(())
    }

    pub async fn mark_spam(&self, id: &str) -> Result<()> {
        use crate::menu::{click_menu_item, open_context_menu};

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        let selector = crate::browser::message_selector(id);
        open_context_menu(&page, &selector).await?;

        // Click "Report" to open submenu
        click_menu_item(&page, "report", Some(500)).await?;

        // Click "Junk" in submenu
        click_menu_item(&page, "junk", None).await?;

        Ok(())
    }

    pub async fn unspam(&self, id: &str) -> Result<()> {
        use crate::browser::navigate_to_junk;
        use crate::menu::{click_menu_item, open_context_menu};

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        navigate_to_junk(&page).await?;

        let selector = crate::browser::message_selector(id);
        open_context_menu(&page, &selector).await?;

        // Try "Not junk" first, fall back to "Move to" -> "Inbox"
        if click_menu_item(&page, "not junk", None).await.is_err() {
            click_menu_item(&page, "move", Some(500)).await?;
            click_menu_item(&page, "inbox", None).await?;
        }

        Ok(())
    }

    pub async fn mark_read(&self, id: &str) -> Result<()> {
        use crate::menu::{click_menu_item, open_context_menu};

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        let selector = crate::browser::message_selector(id);
        open_context_menu(&page, &selector).await?;

        click_menu_item(&page, "mark as read", None).await?;
        Ok(())
    }

    pub async fn mark_unread(&self, id: &str) -> Result<()> {
        use crate::menu::{click_menu_item, open_context_menu};

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        let selector = crate::browser::message_selector(id);
        open_context_menu(&page, &selector).await?;

        click_menu_item(&page, "mark as unread", None).await?;
        Ok(())
    }

    pub async fn clear_labels(&self, id: &str) -> Result<()> {
        use crate::menu::{click_menu_item, open_context_menu};

        let browser = connect_or_start_browser(self.port).await?;
        let page = find_outlook_page(&browser).await?;

        let selector = crate::browser::message_selector(id);
        open_context_menu(&page, &selector).await?;

        click_menu_item(&page, "categorize", Some(500)).await?;
        click_menu_item(&page, "clear", None).await?;
        Ok(())
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

            menu::right_click(&page, x, y, Some(1000)).await?;

            if !menu::is_context_menu_open(&page).await? {
                anyhow::bail!("Context menu didn't open");
            }
        }

        // Step 2: Click Categorize to open submenu
        menu::click_categorize(&page, Some(800)).await?;

        // Step 3: Click "Manage Categories" to open the full list
        menu::click_manage_categories(&page, Some(1000)).await?;

        // Step 4: Extract categories from the dialog
        let categories = menu::extract_categories_from_dialog(&page).await?;

        // Close dialog by clicking the X button or backdrop
        let close_script = r#"
            (() => {
                // Try close button first
                const closeBtn = document.querySelector('[role="dialog"] button[aria-label*="Close"], [role="dialog"] button[aria-label*="close"]');
                if (closeBtn) { closeBtn.click(); return true; }
                // Try clicking backdrop/overlay
                const backdrop = document.querySelector('[class*="overlay"], [class*="backdrop"]');
                if (backdrop) { backdrop.click(); return true; }
                // Click outside dialog
                document.body.click();
                return false;
            })()
        "#;
        page.evaluate(close_script).await?;
        tokio::time::sleep(tokio::time::Duration::from_millis(300)).await;

        // If we didn't find categories in the dialog, fall back to the submenu items
        if categories.is_empty() {
            let fallback_categories = menu::extract_categories_from_submenu(&page).await?;

            // Close menu by pressing Escape
            crate::browser::press_key(&page, "Escape", None, None).await?;

            return Ok(fallback_categories);
        }

        Ok(categories)
    }
}
