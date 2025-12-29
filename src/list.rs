use crate::api::Message;
use crate::browser::{connect_or_start_browser, find_outlook_page, navigate_to_inbox};
use anyhow::{Context, Result};

/// JavaScript function to extract labels from an element
const EXTRACT_LABELS_JS: &str = r#"
    function extractLabels(el) {
        const labels = [];
        // Look for "Remove X" buttons
        el.querySelectorAll('button[aria-label^="Remove "]').forEach(btn => {
            const label = btn.getAttribute('aria-label').replace('Remove ', '');
            if (label && !labels.includes(label)) labels.push(label);
        });
        // Fallback: look for elements with category title
        if (labels.length === 0) {
            el.querySelectorAll('[title^="Search for all messages with the category "]').forEach(el => {
                const title = el.getAttribute('title');
                const label = title.replace('Search for all messages with the category ', '');
                if (label && !labels.includes(label)) labels.push(label);
            });
        }
        return labels;
    }
"#;

/// Extract messages from the current page view
async fn extract_message_list(page: &chromiumoxide::Page, max: u32) -> Result<Vec<Message>> {
    let script = format!(r#"
        (() => {{
            {extract_labels}
            const messages = [];
            const items = document.querySelectorAll('[data-convid]');
            items.forEach(item => {{
                const id = item.getAttribute('data-convid');
                const ariaLabel = item.getAttribute('aria-label') || '';
                const labels = extractLabels(item);

                // Extract from DOM elements using stable patterns
                let from = '';
                let subject = '';
                let preview = '';

                // Sender: span with email address in title attribute
                const senderEl = item.querySelector('span[title*="@"]');
                if (senderEl) {{
                    from = senderEl.textContent?.trim() || '';
                }}

                // Subject and preview: find text spans that aren't the sender
                const allSpans = item.querySelectorAll('span[title]');
                for (const span of allSpans) {{
                    const title = span.getAttribute('title') || '';
                    const text = span.textContent?.trim() || '';
                    // Skip sender (has @ in title) and empty spans
                    if (title.includes('@') || !text) continue;
                    // Skip time spans (contain : like "15:38")
                    if (/^\d{{1,2}}:\d{{2}}$/.test(text)) continue;
                    // Skip recipient lists (names separated by semicolons like "John; Jane" or "A; B; C")
                    // These are CC/To lists, not subject lines - but use first name as sender if needed
                    if (text.includes(';')) {{
                        const parts = text.split(';').map(p => p.trim());
                        // Detect if any part contains an email address
                        const hasEmails = parts.some(p => p.includes('@'));
                        // Or if all parts look like names (start with capital, reasonably short)
                        const looksLikeNames = parts.every(p => p.length > 0 && p.length < 40 && /^[A-Z]/.test(p));
                        if ((hasEmails || looksLikeNames) && parts.length >= 2) {{
                            // Use first name as sender if we don't have one yet
                            if (!from && parts[0] && !parts[0].includes('@')) {{
                                from = parts[0];
                            }}
                            continue;
                        }}
                    }}
                    // First non-sender span with title is likely subject
                    if (!subject) {{
                        subject = title || text;
                    }}
                }}

                // Preview: look for longer text content that's not the subject
                const textSpans = item.querySelectorAll('span');
                for (const span of textSpans) {{
                    const text = span.textContent?.trim() || '';
                    if (text.length > 50 && text !== subject && !text.includes('@')) {{
                        preview = text;
                        break;
                    }}
                }}

                // Check for Unread marker
                const isUnread = ariaLabel.toLowerCase().includes('unread');

                if (id) {{
                    messages.push({{ id, subject, from, preview, labels, isUnread }});
                }}
            }});
            return JSON.stringify(messages);
        }})()
    "#, extract_labels = EXTRACT_LABELS_JS);

    let result = page.evaluate(script).await?;
    let messages_str = result.into_value::<String>().unwrap_or_default();
    let mut parsed: Vec<Message> = serde_json::from_str(&messages_str).unwrap_or_default();
    parsed.truncate(max as usize);
    Ok(parsed)
}

pub async fn list_messages(port: u16, max: u32) -> Result<Vec<Message>> {
    let browser = connect_or_start_browser(port).await?;
    let page = find_outlook_page(&browser).await?;
    navigate_to_inbox(&page).await?;
    extract_message_list(&page, max).await
}

pub async fn list_spam(port: u16, max: u32) -> Result<Vec<Message>> {
    use crate::browser::navigate_to_junk;

    let browser = connect_or_start_browser(port).await?;
    let page = find_outlook_page(&browser).await?;
    navigate_to_junk(&page).await?;
    extract_message_list(&page, max).await
}

pub async fn get_message(port: u16, id: &str) -> Result<Message> {
    use crate::browser::click_element;

    let browser = connect_or_start_browser(port).await?;
    let page = find_outlook_page(&browser).await?;

    let selector = crate::browser::message_selector(id);
    click_element(&page, &selector, Some(2000)).await?;

    let read_script = format!(r#"
        (() => {{
            {extract_labels}
            const labels = extractLabels(document);

            // Get subject - prefer title attribute for full text
            let subject = '';
            const subjectEl = document.querySelector('.allowTextSelection, [class*="SubjectLine"], [class*="JdFsz"]');
            if (subjectEl) {{
                subject = subjectEl.getAttribute('title') || subjectEl.textContent?.trim() || '';
            }}

            // From
            let from = '';
            const senderEl = document.querySelector('[class*="Sender"], [class*="sender"], [class*="From"], button[class*="Persona"]');
            if (senderEl) {{
                from = senderEl.textContent?.trim();
            }}
            if (!from) {{
                const selected = document.querySelector('[data-convid][aria-selected="true"]');
                if (selected) {{
                    const label = selected.getAttribute('aria-label') || '';
                    const match = label.match(/^([^<]+?)(?:\s+(?:Re:|Fw:|New\s|Your\s|Microsoft|Amazon))/i);
                    if (match) from = match[1].trim();
                }}
            }}

            // Body
            const bodyEl = document.querySelector('div[role="document"]');
            const body = bodyEl?.innerText?.trim();

            // Get ID from selected item
            const selected = document.querySelector('[data-convid][aria-selected="true"]');
            const id = selected?.getAttribute('data-convid') || '';

            return JSON.stringify({{ id, subject, from, body, labels, isUnread: false }});
        }})()
    "#, extract_labels = EXTRACT_LABELS_JS);

    let result = page.evaluate(read_script).await?;
    let message_str = result.into_value::<String>().unwrap_or_default();
    let message: Message =
        serde_json::from_str(&message_str).context("Failed to parse message")?;
    Ok(message)
}
