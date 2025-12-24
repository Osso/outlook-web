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

                // aria-label format: "[Unread] Sender Subject Time Preview..."
                let from = '';
                let subject = '';
                let preview = '';

                // Remove "Unread" prefix if present
                let label = ariaLabel.replace(/^Unread\s+/i, '').trim();

                // Common subject starters to find where sender ends
                const subjectStarters = /\s+(Re:|Fw:|FW:|RE:|New\s|Your\s|Action\s|Welcome|Microsoft|Amazon|Google|Apple|Thanks|Thank\s|Confirm|Verify|Update|Alert|Notice|Reminder|Invoice|Order|Shipping|Delivery)/i;
                const match = label.match(subjectStarters);

                if (match) {{
                    from = label.substring(0, match.index).trim();
                    const rest = label.substring(match.index).trim();
                    // Split on time pattern to separate subject from preview
                    const timeParts = rest.split(/\s+\d{{1,2}}:\d{{2}}\s+|\s+\d{{4}}-\d{{2}}-\d{{2}}\s+/);
                    subject = timeParts[0]?.trim() || '';
                    preview = timeParts.slice(1).join(' ').trim();
                }} else {{
                    // Fallback: look for time pattern to split
                    const timeSplit = label.split(/\s+\d{{1,2}}:\d{{2}}\s+|\s+\d{{4}}-\d{{2}}-\d{{2}}\s+/);
                    if (timeSplit.length > 1) {{
                        // First part has sender + subject, need to split by common patterns
                        const firstPart = timeSplit[0];
                        // Try splitting on < which often separates display name from email
                        const emailMatch = firstPart.match(/^(.+?)<[^>]+>\s*(.*)/);
                        if (emailMatch) {{
                            from = emailMatch[1].trim();
                            subject = emailMatch[2].trim();
                        }} else {{
                            // Assume first few words are sender
                            const words = firstPart.split(/\s+/);
                            from = words.slice(0, 3).join(' ');
                            subject = words.slice(3).join(' ');
                        }}
                        preview = timeSplit.slice(1).join(' ').trim();
                    }} else {{
                        from = label;
                    }}
                }}

                // Clean subject - remove labels
                labels.forEach(label => {{ subject = subject.replace(label, ''); }});
                subject = subject.trim();

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
