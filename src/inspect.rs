use crate::browser::{connect_or_start_browser, find_outlook_page};
use anyhow::Result;

pub async fn inspect_dom(port: u16) -> Result<String> {
    let browser = connect_or_start_browser(port).await?;
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
    Ok(info)
}
