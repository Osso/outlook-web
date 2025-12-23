use anyhow::Result;
use chromiumoxide::cdp::browser_protocol::input::{
    DispatchMouseEventParams, DispatchMouseEventType, MouseButton,
};
use chromiumoxide::Page;

/// Check if the category submenu is open and the specific category is visible
pub async fn is_category_visible(page: &Page, label: &str) -> Result<bool> {
    let script = format!(
        r#"
        (() => {{
            const items = document.querySelectorAll('[role="menuitemcheckbox"], [role="menuitem"]');
            for (const item of items) {{
                if (item.textContent?.endsWith('{}')) {{
                    return true;
                }}
            }}
            return false;
        }})()
        "#,
        label
    );

    let result = page.evaluate(script).await?;
    Ok(result.into_value::<bool>().unwrap_or(false))
}

/// Check if a context menu is open (any menu with menuitems)
pub async fn is_context_menu_open(page: &Page) -> Result<bool> {
    let script = r#"
        (() => {
            const menu = document.querySelector('[role="menu"]');
            return menu !== null;
        })()
    "#;

    let result = page.evaluate(script).await?;
    Ok(result.into_value::<bool>().unwrap_or(false))
}

/// Check if the Categorize button is visible in the context menu
pub async fn is_categorize_button_visible(page: &Page) -> Result<bool> {
    let script = r#"
        (() => {
            const items = document.querySelectorAll('[role="menuitem"]');
            for (const item of items) {
                if (item.textContent?.toLowerCase().includes('categor')) {
                    return true;
                }
            }
            return false;
        })()
    "#;

    let result = page.evaluate(script).await?;
    Ok(result.into_value::<bool>().unwrap_or(false))
}

/// Right-click at the specified coordinates using CDP
pub async fn right_click(page: &Page, x: f64, y: f64) -> Result<()> {
    // Move mouse to position
    let move_params = DispatchMouseEventParams::builder()
        .r#type(DispatchMouseEventType::MouseMoved)
        .x(x)
        .y(y)
        .build()
        .unwrap();
    page.execute(move_params).await?;

    // Press right mouse button
    let down_params = DispatchMouseEventParams::builder()
        .r#type(DispatchMouseEventType::MousePressed)
        .x(x)
        .y(y)
        .button(MouseButton::Right)
        .click_count(1)
        .build()
        .unwrap();
    page.execute(down_params).await?;

    // Release right mouse button
    let up_params = DispatchMouseEventParams::builder()
        .r#type(DispatchMouseEventType::MouseReleased)
        .x(x)
        .y(y)
        .button(MouseButton::Right)
        .click_count(1)
        .build()
        .unwrap();
    page.execute(up_params).await?;

    Ok(())
}

/// Get the center position of a message element
pub async fn get_message_position(page: &Page, id: &str) -> Result<Option<(f64, f64)>> {
    let script = format!(
        r#"
        (() => {{
            const item = document.querySelector('[data-convid="{}"]');
            if (!item) return null;
            const rect = item.getBoundingClientRect();
            return {{ x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 }};
        }})()
        "#,
        id
    );

    let result = page.evaluate(script).await?;
    let pos: Option<serde_json::Value> = result.into_value().ok();

    Ok(pos.and_then(|p| {
        let x = p.get("x").and_then(|v| v.as_f64())?;
        let y = p.get("y").and_then(|v| v.as_f64())?;
        Some((x, y))
    }))
}

/// Wait for a message element to be visible
pub async fn wait_for_message(page: &Page, id: &str) -> Result<bool> {
    let script = format!(
        r#"
        (async () => {{
            for (let i = 0; i < 50; i++) {{
                const item = document.querySelector('[data-convid="{}"]');
                if (item && item.getBoundingClientRect().height > 0) {{
                    return true;
                }}
                await new Promise(r => setTimeout(r, 100));
            }}
            return false;
        }})()
        "#,
        id
    );

    let result = page.evaluate(script).await?;
    Ok(result.into_value::<bool>().unwrap_or(false))
}

/// Click on the "Categorize" menu item to open the submenu
/// Returns true if clicked, false if not found
pub async fn click_categorize(page: &Page) -> Result<bool> {
    let script = r#"
        (() => {
            const menuItems = document.querySelectorAll('[role="menuitem"]');
            for (const mi of menuItems) {
                const text = mi.textContent?.toLowerCase() || '';
                if (text.includes('categor')) {
                    mi.click();
                    return true;
                }
            }
            return false;
        })()
    "#;

    let result = page.evaluate(script).await?;
    Ok(result.into_value::<bool>().unwrap_or(false))
}

#[derive(serde::Deserialize, Default, Debug)]
pub struct ClickCategoryResult {
    #[serde(default)]
    pub status: String,
    #[serde(default)]
    pub categories: Vec<String>,
}

/// Click on a specific category in the submenu
/// Assumes the category submenu is already open
pub async fn click_category(page: &Page, label: &str) -> Result<ClickCategoryResult> {
    let script = format!(
        r#"
        (() => {{
            const items = document.querySelectorAll('[role="menuitemcheckbox"], [role="menuitem"]');
            const categoryTexts = [];

            for (const item of items) {{
                const text = item.textContent?.trim() || '';
                if (text.endsWith('{label}')) {{
                    item.click();
                    return JSON.stringify({{ status: 'success' }});
                }}
                // Collect category names for error reporting
                if (text.includes('category') ||
                    ['Account', 'Classified', 'Security', 'Urgent', 'Green', 'Orange', 'Red', 'Yellow'].some(c => text.includes(c))) {{
                    categoryTexts.push(text);
                }}
            }}

            return JSON.stringify({{ status: 'category_not_found', categories: categoryTexts }});
        }})()
        "#,
        label = label
    );

    let result = page.evaluate(script).await?;
    let status_json = result.into_value::<String>().unwrap_or_default();

    let mut parsed: ClickCategoryResult = serde_json::from_str(&status_json).unwrap_or_default();
    if parsed.status.is_empty() {
        parsed.status = status_json;
    }

    Ok(parsed)
}
