use crate::model::{FileEntry, Suggestion};
use serde::Serialize;
use tauri::{AppHandle, Emitter, Manager, PhysicalPosition};

#[derive(Serialize, Clone)]
struct ToastPayload<'a> {
    entry: &'a FileEntry,
    suggestions: &'a [Suggestion],
}

/// 토스트 창에 새 파일 이벤트를 보내고, 우하단에 표시한다.
pub fn show_toast(app: &AppHandle, entry: &FileEntry, suggestions: &[Suggestion]) {
    let Some(win) = app.get_webview_window("toast") else { return };

    let _ = app.emit_to("toast", "toast-file", ToastPayload { entry, suggestions });

    if let Ok(Some(monitor)) = win.primary_monitor() {
        let screen = monitor.size();
        let size = win.outer_size().unwrap_or(tauri::PhysicalSize::new(400, 260));
        let x = screen.width.saturating_sub(size.width + 16) as i32;
        let y = screen.height.saturating_sub(size.height + 72) as i32;
        let _ = win.set_position(PhysicalPosition::new(x, y));
    }
    let _ = win.show();
}

pub fn hide_toast(app: &AppHandle) {
    if let Some(win) = app.get_webview_window("toast") {
        let _ = win.hide();
    }
}
