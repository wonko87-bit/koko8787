// 콘솔 창 없이 실행 (릴리스 빌드)
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod google;

use std::{
    fs,
    path::{Path, PathBuf},
    time::Duration,
};

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Workbook};
use serde::Deserialize;
use tauri::{
    menu::{Menu, MenuItem, PredefinedMenuItem},
    tray::{MouseButton, MouseButtonState, TrayIconBuilder, TrayIconEvent},
    AppHandle, Emitter, Manager, WindowEvent,
};
use tauri_plugin_autostart::{MacosLauncher, ManagerExt};
use tauri_plugin_global_shortcut::{Code, GlobalShortcutExt, Modifiers, Shortcut, ShortcutState};
use tauri_plugin_opener::OpenerExt;

const DATA_FILE: &str = "planner-data.json";
const BACKUP_KEEP: usize = 14;
const HIDDEN_ARG: &str = "--hidden";

/* ───────────────────────── 경로 ───────────────────────── */

/// 문서\Planner (없으면 생성)
fn data_root(app: &AppHandle) -> Result<PathBuf, String> {
    let base = app
        .path()
        .document_dir()
        .or_else(|_| app.path().app_data_dir())
        .map_err(|e| e.to_string())?;
    let dir = base.join("Planner");
    fs::create_dir_all(&dir).map_err(|e| format!("폴더를 만들 수 없음: {e}"))?;
    Ok(dir)
}

fn sub_dir(app: &AppHandle, name: &str) -> Result<PathBuf, String> {
    let dir = data_root(app)?.join(name);
    fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir)
}

/// 파일 이름에 쓸 수 없는 문자 제거 + 중복 시 (2), (3)…
fn unique_path(dir: &Path, stem: &str, ext: &str) -> PathBuf {
    let clean: String = stem
        .chars()
        .map(|c| if "<>:\"/\\|?*".contains(c) || c.is_control() { '_' } else { c })
        .collect();
    let mut p = dir.join(format!("{clean}.{ext}"));
    let mut n = 2;
    while p.exists() {
        p = dir.join(format!("{clean} ({n}).{ext}"));
        n += 1;
    }
    p
}

/* ───────────────────────── 데이터 저장 ───────────────────────── */

#[tauri::command]
fn load_data(app: AppHandle) -> Result<Option<String>, String> {
    let f = data_root(&app)?.join(DATA_FILE);
    if !f.exists() {
        return Ok(None);
    }
    fs::read_to_string(&f).map(Some).map_err(|e| e.to_string())
}

#[tauri::command]
fn save_data(app: AppHandle, json: String) -> Result<(), String> {
    // 깨진 데이터가 저장되지 않도록 검사
    serde_json::from_str::<serde_json::Value>(&json).map_err(|e| format!("잘못된 데이터: {e}"))?;
    let root = data_root(&app)?;
    let file = root.join(DATA_FILE);

    // 하루 한 번, 덮어쓰기 전 상태를 backups\YYYY-MM-DD.json 으로 보관
    if file.exists() {
        if let Ok(bdir) = sub_dir(&app, "backups") {
            let b = bdir.join(format!("{}.json", chrono::Local::now().format("%Y-%m-%d")));
            if !b.exists() {
                let _ = fs::copy(&file, &b);
                prune_backups(&bdir);
            }
        }
    }

    // 임시 파일에 쓴 뒤 교체 (저장 중 꺼져도 원본 보호)
    let tmp = root.join(format!("{DATA_FILE}.tmp"));
    fs::write(&tmp, json.as_bytes()).map_err(|e| e.to_string())?;
    fs::rename(&tmp, &file).map_err(|e| e.to_string())
}

fn prune_backups(dir: &Path) {
    let mut files: Vec<PathBuf> = match fs::read_dir(dir) {
        Ok(rd) => rd
            .filter_map(|e| e.ok().map(|e| e.path()))
            .filter(|p| p.extension().map(|x| x == "json").unwrap_or(false))
            .collect(),
        Err(_) => return,
    };
    files.sort();
    while files.len() > BACKUP_KEEP {
        let _ = fs::remove_file(files.remove(0));
    }
}

#[tauri::command]
fn data_dir(app: AppHandle) -> Result<String, String> {
    Ok(data_root(&app)?.to_string_lossy().into_owned())
}

#[tauri::command]
fn open_data_dir(app: AppHandle) -> Result<(), String> {
    let dir = data_root(&app)?;
    app.opener()
        .open_path(dir.to_string_lossy(), None::<&str>)
        .map_err(|e| e.to_string())
}

#[tauri::command]
fn export_json(app: AppHandle, json: String, name: String) -> Result<String, String> {
    let dir = sub_dir(&app, "exports")?;
    let stem = name.trim_end_matches(".json");
    let path = unique_path(&dir, stem, "json");
    fs::write(&path, json.as_bytes()).map_err(|e| e.to_string())?;
    let _ = app.opener().reveal_item_in_dir(&path);
    Ok(path.to_string_lossy().into_owned())
}

/* ───────────────────────── 엑셀 내보내기 ───────────────────────── */

#[derive(Deserialize)]
struct SheetIn {
    name: String,
    columns: Vec<String>,
    #[serde(default)]
    widths: Vec<f64>,
    rows: Vec<Vec<String>>,
}

#[derive(Deserialize)]
struct ExportIn {
    sheets: Vec<SheetIn>,
}

#[tauri::command]
fn export_xlsx(app: AppHandle, payload: String, name: String) -> Result<String, String> {
    let input: ExportIn = serde_json::from_str(&payload).map_err(|e| e.to_string())?;
    let dir = sub_dir(&app, "exports")?;
    let path = unique_path(&dir, &name, "xlsx");
    write_xlsx(&input, &path).map_err(|e| e.to_string())?;
    let _ = app.opener().reveal_item_in_dir(&path);
    Ok(path.to_string_lossy().into_owned())
}

fn write_xlsx(input: &ExportIn, path: &Path) -> Result<(), rust_xlsxwriter::XlsxError> {
    const FONT: &str = "맑은 고딕";
    let line = Color::RGB(0xE3E8EF);
    let head = Format::new()
        .set_font_name(FONT)
        .set_bold()
        .set_font_color(Color::White)
        .set_background_color(Color::RGB(0x1F3A5F))
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter);
    let base = Format::new()
        .set_font_name(FONT)
        .set_font_size(10)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap()
        .set_border_bottom(FormatBorder::Thin)
        .set_border_bottom_color(line);
    let center = base.clone().set_align(FormatAlign::Center);
    let done = center.clone().set_font_color(Color::RGB(0x1F9D5B)).set_bold();
    let carry = center.clone().set_font_color(Color::RGB(0x2E6BD1)).set_bold();
    let cancel = center.clone().set_font_color(Color::RGB(0xA9B1BD));
    let red = base.clone().set_font_color(Color::RGB(0xD9434B));

    let mut wb = Workbook::new();
    for s in &input.sheets {
        let ws = wb.add_worksheet();
        ws.set_name(&s.name)?;
        ws.set_row_height(0, 24)?;
        for (c, col) in s.columns.iter().enumerate() {
            ws.write_with_format(0, c as u16, col, &head)?;
            ws.set_column_width(c as u16, *s.widths.get(c).unwrap_or(&12.0))?;
        }
        for (r, row) in s.rows.iter().enumerate() {
            let r = (r + 1) as u32;
            for (c, v) in row.iter().enumerate() {
                let col = s.columns.get(c).map(String::as_str).unwrap_or("");
                let fmt = match col {
                    "상태" if v.starts_with('✔') => &done,
                    "상태" if v.starts_with('→') => &carry,
                    "상태" if v.starts_with('✕') => &cancel,
                    "상태" | "요일" | "시간" | "구분" | "날짜" | "마감일" => &center,
                    "공휴일" => &red,
                    _ => &base,
                };
                ws.write_with_format(r, c as u16, v, fmt)?;
            }
        }
        ws.set_freeze_panes(1, 0)?;
        if !s.rows.is_empty() && !s.columns.is_empty() {
            ws.autofilter(0, 0, s.rows.len() as u32, (s.columns.len() - 1) as u16)?;
        }
    }
    wb.save(path)
}

/* ───────────────────────── 자동 실행 / 종료 ───────────────────────── */

#[tauri::command]
fn get_autostart(app: AppHandle) -> Result<bool, String> {
    app.autolaunch().is_enabled().map_err(|e| e.to_string())
}

#[tauri::command]
fn set_autostart(app: AppHandle, enabled: bool) -> Result<(), String> {
    let al = app.autolaunch();
    if enabled { al.enable() } else { al.disable() }.map_err(|e| e.to_string())
}

#[tauri::command]
fn quit_app(app: AppHandle) {
    app.exit(0);
}

/* ───────────────────────── 동기화 (Google 드라이브) ───────────────────────── */

#[derive(serde::Serialize)]
struct SyncStatus {
    available: bool,
    connected: bool,
}

#[tauri::command]
fn sync_status(app: AppHandle) -> SyncStatus {
    SyncStatus { available: google::client().is_some(), connected: google::is_connected(&app) }
}

/// 블로킹 작업을 별도 스레드에서 실행
async fn blocking<T: Send + 'static>(
    app: AppHandle,
    f: impl FnOnce(&AppHandle, &google::Shared) -> Result<T, String> + Send + 'static,
) -> Result<T, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let shared = app.state::<google::Shared>();
        f(&app, shared.inner())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn google_login(app: AppHandle) -> Result<(), String> {
    let res = blocking(app.clone(), |a, s| google::login(a, s)).await;
    show_main(&app); // 브라우저에서 돌아오면 플래너를 앞으로
    res
}

#[tauri::command]
async fn google_logout(app: AppHandle) -> Result<(), String> {
    blocking(app, |a, s| {
        google::logout(a, s);
        Ok(())
    })
    .await
}

#[tauri::command]
async fn drive_pull(app: AppHandle) -> Result<Vec<String>, String> {
    blocking(app, |a, s| google::pull(a, s)).await
}

#[tauri::command]
async fn drive_push(app: AppHandle, content: String) -> Result<(), String> {
    blocking(app, move |a, s| google::push(a, s, &content)).await
}

/// 동기화용 데이터 변환 직전 원본 보관 (한 번만)
#[tauri::command]
fn backup_once(app: AppHandle, json: String, name: String) -> Result<(), String> {
    let dir = sub_dir(&app, "backups")?;
    let stem: String = name.chars().filter(|c| c.is_ascii_alphanumeric() || *c == '-').collect();
    let path = dir.join(format!("{stem}.json.bak"));
    if !path.exists() {
        fs::write(&path, json.as_bytes()).map_err(|e| e.to_string())?;
    }
    Ok(())
}

/* ───────────────────────── 창 / 트레이 ───────────────────────── */

fn show_main(app: &AppHandle) {
    if let Some(w) = app.get_webview_window("main") {
        let _ = w.unminimize();
        let _ = w.show();
        let _ = w.set_focus();
        let _ = w.emit("window-shown", ());
    }
}

fn toggle_main(app: &AppHandle) {
    if let Some(w) = app.get_webview_window("main") {
        let visible = w.is_visible().unwrap_or(false);
        let minimized = w.is_minimized().unwrap_or(false);
        let focused = w.is_focused().unwrap_or(false);
        if visible && !minimized && focused {
            let _ = w.hide();
        } else {
            show_main(app);
        }
    }
}

/// 화면 쪽에 저장을 맡기고 종료. 응답이 없으면 2초 뒤 강제 종료.
fn request_quit(app: &AppHandle) {
    let _ = app.emit("request-quit", ());
    let handle = app.clone();
    std::thread::spawn(move || {
        std::thread::sleep(Duration::from_secs(2));
        handle.exit(0);
    });
}

fn build_tray(app: &AppHandle) -> tauri::Result<()> {
    let open = MenuItem::with_id(app, "open", "플래너 열기  (Ctrl+Alt+P)", true, None::<&str>)?;
    let sep = PredefinedMenuItem::separator(app)?;
    let quit = MenuItem::with_id(app, "quit", "종료", true, None::<&str>)?;
    let menu = Menu::with_items(app, &[&open, &sep, &quit])?;

    let mut tray = TrayIconBuilder::with_id("main")
        .tooltip("Planner  ·  Ctrl+Alt+P")
        .menu(&menu)
        .show_menu_on_left_click(false)
        .on_menu_event(|app, event| match event.id.as_ref() {
            "open" => show_main(app),
            "quit" => request_quit(app),
            _ => {}
        })
        .on_tray_icon_event(|tray, event| {
            if let TrayIconEvent::Click {
                button: MouseButton::Left,
                button_state: MouseButtonState::Up,
                ..
            } = event
            {
                toggle_main(tray.app_handle());
            }
        });
    if let Some(icon) = app.default_window_icon() {
        tray = tray.icon(icon.clone());
    }
    tray.build(app)?;
    Ok(())
}

fn main() {
    let hotkey = Shortcut::new(Some(Modifiers::CONTROL | Modifiers::ALT), Code::KeyP);

    tauri::Builder::default()
        // 두 번째 실행 시 기존 창을 앞으로
        .plugin(tauri_plugin_single_instance::init(|app, _args, _cwd| show_main(app)))
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_autostart::init(
            MacosLauncher::LaunchAgent,
            Some(vec![HIDDEN_ARG]),
        ))
        .plugin(
            tauri_plugin_global_shortcut::Builder::new()
                .with_handler(move |app, sc, event| {
                    if event.state() == ShortcutState::Pressed && sc == &hotkey {
                        toggle_main(app);
                    }
                })
                .build(),
        )
        .manage(google::Shared::default())
        .setup(move |app| {
            let handle = app.handle();
            build_tray(handle)?;
            if let Err(e) = handle.global_shortcut().register(hotkey) {
                eprintln!("단축키 등록 실패 (다른 프로그램이 사용 중일 수 있음): {e}");
            }
            // 윈도우 시작 시 자동 실행(--hidden)이면 트레이에만 둔다
            let start_hidden = std::env::args().any(|a| a == HIDDEN_ARG);
            if !start_hidden {
                show_main(handle);
            }
            Ok(())
        })
        .on_window_event(|window, event| {
            // X 버튼 = 트레이로 숨기기 (종료는 트레이 메뉴에서)
            if let WindowEvent::CloseRequested { api, .. } = event {
                api.prevent_close();
                let _ = window.hide();
            }
        })
        .invoke_handler(tauri::generate_handler![
            load_data,
            save_data,
            data_dir,
            open_data_dir,
            export_json,
            export_xlsx,
            get_autostart,
            set_autostart,
            quit_app,
            sync_status,
            google_login,
            google_logout,
            drive_pull,
            drive_push,
            backup_once
        ])
        .run(tauri::generate_context!())
        .expect("Planner 실행 중 오류");
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xlsx_and_helpers() {
        let dir = std::env::temp_dir().join("planner-test");
        let _ = fs::remove_dir_all(&dir);
        fs::create_dir_all(&dir).unwrap();
        let input: ExportIn = serde_json::from_str(r#"{"sheets":[
            {"name":"할일·일정","columns":["날짜","요일","공휴일","상태","시간","할일·일정","구분","비고"],"widths":[12,6,12,9,8,48,8,16],
             "rows":[["2026-12-25","금","성탄절","✔ 완료","09:00","테스트","업무",""],["2026-12-26","토","","→ 이월","","보고서","개인","12/25에서 이월"]]},
            {"name":"월간","columns":["월","종류","상태","내용","마감일","구분"],"rows":[]}]}"#).unwrap();
        let p = unique_path(&dir, "플래너:이번 주?", "xlsx");
        assert!(p.file_name().unwrap().to_string_lossy().starts_with("플래너_이번 주_"));
        write_xlsx(&input, &p).unwrap();
        assert!(fs::metadata(&p).unwrap().len() > 1000);
        let p2 = unique_path(&dir, "플래너:이번 주?", "xlsx");
        assert!(p2.to_string_lossy().ends_with("(2).xlsx"));

        let b = dir.join("backups");
        fs::create_dir_all(&b).unwrap();
        for d in 1..=20 {
            fs::write(b.join(format!("2026-11-{d:02}.json")), "{}").unwrap();
        }
        prune_backups(&b);
        let mut left: Vec<_> = fs::read_dir(&b).unwrap().map(|e| e.unwrap().file_name().into_string().unwrap()).collect();
        left.sort();
        assert_eq!(left.len(), BACKUP_KEEP);
        assert_eq!(left[0], "2026-11-07.json");
        println!("xlsx: {}", p.display());
    }
}
