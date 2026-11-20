#[cfg(windows)]
mod desktop;
#[cfg(windows)]
mod dropzone;
mod geometry;
mod icons;
mod launch;
mod model;
mod store;

use geometry::{GeometryStore, Rect};
use model::Board;
use std::sync::Arc;
use store::Store;
use tauri::menu::{Menu, MenuItem};
use tauri::tray::{MouseButton, MouseButtonState, TrayIconBuilder, TrayIconEvent};
use tauri::{AppHandle, Manager, PhysicalPosition, PhysicalSize, RunEvent, State, WebviewWindow, WindowEvent};
use tauri_plugin_global_shortcut::{GlobalShortcutExt, ShortcutState};

#[tauri::command]
fn load_board(store: State<'_, Store>) -> Board {
    store.get()
}

#[tauri::command]
fn save_board(store: State<'_, Store>, board: Board) -> Result<(), String> {
    store.replace(board)
}

// 느린 네트워크 경로나 셸 확장 때문에 오래 걸릴 수 있어 메인 스레드 밖에서 돈다.
#[tauri::command(async)]
fn launch_target(target: String, elevated: bool) -> Result<(), String> {
    launch::open(&target, elevated)
}

#[tauri::command(async)]
fn classify_target(target: String) -> launch::TargetInfo {
    launch::classify(&target)
}

/// 타일에 그릴 아이콘을 PNG data URL 로 준다. 구하지 못하면 null.
/// `thumbnail` 은 사용자가 그림 파일을 아이콘으로 골랐을 때 그 그림 자체를 쓰게 한다.
#[tauri::command(async)]
fn tile_icon(path: String, thumbnail: bool) -> Option<String> {
    if launch::is_link(path.trim()) {
        return None;
    }
    icons::load(&path, thumbnail).as_ref().and_then(icons::to_data_url)
}

#[tauri::command(async)]
fn list_fonts() -> Vec<String> {
    icons::font_families()
}

#[tauri::command]
fn export_board(store: State<'_, Store>, path: String) -> Result<(), String> {
    let json = serde_json::to_string_pretty(&store.get()).map_err(|e| e.to_string())?;
    std::fs::write(&path, json).map_err(|e| format!("백업 파일을 쓰지 못했습니다: {e}"))
}

/// 백업 파일의 내용으로 보드를 통째로 바꾸고, 바뀐 보드를 돌려준다.
#[tauri::command]
fn import_board(app: AppHandle, store: State<'_, Store>, path: String) -> Result<Board, String> {
    let text = std::fs::read_to_string(&path).map_err(|e| format!("백업 파일을 읽지 못했습니다: {e}"))?;
    store.replace(store::parse_backup(&text)?)?;
    let board = store.get();
    if let Err(e) = apply_hotkey(&app, &board.settings.hotkey) {
        eprintln!("[deskboard] {e}");
    }
    Ok(board)
}

/// 자동 백업이 쌓이는 폴더의 경로
#[tauri::command]
fn backup_dir(store: State<'_, Store>) -> String {
    let dir = store.backup_dir();
    // 아직 백업이 한 번도 없었어도 폴더를 열어 볼 수 있게 만들어 둔다.
    let _ = std::fs::create_dir_all(&dir);
    dir.to_string_lossy().into_owned()
}

/// 불러오기 단축키를 바꾼다. 다른 프로그램이 쓰고 있어 등록하지 못하면 예전 단축키로 되돌리고 알린다.
#[tauri::command]
fn set_hotkey(app: AppHandle, store: State<'_, Store>, hotkey: String) -> Result<(), String> {
    apply_hotkey(&app, &hotkey).inspect_err(|_| {
        let _ = apply_hotkey(&app, &store.get().settings.hotkey);
    })
}

fn apply_hotkey(app: &AppHandle, hotkey: &str) -> Result<(), String> {
    let shortcuts = app.global_shortcut();
    shortcuts.unregister_all().map_err(|e| e.to_string())?;
    let hotkey = hotkey.trim();
    if hotkey.is_empty() {
        return Ok(());
    }
    shortcuts
        .register(hotkey)
        .map_err(|e| format!("단축키 {hotkey} 를 등록하지 못했습니다. 다른 프로그램이 쓰고 있을 수 있습니다. ({e})"))
}

/// 위젯을 다른 창들 위로 불러온다 (이미 불러와 있으면 내려보낸다).
/// 불러왔을 때는 화면에 알려서, 바로 타자를 쳐 바로가기를 찾을 수 있게 한다.
fn toggle_summon(app: &AppHandle) {
    #[cfg(windows)]
    {
        let handle = app.clone();
        let _ = app.run_on_main_thread(move || {
            if desktop::toggle_summon() {
                let _ = tauri::Emitter::emit(&handle, "summoned", ());
            }
        });
    }
    #[cfg(not(windows))]
    let _ = app;
}

fn build_tray(app: &tauri::App) -> tauri::Result<()> {
    let summon = MenuItem::with_id(app, "summon", "위젯 불러오기", true, None::<&str>)?;
    let quit = MenuItem::with_id(app, "quit", "DeskBoard 종료", true, None::<&str>)?;
    let menu = Menu::with_items(app, &[&summon, &quit])?;
    let mut tray = TrayIconBuilder::with_id("deskboard-tray")
        .tooltip("DeskBoard")
        .menu(&menu)
        .show_menu_on_left_click(false)
        .on_menu_event(|app, event| match event.id.as_ref() {
            "summon" => toggle_summon(app),
            "quit" => app.exit(0),
            _ => {}
        })
        .on_tray_icon_event(|tray, event| {
            if let TrayIconEvent::Click { button: MouseButton::Left, button_state: MouseButtonState::Up, .. } = event {
                toggle_summon(tray.app_handle());
            }
        });
    if let Some(icon) = app.default_window_icon() {
        tray = tray.icon(icon.clone());
    }
    tray.build(app)?;
    Ok(())
}

/// 위젯을 주 모니터의 기본 위치·크기로 되돌린다.
#[tauri::command]
fn reset_window(window: WebviewWindow) -> Result<(), String> {
    place(&window, None).map_err(|e| e.to_string())
}

#[tauri::command]
fn quit_app(app: AppHandle) {
    app.exit(0);
}

fn monitor_rect(monitor: &tauri::Monitor) -> Rect {
    Rect {
        x: monitor.position().x,
        y: monitor.position().y,
        width: monitor.size().width,
        height: monitor.size().height,
    }
}

/// 저장된 자리가 아직 화면에 있으면 거기에, 아니면 기본 자리에 창을 놓는다.
fn place(window: &WebviewWindow, saved: Option<Rect>) -> tauri::Result<()> {
    let monitors = window.available_monitors()?;
    let Some(primary) = window.primary_monitor()?.or_else(|| monitors.first().cloned()) else {
        return Ok(());
    };
    let rects: Vec<Rect> = monitors.iter().map(monitor_rect).collect();
    let rect = geometry::resolve(saved, &rects, monitor_rect(&primary), primary.scale_factor());
    window.set_position(PhysicalPosition::new(rect.x, rect.y))?;
    window.set_size(PhysicalSize::new(rect.width, rect.height))?;
    Ok(())
}

/// 창을 옮기거나 크기를 바꿀 때마다 자리를 기억해 둔다.
fn remember_geometry(window: &WebviewWindow, geometry: Arc<GeometryStore>) {
    let tracked = window.clone();
    window.on_window_event(move |event| {
        if !matches!(event, WindowEvent::Moved(_) | WindowEvent::Resized(_)) {
            return;
        }
        let (Ok(position), Ok(size)) = (tracked.outer_position(), tracked.inner_size()) else {
            return;
        };
        // 최소화된 창은 화면 밖 좌표와 0 크기를 알려 오므로 기억하지 않는다.
        if size.width == 0 || size.height == 0 || position.x <= -30000 || position.y <= -30000 {
            return;
        }
        geometry.set(Rect { x: position.x, y: position.y, width: size.width, height: size.height });
    });
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let app = tauri::Builder::default()
        // 두 번 실행하면 같은 파일을 두 창이 번갈아 덮어쓰게 되므로 하나만 띄운다.
        .plugin(tauri_plugin_single_instance::init(|_app, _args, _cwd| {}))
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_autostart::init(tauri_plugin_autostart::MacosLauncher::LaunchAgent, None))
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_process::init())
        .plugin(
            tauri_plugin_global_shortcut::Builder::new()
                .with_handler(|app, _shortcut, event| {
                    if event.state() == ShortcutState::Pressed {
                        toggle_summon(app);
                    }
                })
                .build(),
        )
        .setup(|app| {
            let data_dir = app.path().app_data_dir()?;
            app.manage(Store::open(data_dir.join("board.json")));

            let geometry = Arc::new(GeometryStore::load(data_dir.join("window.json")));
            app.manage(geometry.clone());

            let window = app.get_webview_window("main").ok_or("main window is missing")?;
            if let Err(e) = place(&window, geometry.get()) {
                eprintln!("[deskboard] could not place the window: {e}");
            }
            remember_geometry(&window, geometry.clone());
            geometry::spawn_flusher(geometry);

            // 바탕화면에 붙이면서, 쓰던 프로그램의 포커스를 뺏지 않고 조용히 보인다.
            #[cfg(windows)]
            {
                let widget = windows::Win32::Foundation::HWND(window.hwnd()?.0 as _);
                desktop::install(widget);
                desktop::sync_visibility(|| {
                    let _ = window.show();
                });
                // 다른 프로그램에서 끌어온 파일을 받아 화면에 알린다.
                let app = app.handle().clone();
                dropzone::watch(widget, move |drop| {
                    let _ = tauri::Emitter::emit(&app, "file-drop", drop);
                });
            }
            #[cfg(not(windows))]
            window.show()?;

            // 단축키 등록이나 트레이 아이콘이 안 되어도 위젯 자체는 쓸 수 있어야 한다.
            let hotkey = app.state::<Store>().get().settings.hotkey;
            if let Err(e) = apply_hotkey(app.handle(), &hotkey) {
                eprintln!("[deskboard] {e}");
            }
            if let Err(e) = build_tray(app) {
                eprintln!("[deskboard] tray icon unavailable: {e}");
            }

            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            load_board,
            save_board,
            launch_target,
            classify_target,
            tile_icon,
            list_fonts,
            export_board,
            import_board,
            backup_dir,
            set_hotkey,
            reset_window,
            quit_app,
        ])
        .build(tauri::generate_context!())
        .expect("error while building DeskBoard");

    app.run(|app, event| {
        if let RunEvent::Exit = event {
            app.state::<Arc<GeometryStore>>().flush();
        }
    });
}
