#[cfg(windows)]
mod desktop;
mod geometry;
mod launch;
mod model;
mod store;

use geometry::{GeometryStore, Rect};
use model::Board;
use std::sync::Arc;
use store::Store;
use tauri::{AppHandle, Manager, PhysicalPosition, PhysicalSize, RunEvent, State, WebviewWindow, WindowEvent};

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
fn launch_target(target: String) -> Result<(), String> {
    launch::open(&target)
}

#[tauri::command(async)]
fn classify_target(target: String) -> launch::TargetInfo {
    launch::classify(&target)
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
            desktop::install(windows::Win32::Foundation::HWND(window.hwnd()?.0 as _));
            #[cfg(not(windows))]
            window.show()?;

            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            load_board,
            save_board,
            launch_target,
            classify_target,
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
