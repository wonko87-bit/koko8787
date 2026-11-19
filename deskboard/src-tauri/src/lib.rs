#[cfg(windows)]
mod desktop;
mod geometry;
mod launch;
mod model;
mod store;

use geometry::{GeometryStore, Rect};
use model::Board;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use std::time::Duration;
use store::Store;
use tauri::{
    AppHandle, Manager, PhysicalPosition, PhysicalSize, RunEvent, State, WebviewWindow,
    WebviewWindowBuilder, WindowEvent,
};

/// 사용자가 위젯 창을 직접 닫았는지. 닫지 않았는데 창이 없어졌다면 다시 만들어야 한다.
#[derive(Default)]
struct ClosedByUser(AtomicBool);

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

/// 창을 옮기거나 크기를 바꿀 때마다 자리를 기억해 두고, 사용자가 닫으면 그것도 적어 둔다.
fn watch_window(window: &WebviewWindow, geometry: Arc<GeometryStore>) {
    let tracked = window.clone();
    window.on_window_event(move |event| match event {
        WindowEvent::Moved(_) | WindowEvent::Resized(_) => {
            let (Ok(position), Ok(size)) = (tracked.outer_position(), tracked.inner_size()) else {
                return;
            };
            // 최소화된 창은 화면 밖 좌표와 0 크기를 알려 오므로 기억하지 않는다.
            if size.width == 0 || size.height == 0 || position.x <= -30000 || position.y <= -30000 {
                return;
            }
            geometry.set(Rect { x: position.x, y: position.y, width: size.width, height: size.height });
        }
        WindowEvent::CloseRequested { .. } => {
            tracked.state::<ClosedByUser>().0.store(true, Ordering::Release);
        }
        _ => {}
    });
}

/// 위젯 창을 만들어 기억해 둔 자리에 놓고, 바탕화면에 붙인 뒤 보인다.
fn open_widget(app: &AppHandle) -> tauri::Result<()> {
    let config = app.config().app.windows.first().cloned().expect("tauri.conf.json has no window");
    let window = WebviewWindowBuilder::from_config(app, &config)?.build()?;
    let geometry = app.state::<Arc<GeometryStore>>().inner().clone();
    if let Err(e) = place(&window, geometry.get()) {
        eprintln!("[deskboard] could not place the window: {e}");
    }
    watch_window(&window, geometry);

    // 창 순서를 다루는 일은 창을 만든 스레드에서 해야 한다.
    #[cfg(windows)]
    let handle = window.hwnd()?.0 as isize;
    let shown = window.clone();
    window.run_on_main_thread(move || {
        #[cfg(windows)]
        desktop::install(windows::Win32::Foundation::HWND(handle as _));
        let _ = shown.show();
    })
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

            geometry::spawn_flusher(geometry);
            app.manage(ClosedByUser::default());

            open_widget(app.handle())?;

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

    app.run(|app, event| match event {
        // 닫으라고 한 적이 없는데 창이 없어졌다: 위젯이 붙어 있던 바탕화면 창이
        // 탐색기와 함께 사라지면서 같이 없어진 것이다. 끝내지 않고 창을 다시 만든다.
        RunEvent::ExitRequested { code: None, api, .. }
            if !app.state::<ClosedByUser>().0.load(Ordering::Acquire) =>
        {
            api.prevent_exit();
            let app = app.clone();
            std::thread::spawn(move || {
                std::thread::sleep(Duration::from_millis(500));
                if let Err(e) = open_widget(&app) {
                    eprintln!("[deskboard] could not reopen the widget: {e}");
                    app.exit(1);
                }
            });
        }
        RunEvent::Exit => app.state::<Arc<GeometryStore>>().flush(),
        _ => {}
    });
}
