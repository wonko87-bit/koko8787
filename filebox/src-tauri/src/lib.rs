mod commands;
mod model;
mod rules;
mod store;
mod suggest;
mod toast;
mod watcher;

use model::Settings;
use std::sync::Mutex;
use store::Store;
use tauri::menu::{Menu, MenuItem};
use tauri::tray::{MouseButton, MouseButtonState, TrayIconBuilder, TrayIconEvent};
use tauri::{Emitter, Manager, WindowEvent};

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_notification::init())
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_process::init())
        .plugin(tauri_plugin_autostart::init(
            tauri_plugin_autostart::MacosLauncher::LaunchAgent,
            None,
        ))
        .setup(|app| {
            // ---- 스토어 로드 & 기본 설정 ----
            let data_dir = app.path().app_data_dir()?;
            let store = Store::load(data_dir.join("store.json"));

            let download_dir = app
                .path()
                .download_dir()
                .unwrap_or_else(|_| std::env::temp_dir());
            let inbox_dir = app
                .path()
                .home_dir()
                .map(|h| h.join("FileBox"))
                .unwrap_or_else(|_| std::env::temp_dir().join("FileBox"));
            store.update(|d| {
                if d.settings.is_none() {
                    d.settings = Some(Settings::new(download_dir, inbox_dir));
                }
            });
            app.manage(store);

            // ---- 감시 스레드 ----
            let tx = watcher::spawn(app.handle().clone());
            app.manage(watcher::WatcherCtl(Mutex::new(tx)));

            // ---- 토스트 창 (숨김 상태로 준비) ----
            tauri::WebviewWindowBuilder::new(
                app,
                "toast",
                tauri::WebviewUrl::App("toast.html".into()),
            )
            .title("FileBox 알림")
            .inner_size(380.0, 300.0)
            .decorations(false)
            .resizable(false)
            .always_on_top(true)
            .skip_taskbar(true)
            .visible(false)
            .build()?;

            // ---- 메인 창 닫기 → 트레이로 숨김 ----
            if let Some(main) = app.get_webview_window("main") {
                let main2 = main.clone();
                main.on_window_event(move |event| {
                    if let WindowEvent::CloseRequested { api, .. } = event {
                        api.prevent_close();
                        let _ = main2.hide();
                    }
                });
            }

            // ---- 트레이 아이콘 ----
            let open_item = MenuItem::with_id(app, "open", "FileBox 열기", true, None::<&str>)?;
            let pause_item =
                MenuItem::with_id(app, "pause", "감시 일시정지/재개", true, None::<&str>)?;
            let quit_item = MenuItem::with_id(app, "quit", "종료", true, None::<&str>)?;
            let menu = Menu::with_items(app, &[&open_item, &pause_item, &quit_item])?;

            // 트레이 생성 실패(리눅스에 appindicator 부재 등)는 치명적이지 않게 처리
            let tray_result = TrayIconBuilder::with_id("filebox-tray")
                .icon(app.default_window_icon().unwrap().clone())
                .tooltip("FileBox — 파일 관리함")
                .menu(&menu)
                .show_menu_on_left_click(false)
                .on_menu_event(|app, event| match event.id.as_ref() {
                    "open" => show_main(app),
                    "pause" => {
                        let store = app.state::<Store>();
                        store.update(|d| {
                            if let Some(s) = d.settings.as_mut() {
                                s.paused = !s.paused;
                            }
                        });
                        let _ = app.emit("settings-changed", ());
                    }
                    "quit" => app.exit(0),
                    _ => {}
                })
                .on_tray_icon_event(|tray, event| {
                    if let TrayIconEvent::Click {
                        button: MouseButton::Left,
                        button_state: MouseButtonState::Up,
                        ..
                    } = event
                    {
                        show_main(tray.app_handle());
                    }
                })
                .build(app);
            if let Err(e) = tray_result {
                eprintln!("[filebox] tray init failed (continuing without tray): {e}");
            }

            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            commands::get_settings,
            commands::update_settings,
            commands::list_entries,
            commands::get_suggestions,
            commands::send_to_favorite,
            commands::send_to_path,
            commands::set_category,
            commands::set_tags,
            commands::remove_entry,
            commands::remove_entries,
            commands::set_category_many,
            commands::send_many_to_favorite,
            commands::send_many_to_path,
            commands::undo_move,
            commands::clear_history,
            commands::list_favorites,
            commands::add_favorite,
            commands::remove_favorite,
            commands::list_rules,
            commands::upsert_rule,
            commands::remove_rule,
            commands::list_uncollected,
            commands::collect_paths,
            commands::open_entry,
            commands::open_folder,
            commands::reveal_entry,
            commands::hide_toast,
            commands::show_main_window,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

fn show_main(app: &tauri::AppHandle) {
    if let Some(win) = app.get_webview_window("main") {
        let _ = win.show();
        let _ = win.unminimize();
        let _ = win.set_focus();
    }
}
