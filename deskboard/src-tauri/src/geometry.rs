//! 위젯 창의 위치와 크기를 기억했다가 다음 실행 때 되돌린다.

use crate::store::write_atomic;
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::Duration;

/// 기본 크기 (배율 100% 기준). tauri.conf.json 의 창 크기와 맞춘다.
pub const DEFAULT_W: f64 = 780.0;
pub const DEFAULT_H: f64 = 460.0;
const MARGIN: f64 = 48.0;
/// 저장된 위치가 화면에 이만큼도 걸치지 않으면 잡을 수 없다고 보고 기본 위치로 돌린다.
const MIN_VISIBLE_W: i64 = 120;
const MIN_VISIBLE_H: i64 = 80;

/// 물리 픽셀 기준 사각형
#[derive(Serialize, Deserialize, Clone, Copy, Debug, PartialEq)]
pub struct Rect {
    pub x: i32,
    pub y: i32,
    pub width: u32,
    pub height: u32,
}

/// 주 모니터 오른쪽 위에 기본 크기로 놓는다. 화면보다 크면 화면에 맞춘다.
pub fn default_rect(primary: Rect, scale: f64) -> Rect {
    let margin = (MARGIN * scale).round() as i64;
    let fit = |want: f64, room: u32| -> u32 {
        let room = (room as i64 - 2 * margin).max(200);
        ((want * scale).round() as i64).min(room) as u32
    };
    let width = fit(DEFAULT_W, primary.width);
    let height = fit(DEFAULT_H, primary.height);
    Rect {
        x: (primary.x as i64 + primary.width as i64 - width as i64 - margin) as i32,
        y: (primary.y as i64 + margin) as i32,
        width,
        height,
    }
}

/// 저장된 위치가 지금 연결된 모니터 어딘가에 충분히 보이면 그대로 쓰고,
/// 아니면(모니터를 뗐거나 해상도가 바뀐 경우) 기본 위치를 쓴다.
pub fn resolve(saved: Option<Rect>, monitors: &[Rect], primary: Rect, scale: f64) -> Rect {
    match saved {
        Some(rect) if monitors.iter().any(|m| visible_enough(rect, *m)) => rect,
        _ => default_rect(primary, scale),
    }
}

fn visible_enough(rect: Rect, monitor: Rect) -> bool {
    let overlap = |a0: i32, a_len: u32, b0: i32, b_len: u32| -> i64 {
        let start = (a0 as i64).max(b0 as i64);
        let end = (a0 as i64 + a_len as i64).min(b0 as i64 + b_len as i64);
        end - start
    };
    rect.width > 0
        && rect.height > 0
        && overlap(rect.x, rect.width, monitor.x, monitor.width) >= MIN_VISIBLE_W.min(rect.width as i64)
        && overlap(rect.y, rect.height, monitor.y, monitor.height) >= MIN_VISIBLE_H.min(rect.height as i64)
}

pub struct GeometryStore {
    path: PathBuf,
    current: Mutex<Option<Rect>>,
    dirty: AtomicBool,
}

impl GeometryStore {
    pub fn load(path: PathBuf) -> Self {
        let current =
            fs::read_to_string(&path).ok().and_then(|text| serde_json::from_str::<Rect>(&text).ok());
        GeometryStore { path, current: Mutex::new(current), dirty: AtomicBool::new(false) }
    }

    pub fn get(&self) -> Option<Rect> {
        *self.current.lock().unwrap()
    }

    pub fn set(&self, rect: Rect) {
        let mut current = self.current.lock().unwrap();
        if *current != Some(rect) {
            *current = Some(rect);
            self.dirty.store(true, Ordering::Release);
        }
    }

    /// 창을 끄는 동안 이벤트가 쏟아지므로, 바뀐 내용은 모아 두었다가 한 번에 쓴다.
    pub fn flush(&self) {
        if !self.dirty.swap(false, Ordering::AcqRel) {
            return;
        }
        if let Some(rect) = self.get() {
            if let Ok(json) = serde_json::to_string(&rect) {
                let _ = write_atomic(&self.path, json.as_bytes());
            }
        }
    }
}

pub fn spawn_flusher(store: Arc<GeometryStore>) {
    std::thread::spawn(move || loop {
        std::thread::sleep(Duration::from_secs(1));
        store.flush();
    });
}

#[cfg(test)]
mod tests {
    use super::*;

    const FHD: Rect = Rect { x: 0, y: 0, width: 1920, height: 1080 };

    #[test]
    fn default_sits_at_the_top_right_of_the_primary_monitor() {
        let rect = default_rect(FHD, 1.0);
        assert_eq!(rect, Rect { x: 1920 - 780 - 48, y: 48, width: 780, height: 460 });
    }

    #[test]
    fn default_scales_with_display_scaling() {
        let rect = default_rect(Rect { x: 0, y: 0, width: 3840, height: 2160 }, 2.0);
        assert_eq!((rect.width, rect.height), (1560, 920));
        assert_eq!(rect.x + rect.width as i32, 3840 - 96);
    }

    #[test]
    fn default_shrinks_to_fit_a_small_screen() {
        let small = Rect { x: 0, y: 0, width: 800, height: 480 };
        let rect = default_rect(small, 1.0);
        assert!(rect.x >= 0 && rect.y >= 0);
        assert!(rect.x + rect.width as i32 <= 800);
        assert!(rect.y + rect.height as i32 <= 480);
    }

    #[test]
    fn saved_position_on_a_connected_monitor_is_kept() {
        let second = Rect { x: -1920, y: 0, width: 1920, height: 1080 };
        let saved = Rect { x: -1500, y: 200, width: 600, height: 400 };
        assert_eq!(resolve(Some(saved), &[FHD, second], FHD, 1.0), saved);
    }

    #[test]
    fn saved_position_on_a_removed_monitor_falls_back_to_default() {
        let saved = Rect { x: -1500, y: 200, width: 600, height: 400 };
        assert_eq!(resolve(Some(saved), &[FHD], FHD, 1.0), default_rect(FHD, 1.0));
    }

    #[test]
    fn barely_visible_sliver_counts_as_lost() {
        let saved = Rect { x: 1900, y: 100, width: 600, height: 400 };
        assert_eq!(resolve(Some(saved), &[FHD], FHD, 1.0), default_rect(FHD, 1.0));
    }

    #[test]
    fn nothing_saved_uses_default() {
        assert_eq!(resolve(None, &[FHD], FHD, 1.0), default_rect(FHD, 1.0));
    }

    #[test]
    fn flush_writes_only_what_changed_and_reload_reads_it_back() {
        let dir = std::env::temp_dir().join(format!("deskboard-geo-{}", std::process::id()));
        let _ = fs::remove_dir_all(&dir);
        let path = dir.join("window.json");
        let store = GeometryStore::load(path.clone());
        assert_eq!(store.get(), None);
        store.flush();
        assert!(!path.exists());

        let rect = Rect { x: 10, y: 20, width: 300, height: 200 };
        store.set(rect);
        store.flush();
        assert_eq!(GeometryStore::load(path).get(), Some(rect));
    }
}
