//! 다른 프로그램에서 끌어온 파일을 위젯이 직접 받는다.
//!
//! 웹뷰 라이브러리에 맡기면, 창이 만들어지는 순간에 있던 웹뷰 내부 창에만 "놓기 받기"가 등록된다.
//! 내부 창이 그보다 늦게 생기면 등록이 빠지는데, 그러면 파일을 끌어와도 금지 표시 커서만 뜬다
//! (Windows 11 에서 실행할 때마다 되다 안 되다 했다). 그래서 위젯이 직접 등록하고,
//! 내부 창이 새로 생기지 않았는지 계속 살펴 생기는 대로 등록한다.

use crate::desktop::log;
use std::cell::{Cell, RefCell};
use std::collections::HashSet;
use windows::core::{implement, BOOL};
use windows::Win32::Foundation::{HWND, LPARAM, POINT, POINTL};
use windows::Win32::Graphics::Gdi::ScreenToClient;
use windows::Win32::System::Com::{IDataObject, DVASPECT_CONTENT, FORMATETC, TYMED_HGLOBAL};
use windows::Win32::System::Ole::{
    IDropTarget, IDropTarget_Impl, RegisterDragDrop, ReleaseStgMedium, RevokeDragDrop, CF_HDROP,
    DROPEFFECT, DROPEFFECT_COPY, DROPEFFECT_NONE,
};
use windows::Win32::System::SystemServices::MODIFIERKEYS_FLAGS;
use windows::Win32::UI::Shell::{DragQueryFileW, HDROP};
use windows::Win32::UI::WindowsAndMessaging::{EnumChildWindows, SetTimer};

const TIMER_ID: usize = 0xDE5B_0010;
const CHECK_MS: u32 = 500;

/// 화면(프런트엔드)에 알리는 일. 좌표는 위젯 창 왼쪽 위 기준의 물리 픽셀이다.
#[derive(Clone, Debug, PartialEq, serde::Serialize)]
#[serde(tag = "type", rename_all = "lowercase")]
pub enum FileDrop {
    /// 파일을 끌고 위젯 위를 지나가는 중
    Over { x: i32, y: i32 },
    Drop { x: i32, y: i32, paths: Vec<String> },
    /// 놓지 않고 위젯 밖으로 나갔다
    Leave,
}

struct Watcher {
    widget: HWND,
    target: IDropTarget,
    /// 이미 등록해 둔 창들
    registered: HashSet<isize>,
}

thread_local! {
    // 놓기 받기는 창을 만든 스레드에서만 등록하고 불린다.
    static WATCHER: RefCell<Option<Watcher>> = const { RefCell::new(None) };
}

/// 위젯과 그 안의 모든 창에서 파일 놓기를 받기 시작한다. 창을 만든 스레드에서 한 번 부른다.
pub fn watch(widget: HWND, listener: impl Fn(FileDrop) + 'static) {
    let target: IDropTarget = Target {
        widget: widget.0 as isize,
        listener: Box::new(listener),
        accepting: Cell::new(false),
        last: Cell::new((i32::MIN, i32::MIN)),
    }
    .into();
    WATCHER.with(|watcher| {
        *watcher.borrow_mut() = Some(Watcher { widget, target, registered: HashSet::new() });
    });
    register_new_windows();
    unsafe {
        SetTimer(Some(widget), TIMER_ID, CHECK_MS, Some(on_timer));
    }
}

unsafe extern "system" fn on_timer(_window: HWND, _message: u32, _id: usize, _time: u32) {
    register_new_windows();
}

/// 아직 등록하지 않은 창(위젯 자신과 그 안의 창들)에 놓기 받기를 등록한다.
fn register_new_windows() {
    WATCHER.with(|watcher| {
        let mut watcher = watcher.borrow_mut();
        let Some(watcher) = watcher.as_mut() else { return };

        let mut windows: Vec<HWND> = vec![watcher.widget];
        unsafe extern "system" fn collect(window: HWND, out: LPARAM) -> BOOL {
            (*(out.0 as *mut Vec<HWND>)).push(window);
            true.into()
        }
        unsafe {
            let _ = EnumChildWindows(Some(watcher.widget), Some(collect), LPARAM(&mut windows as *mut _ as isize));
        }

        // 없어진 창은 잊는다. 같은 번호가 새 창에 다시 쓰일 수 있다.
        watcher.registered.retain(|known| windows.iter().any(|w| w.0 as isize == *known));
        for window in windows {
            if watcher.registered.contains(&(window.0 as isize)) {
                continue;
            }
            unsafe {
                // 웹뷰가 자기 것을 등록해 두었으면 걷어 내고 우리 것을 넣는다.
                let _ = RevokeDragDrop(window);
                match RegisterDragDrop(window, &watcher.target) {
                    Ok(()) => {
                        watcher.registered.insert(window.0 as isize);
                        log(|| format!("file drops are now accepted on window {:?}", window.0));
                    }
                    Err(e) => log(|| format!("could not accept file drops on window {:?}: {e}", window.0)),
                }
            }
        }
    });
}

#[implement(IDropTarget)]
struct Target {
    widget: isize,
    listener: Box<dyn Fn(FileDrop)>,
    /// 지금 끌고 있는 것이 파일이라 받아 줄 수 있는지
    accepting: Cell<bool>,
    /// 마지막으로 알린 자리. 같은 자리를 되풀이해 알리지 않는다.
    last: Cell<(i32, i32)>,
}

impl Target {
    /// 화면 좌표를 위젯 창 안의 좌표로 바꾼다.
    fn local(&self, point: &POINTL) -> (i32, i32) {
        let mut local = POINT { x: point.x, y: point.y };
        unsafe {
            let _ = ScreenToClient(HWND(self.widget as *mut core::ffi::c_void), &mut local);
        }
        (local.x, local.y)
    }
}

/// 끌려온 것에 들어 있는 파일·폴더 경로. 파일이 아니면(글자, 웹 링크 등) 비어 있다.
unsafe fn dragged_paths(data: &IDataObject) -> Vec<String> {
    let format = FORMATETC {
        cfFormat: CF_HDROP.0,
        ptd: std::ptr::null_mut(),
        dwAspect: DVASPECT_CONTENT.0,
        lindex: -1,
        tymed: TYMED_HGLOBAL.0 as u32,
    };
    let Ok(mut medium) = data.GetData(&format) else {
        return Vec::new();
    };
    let files = HDROP(medium.u.hGlobal.0 as _);
    let count = DragQueryFileW(files, u32::MAX, None);
    let mut paths = Vec::with_capacity(count as usize);
    for index in 0..count {
        let length = DragQueryFileW(files, index, None) as usize;
        let mut buffer = vec![0u16; length + 1];
        DragQueryFileW(files, index, Some(&mut buffer));
        paths.push(String::from_utf16_lossy(&buffer[..length]));
    }
    ReleaseStgMedium(&mut medium);
    paths
}

#[allow(non_snake_case)]
impl IDropTarget_Impl for Target_Impl {
    fn DragEnter(
        &self,
        data: windows::core::Ref<'_, IDataObject>,
        _keys: MODIFIERKEYS_FLAGS,
        point: &POINTL,
        effect: *mut DROPEFFECT,
    ) -> windows::core::Result<()> {
        let count = data.as_ref().map_or(0, |data| unsafe { dragged_paths(data).len() });
        self.accepting.set(count > 0);
        self.last.set((i32::MIN, i32::MIN));
        log(|| format!("something is dragged over the widget: {count} file(s)"));
        self.DragOver(_keys, point, effect)
    }

    fn DragOver(&self, _keys: MODIFIERKEYS_FLAGS, point: &POINTL, effect: *mut DROPEFFECT) -> windows::core::Result<()> {
        if self.accepting.get() {
            let at = self.local(point);
            if self.last.replace(at) != at {
                (self.listener)(FileDrop::Over { x: at.0, y: at.1 });
            }
        }
        unsafe {
            *effect = if self.accepting.get() { DROPEFFECT_COPY } else { DROPEFFECT_NONE };
        }
        Ok(())
    }

    fn DragLeave(&self) -> windows::core::Result<()> {
        if self.accepting.replace(false) {
            (self.listener)(FileDrop::Leave);
        }
        Ok(())
    }

    fn Drop(
        &self,
        data: windows::core::Ref<'_, IDataObject>,
        _keys: MODIFIERKEYS_FLAGS,
        point: &POINTL,
        effect: *mut DROPEFFECT,
    ) -> windows::core::Result<()> {
        let paths = data.as_ref().map_or_else(Vec::new, |data| unsafe { dragged_paths(data) });
        self.accepting.set(false);
        let accepted = !paths.is_empty();
        log(|| format!("dropped on the widget: {} file(s)", paths.len()));
        if accepted {
            let (x, y) = self.local(point);
            (self.listener)(FileDrop::Drop { x, y, paths });
        }
        unsafe {
            *effect = if accepted { DROPEFFECT_COPY } else { DROPEFFECT_NONE };
        }
        Ok(())
    }
}
