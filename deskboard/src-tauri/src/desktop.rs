//! 위젯 창을 바탕화면에 붙여 둔다.
//!
//! 평소에는 모든 창의 맨 아래에 깔려 있어 다른 프로그램을 가리지 않는다.
//! Win+D(바탕화면 보기)를 누르면 셸이 바탕화면 창을 다른 창들 위로 끌어올려
//! 가리는데, 그동안에만 위젯을 "항상 위"로 올려 바탕화면 위에 보이게 한다.
//! (Rainmeter 의 "바탕화면에 고정"과 같은 방식)
//!
//! 바탕화면이 올라왔는지는 맨 아래에 둔 보이지 않는 표식 창으로 판단한다.
//! 셸의 바탕화면 창이 표식 창보다 위에 있으면 바탕화면 보기 상태다.

use std::io::Write;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, Ordering::Relaxed};
use std::sync::OnceLock;
use windows::core::w;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::UI::Accessibility::{SetWinEventHook, HWINEVENTHOOK};
use windows::Win32::UI::Shell::{DefSubclassProc, RemoveWindowSubclass, SetWindowSubclass};
use windows::Win32::UI::WindowsAndMessaging::*;

const SUBCLASS_ID: usize = 0xDE5B;
/// 놓친 변화가 있어도 곧 바로잡도록 느리게 계속 확인한다.
const TIMER_POLL: usize = 0xDE5B_0001;
const POLL_MS: u32 = 300;
/// 앞에 나온 창이 바뀐 직후에는 창 순서가 조금 늦게 바뀌므로 잠깐 촘촘히 확인한다.
const TIMER_FAST: usize = 0xDE5B_0002;
const FAST_MS: u32 = 40;
const FAST_TICKS: u32 = 12;

const Z_ONLY: SET_WINDOW_POS_FLAGS =
    SET_WINDOW_POS_FLAGS(SWP_NOMOVE.0 | SWP_NOSIZE.0 | SWP_NOACTIVATE.0);

// 아래 값들은 모두 창을 만든 스레드(메인 스레드)에서만 읽고 쓴다.
static WIDGET: AtomicIsize = AtomicIsize::new(0);
static MARKER: AtomicIsize = AtomicIsize::new(0);
static RAISED: AtomicBool = AtomicBool::new(false);
static MOVING_SELF: AtomicBool = AtomicBool::new(false);
static FAST_LEFT: AtomicU32 = AtomicU32::new(0);

fn hwnd(value: &AtomicIsize) -> HWND {
    HWND(value.load(Relaxed) as *mut core::ffi::c_void)
}

/// 창을 바탕화면에 붙이고 화면에 보인다. 창을 만든 스레드에서 한 번만 부른다.
pub fn install(widget: HWND) {
    unsafe {
        // Alt+Tab 목록과 작업 표시줄에 나타나지 않게 한다.
        let ex = GetWindowLongPtrW(widget, GWL_EXSTYLE);
        let ex = (ex | WS_EX_TOOLWINDOW.0 as isize) & !(WS_EX_APPWINDOW.0 as isize);
        SetWindowLongPtrW(widget, GWL_EXSTYLE, ex);

        let marker = CreateWindowExW(
            WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
            w!("Static"),
            w!("DeskBoardMarker"),
            WS_POPUP,
            0,
            0,
            0,
            0,
            None,
            None,
            None,
            None,
        )
        .unwrap_or_default();

        WIDGET.store(widget.0 as isize, Relaxed);
        MARKER.store(marker.0 as isize, Relaxed);

        let _ = SetWindowPos(marker, Some(HWND_BOTTOM), 0, 0, 0, 0, Z_ONLY);
        let _ = SetWindowPos(widget, Some(HWND_BOTTOM), 0, 0, 0, 0, Z_ONLY);

        let _ = SetWindowSubclass(widget, Some(subclass_proc), SUBCLASS_ID, 0);
        SetTimer(Some(widget), TIMER_POLL, POLL_MS, None);
        SetWinEventHook(
            EVENT_SYSTEM_FOREGROUND,
            EVENT_SYSTEM_FOREGROUND,
            None,
            Some(foreground_changed),
            0,
            0,
            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS,
        );

        // 쓰던 프로그램의 포커스를 뺏지 않고 조용히 나타난다.
        let _ = ShowWindow(widget, SW_SHOWNOACTIVATE);
        evaluate();
    }
}

unsafe extern "system" fn subclass_proc(
    window: HWND,
    message: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    _id: usize,
    _data: usize,
) -> LRESULT {
    match message {
        // 클릭해서 활성화되어도 앞으로 튀어나오지 않게, 우리가 정한 순서만 받아들인다.
        WM_WINDOWPOSCHANGING if !MOVING_SELF.load(Relaxed) => {
            let pos = &mut *(lparam.0 as *mut WINDOWPOS);
            pos.flags |= SWP_NOZORDER;
        }
        WM_WINDOWPOSCHANGING | WM_WINDOWPOSCHANGED if MOVING_SELF.load(Relaxed) => {
            let pos = &*(lparam.0 as *const WINDOWPOS);
            let name = if message == WM_WINDOWPOSCHANGING { "changing" } else { "changed" };
            log(|| format!("  {name}: insert_after={:?} flags={:#x}", pos.hwndInsertAfter.0, pos.flags.0));
        }
        WM_TIMER if wparam.0 == TIMER_POLL => {
            evaluate();
            return LRESULT(0);
        }
        WM_TIMER if wparam.0 == TIMER_FAST => {
            evaluate();
            if FAST_LEFT.load(Relaxed) <= 1 {
                FAST_LEFT.store(0, Relaxed);
                let _ = KillTimer(Some(window), TIMER_FAST);
            } else {
                FAST_LEFT.fetch_sub(1, Relaxed);
            }
            return LRESULT(0);
        }
        WM_NCDESTROY => {
            let _ = KillTimer(Some(window), TIMER_POLL);
            let _ = KillTimer(Some(window), TIMER_FAST);
            let _ = RemoveWindowSubclass(window, Some(subclass_proc), SUBCLASS_ID);
            let _ = DestroyWindow(hwnd(&MARKER));
            WIDGET.store(0, Relaxed);
            MARKER.store(0, Relaxed);
        }
        _ => {}
    }
    DefSubclassProc(window, message, wparam, lparam)
}

unsafe extern "system" fn foreground_changed(
    _hook: HWINEVENTHOOK,
    _event: u32,
    _window: HWND,
    _object: i32,
    _child: i32,
    _thread: u32,
    _time: u32,
) {
    let widget = hwnd(&WIDGET);
    if widget.is_invalid() {
        return;
    }
    evaluate();
    FAST_LEFT.store(FAST_TICKS, Relaxed);
    SetTimer(Some(widget), TIMER_FAST, FAST_MS, None);
}

/// 바탕화면 보기 상태가 바뀌었으면 위젯의 창 순서를 맞춘다.
unsafe fn evaluate() {
    let widget = hwnd(&WIDGET);
    if widget.is_invalid() {
        return;
    }
    if IsIconic(widget).as_bool() {
        let _ = ShowWindow(widget, SW_SHOWNOACTIVATE);
    }
    let was_raised = RAISED.load(Relaxed);
    match raised_desktop() {
        Some(desktop) => {
            if !was_raised || !is_above(widget, desktop) {
                RAISED.store(true, Relaxed);
                lift_above(widget, desktop);
            }
        }
        None if was_raised => {
            RAISED.store(false, Relaxed);
            SetWindowLongPtrW(widget, GWLP_HWNDPARENT, 0);
            set_order(widget, HWND_NOTOPMOST);
            set_order(widget, HWND_BOTTOM);
            log(|| format!("desktop is back down -> widget topmost: {}", is_topmost(widget)));
        }
        None => {}
    }
}

/// 위젯을 올라온 바탕화면 창 위로 올린다. (어느 방법이 통하는지 기록하며 차례로 시도)
unsafe fn lift_above(widget: HWND, desktop: HWND) {
    let report = |step: &str| {
        log(|| {
            format!(
                "lift over {:?} [{}] by {step}: above={}, topmost={}, desktop topmost={}",
                desktop.0,
                class_name(desktop),
                is_above(widget, desktop),
                is_topmost(widget),
                is_topmost(desktop),
            )
        });
        is_above(widget, desktop)
    };

    set_order(widget, HWND_TOPMOST);
    if report("HWND_TOPMOST") {
        return;
    }
    set_order(widget, HWND_TOP);
    if report("HWND_TOP") {
        return;
    }
    if let Ok(above_desktop) = GetWindow(desktop, GW_HWNDPREV) {
        set_order(widget, above_desktop);
        if report("insert just above the desktop") {
            return;
        }
    }
    SetWindowLongPtrW(widget, GWLP_HWNDPARENT, desktop.0 as isize);
    set_order(widget, HWND_TOP);
    report("owner = desktop, then HWND_TOP");
}

/// 표식 창보다 위로 올라온 셸 바탕화면 창을 찾는다. 없으면 평소 상태다.
unsafe fn raised_desktop() -> Option<HWND> {
    let marker = hwnd(&MARKER);
    let shell = GetShellWindow();
    if marker.is_invalid() || shell.is_invalid() {
        return None;
    }
    let shell_process = process_of(shell);
    let mut current = GetTopWindow(None).ok();
    // 창 순서가 도는 도중에 바뀌어도 끝없이 돌지 않게 횟수를 막아 둔다.
    for _ in 0..4096 {
        let window = current?;
        if window == marker {
            return None;
        }
        if IsWindowVisible(window).as_bool()
            && process_of(window) == shell_process
            && is_desktop_class(window)
            && covers_a_screen(window)
        {
            return Some(window);
        }
        current = GetWindow(window, GW_HWNDNEXT).ok();
    }
    None
}

unsafe fn process_of(window: HWND) -> u32 {
    let mut process = 0u32;
    GetWindowThreadProcessId(window, Some(&mut process));
    process
}

unsafe fn class_name(window: HWND) -> String {
    let mut buffer = [0u16; 64];
    let len = GetClassNameW(window, &mut buffer) as usize;
    String::from_utf16_lossy(&buffer[..len.min(buffer.len())])
}

unsafe fn is_desktop_class(window: HWND) -> bool {
    let name = class_name(window);
    name == "WorkerW" || name == "Progman"
}

unsafe fn is_topmost(window: HWND) -> bool {
    GetWindowLongPtrW(window, GWL_EXSTYLE) & WS_EX_TOPMOST.0 as isize != 0
}

unsafe fn covers_a_screen(window: HWND) -> bool {
    let mut rect = RECT::default();
    GetWindowRect(window, &mut rect).is_ok()
        && rect.right - rect.left >= 200
        && rect.bottom - rect.top >= 200
}

unsafe fn is_above(window: HWND, other: HWND) -> bool {
    let mut current = GetWindow(other, GW_HWNDPREV).ok();
    for _ in 0..4096 {
        match current {
            Some(candidate) if candidate == window => return true,
            Some(candidate) => current = GetWindow(candidate, GW_HWNDPREV).ok(),
            None => return false,
        }
    }
    false
}

unsafe fn set_order(window: HWND, insert_after: HWND) {
    MOVING_SELF.store(true, Relaxed);
    if let Err(e) = SetWindowPos(window, Some(insert_after), 0, 0, 0, 0, Z_ONLY) {
        log(|| format!("SetWindowPos(insert_after={:?}) failed: {e}", insert_after.0));
    }
    MOVING_SELF.store(false, Relaxed);
}

/// 환경 변수 DESKBOARD_LOG 에 파일 경로를 주면 창 순서를 바꿀 때마다 한 줄씩 남긴다.
/// 바탕화면 보기 동작은 Windows 판마다 달라서, 안 보인다는 제보가 왔을 때 원인을 볼 수 있게 해 둔다.
fn log(message: impl FnOnce() -> String) {
    static PATH: OnceLock<Option<PathBuf>> = OnceLock::new();
    let Some(path) = PATH.get_or_init(|| std::env::var_os("DESKBOARD_LOG").map(PathBuf::from)) else {
        return;
    };
    if let Ok(mut file) = std::fs::OpenOptions::new().create(true).append(true).open(path) {
        let _ = writeln!(file, "{}", message());
    }
}
