//! 위젯 창을 바탕화면에 붙여 둔다.
//!
//! 평소에는 모든 창의 맨 아래에 깔려 있어 다른 프로그램을 가리지 않는다.
//! Win+D(바탕화면 보기)를 누르면 셸이 바탕화면 창을 다른 창들 위로 끌어올려
//! 가리는데, 그동안에는 위젯을 그 바탕화면 창에 딸린 창(owned window)으로 만든다.
//! Windows 는 딸린 창을 언제나 주인 창 위에 두므로, 바탕화면을 클릭해도 위젯이 가려지지 않는다.
//!
//! 바탕화면이 올라왔는지는 맨 아래에 둔 보이지 않는 표식 창으로 판단한다.
//! 셸의 바탕화면 창이 표식 창보다 위에 있으면 바탕화면 보기 상태다.
//! 올라오는 창은 Windows 판마다 다르다 (Windows 10: WorkerW, Windows 11 24H2: Progman).
//!
//! 단축키나 트레이로 "불러오면" 위젯이 잠깐 다른 창들 위로("항상 위") 올라오고, 다른 프로그램으로
//! 돌아가면 다시 내려간다. 다만 다른 창에서 마우스를 누르고 있는 동안은 기다린다 —
//! 탐색기에서 파일을 집어 위젯으로 끌어오는 중일 수 있기 때문이다.
//!
//! 주인 창이 다른 프로세스의 것이라, 탐색기가 죽어도 위젯 창은 없어지지 않고 주인만 잃는다.
//! 탐색기가 다시 뜨면 새 바탕화면 창을 기준으로 같은 일을 계속한다.

use std::io::Write;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, Ordering::Relaxed};
use std::sync::OnceLock;
use windows::core::w;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::UI::Accessibility::{SetWinEventHook, HWINEVENTHOOK};
use windows::Win32::UI::Input::KeyboardAndMouse::{GetAsyncKeyState, VK_LBUTTON, VK_RBUTTON};
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
/// 지금 위젯이 딸려 있는 바탕화면 창. 0 이면 평소 상태(맨 아래)다.
static ATTACHED_TO: AtomicIsize = AtomicIsize::new(0);
static MOVING_SELF: AtomicBool = AtomicBool::new(false);
/// 불러와서 다른 창들 위에 올라와 있는 중인지
static SUMMONED: AtomicBool = AtomicBool::new(false);
/// 불러온 뒤 위젯이 실제로 앞 창이 된 적이 있는지. 그 전에는 "다른 창으로 갔다"고 판단하지 않는다.
static SUMMON_TOOK_FOCUS: AtomicBool = AtomicBool::new(false);
/// 불러오기 전에 쓰고 있던 창. 다시 내려갈 때 포커스를 돌려준다.
static RETURN_TO: AtomicIsize = AtomicIsize::new(0);
static FAST_LEFT: AtomicU32 = AtomicU32::new(0);

fn hwnd(value: &AtomicIsize) -> HWND {
    HWND(value.load(Relaxed) as *mut core::ffi::c_void)
}

/// 창을 바탕화면에 붙이고, 활성화하지 않은 채로 보인다. 창을 만든 스레드에서 한 번 부른다.
/// 이어서 `sync_visibility` 를 불러야 한다.
pub fn install(widget: HWND) {
    unsafe {
        // Alt+Tab 목록과 작업 표시줄에 나타나지 않게 한다.
        let ex = GetWindowLongPtrW(widget, GWL_EXSTYLE);
        SetWindowLongPtrW(widget, GWL_EXSTYLE, tool_window_style(ex));

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
        ATTACHED_TO.store(0, Relaxed);

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

        let _ = ShowWindow(widget, SW_SHOWNOACTIVATE);
        evaluate();
    }
}

/// 창 틀에게도 창이 보인다는 것을 알린다 (`show` 가 그 일을 한다).
///
/// 창 틀은 자기 기록에 "숨김"으로 남아 있는 창을, 창 모양을 바꿀 일이 생길 때마다 도로 숨긴다.
/// 그런데 창 틀을 통해 보이게 하면 이미 보이는 창이어도 활성화되므로, 쓰던 프로그램에 포커스를 돌려준다.
pub fn sync_visibility(show: impl FnOnce()) {
    unsafe {
        let widget = hwnd(&WIDGET);
        let before = GetForegroundWindow();
        show();
        if GetForegroundWindow() == widget && before != widget && !before.is_invalid() {
            let returned = SetForegroundWindow(before).as_bool();
            log(|| format!("showing the widget took the focus; gave it back: {returned}"));
        }
    }
}

/// 위젯을 다른 창들 위로 불러온다. 이미 불러와 있으면 도로 내려보낸다. 창을 만든 스레드에서 부른다.
pub fn toggle_summon() {
    unsafe {
        let widget = hwnd(&WIDGET);
        if widget.is_invalid() {
            return;
        }
        if SUMMONED.load(Relaxed) {
            dismiss(true);
            return;
        }
        RETURN_TO.store(GetForegroundWindow().0 as isize, Relaxed);
        SUMMONED.store(true, Relaxed);
        // 앞 창이 되면서 창 순서도 맨 위로 올라가야 하므로, 이때만 순서 변경을 받아들인다.
        MOVING_SELF.store(true, Relaxed);
        let _ = SetForegroundWindow(widget);
        // "항상 위"로 둬야, 다른 창에서 파일을 집으려고 그 창을 눌러도 위젯이 가려지지 않는다.
        let _ = SetWindowPos(widget, Some(HWND_TOPMOST), 0, 0, 0, 0, Z_ONLY);
        MOVING_SELF.store(false, Relaxed);
        // 바로 앞 창이 되었으면 지금 적어 둔다. 다음 확인 때까지 미루면, 그 사이에 사용자가
        // 다른 창을 눌렀을 때 물러날 때를 놓쳐 "불러온 상태"로 남는다.
        let focused = GetForegroundWindow() == widget;
        SUMMON_TOOK_FOCUS.store(focused, Relaxed);
        log(|| format!("summoned (took the focus: {focused}, always on top: {})", is_topmost(widget)));
    }
}

/// 불러온 위젯을 제자리(맨 아래, 또는 보이고 있는 바탕화면 위)로 돌려보낸다.
unsafe fn dismiss(give_focus_back: bool) {
    let widget = hwnd(&WIDGET);
    SUMMONED.store(false, Relaxed);
    set_order(widget, HWND_NOTOPMOST);
    if hwnd(&ATTACHED_TO).is_invalid() {
        set_order(widget, HWND_BOTTOM);
    }
    let previous = hwnd(&RETURN_TO);
    if give_focus_back && previous != widget && IsWindow(Some(previous)).as_bool() {
        let _ = SetForegroundWindow(previous);
    }
    log(|| "dismissed".into());
}

fn tool_window_style(ex_style: isize) -> isize {
    (ex_style | WS_EX_TOOLWINDOW.0 as isize) & !(WS_EX_APPWINDOW.0 as isize)
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
        // 창 틀(tao)은 창 모양을 바꿀 때(크기 조절 허용 등) 확장 스타일을 자기 기록대로 통째로
        // 다시 써서 위에서 넣은 도구 창 표시가 지워지므로, 바뀌는 순간에 다시 넣는다.
        WM_STYLECHANGING if wparam.0 as i32 == GWL_EXSTYLE.0 => {
            let styles = &mut *(lparam.0 as *mut STYLESTRUCT);
            styles.styleNew = tool_window_style(styles.styleNew as isize) as u32;
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
            ATTACHED_TO.store(0, Relaxed);
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

/// 바탕화면 보기 상태가 바뀌었으면 위젯을 붙이거나 뗀다.
unsafe fn evaluate() {
    let widget = hwnd(&WIDGET);
    if widget.is_invalid() {
        return;
    }
    if IsIconic(widget).as_bool() {
        let _ = ShowWindow(widget, SW_SHOWNOACTIVATE);
    }
    if SUMMONED.load(Relaxed) {
        // 불러온 위젯은 다른 프로그램의 창이 앞에 나오면 물러난다.
        // (위젯이 띄운 파일 선택 창 같은 우리 창은 예외)
        // 마우스를 누르고 있는 동안은 물러나지 않는다: 그 창에서 파일을 집어 위젯으로 끌어오는 중일 수 있다.
        let foreground = GetForegroundWindow();
        if !foreground.is_invalid() && process_of(foreground) == process_of(widget) {
            SUMMON_TOOK_FOCUS.store(true, Relaxed);
        } else if SUMMON_TOOK_FOCUS.load(Relaxed) && !mouse_button_held() {
            dismiss(false);
        }
    }
    let attached_to = hwnd(&ATTACHED_TO);
    match raised_desktop() {
        Some(desktop) => {
            if attached_to != desktop || !is_above(widget, desktop) {
                // 바탕화면 창의 딸린 창이 되어 그 위에 선다.
                ATTACHED_TO.store(desktop.0 as isize, Relaxed);
                SetWindowLongPtrW(widget, GWLP_HWNDPARENT, desktop.0 as isize);
                set_order(widget, HWND_TOP);
                log(|| {
                    format!(
                        "desktop {:?} [{}] is shown -> attached, widget above it: {}",
                        desktop.0,
                        class_name(desktop),
                        is_above(widget, desktop),
                    )
                });
            }
        }
        None if !attached_to.is_invalid() => {
            ATTACHED_TO.store(0, Relaxed);
            SetWindowLongPtrW(widget, GWLP_HWNDPARENT, 0);
            if !SUMMONED.load(Relaxed) {
                set_order(widget, HWND_BOTTOM);
            }
            log(|| "desktop is back down -> detached, widget at the bottom".into());
        }
        None => {}
    }
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

unsafe fn is_topmost(window: HWND) -> bool {
    GetWindowLongPtrW(window, GWL_EXSTYLE) & WS_EX_TOPMOST.0 as isize != 0
}

/// 마우스 단추(왼쪽이든 오른쪽이든)가 지금 눌려 있는지
unsafe fn mouse_button_held() -> bool {
    GetAsyncKeyState(VK_LBUTTON.0 as i32) < 0 || GetAsyncKeyState(VK_RBUTTON.0 as i32) < 0
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

/// 환경 변수 DESKBOARD_LOG 에 파일 경로를 주면 위젯을 붙이고 뗄 때, 파일을 끌어올 때마다 한 줄씩 남긴다.
/// 바탕화면 보기 동작은 Windows 판마다 달라서, 안 보인다는 제보가 왔을 때 원인을 볼 수 있게 해 둔다.
pub(crate) fn log(message: impl FnOnce() -> String) {
    static PATH: OnceLock<Option<PathBuf>> = OnceLock::new();
    let Some(path) = PATH.get_or_init(|| std::env::var_os("DESKBOARD_LOG").map(PathBuf::from)) else {
        return;
    };
    if let Ok(mut file) = std::fs::OpenOptions::new().create(true).append(true).open(path) {
        let _ = writeln!(file, "{}", message());
    }
}
