//! 타일이 가리키는 대상을 열고, 대상이 어떤 종류인지 알아낸다.

use serde::Serialize;
use std::path::Path;

#[derive(Serialize, Debug, PartialEq)]
pub struct TargetInfo {
    /// folder | program | file | link
    pub kind: &'static str,
    /// 경로가 지금 실제로 있는지 (링크는 확인하지 않고 true)
    pub exists: bool,
}

const PROGRAM_EXTENSIONS: [&str; 7] = ["exe", "lnk", "bat", "cmd", "com", "msi", "appref-ms"];

/// `https://…`, `mailto:…`, `ms-settings:…`, `shell:startup` 처럼 "이름:" 으로 시작하면 링크로 본다.
/// 드라이브 문자(`C:\…`)와 헷갈리지 않도록 이름은 두 글자 이상이어야 한다.
pub fn is_link(target: &str) -> bool {
    match target.split_once(':') {
        Some((scheme, rest)) => {
            scheme.len() >= 2
                && !rest.is_empty()
                && scheme.starts_with(|c: char| c.is_ascii_alphabetic())
                && scheme.chars().all(|c| c.is_ascii_alphanumeric() || matches!(c, '+' | '-' | '.'))
        }
        None => false,
    }
}

pub fn classify(target: &str) -> TargetInfo {
    let target = target.trim();
    if is_link(target) {
        return TargetInfo { kind: "link", exists: true };
    }
    let path = Path::new(target);
    let extension = path.extension().map(|e| e.to_string_lossy().to_lowercase());
    let by_name = match extension.as_deref() {
        Some(ext) if PROGRAM_EXTENSIONS.contains(&ext) => "program",
        Some(_) => "file",
        None => "folder",
    };
    match std::fs::metadata(path) {
        Ok(meta) if meta.is_dir() => TargetInfo { kind: "folder", exists: true },
        Ok(_) => TargetInfo { kind: if by_name == "folder" { "file" } else { by_name }, exists: true },
        Err(_) => TargetInfo { kind: by_name, exists: false },
    }
}

/// 탐색기에서 더블클릭한 것과 같은 방식으로 연다.
/// `elevated` 면 "관리자 권한으로 실행"과 같아서, Windows 가 먼저 확인 창(UAC)을 띄운다.
#[cfg(windows)]
pub fn open(target: &str, elevated: bool) -> Result<(), String> {
    use std::os::windows::ffi::OsStrExt;
    use windows::core::{w, PCWSTR};
    use windows::Win32::System::Com::{
        CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE,
    };
    use windows::Win32::UI::Shell::ShellExecuteW;
    use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;

    let target = target.trim();
    if target.is_empty() {
        return Err("열 대상이 비어 있습니다.".into());
    }
    let wide = |s: &std::ffi::OsStr| -> Vec<u16> { s.encode_wide().chain(Some(0)).collect() };
    let file = wide(target.as_ref());
    // 프로그램이 자기 폴더를 기준으로 파일을 찾는 경우가 있어, 탐색기처럼 대상의 폴더에서 시작한다.
    let directory = if is_link(target) {
        None
    } else {
        Path::new(target).parent().filter(|p| p.is_dir()).map(|p| wide(p.as_os_str()))
    };

    let code = unsafe {
        // 셸 확장이 COM 을 쓰므로 이 스레드에서 초기화해 둔다.
        let com = CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
        let result = ShellExecuteW(
            None,
            if elevated { w!("runas") } else { PCWSTR::null() },
            PCWSTR(file.as_ptr()),
            PCWSTR::null(),
            directory.as_ref().map_or(PCWSTR::null(), |d| PCWSTR(d.as_ptr())),
            SW_SHOWNORMAL,
        );
        if com.is_ok() {
            CoUninitialize();
        }
        result.0 as isize
    };

    // ShellExecute 는 32 보다 큰 값을 돌려주면 성공이다.
    if code > 32 {
        return Ok(());
    }
    Err(match code {
        2 | 3 => format!("대상을 찾을 수 없습니다: {target}"),
        5 => "접근이 거부되었거나 실행이 취소되었습니다.".into(),
        31 => "이 파일을 열 프로그램이 지정되어 있지 않습니다.".into(),
        other => format!("열지 못했습니다 (오류 {other}): {target}"),
    })
}

#[cfg(not(windows))]
pub fn open(_target: &str, _elevated: bool) -> Result<(), String> {
    Err("바로가기 실행은 Windows 에서만 지원합니다.".into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn urls_and_shell_uris_are_links() {
        for target in ["https://example.com", "mailto:a@b.c", "ms-settings:display", "shell:startup"] {
            assert!(is_link(target), "{target}");
        }
    }

    #[test]
    fn paths_are_not_links() {
        for target in [r"C:\Users\me", "C:/Users/me", r"\\server\share", "notes.txt", "", "https:"] {
            assert!(!is_link(target), "{target}");
        }
    }

    #[test]
    fn existing_paths_are_classified_by_what_is_on_disk() {
        let dir = std::env::temp_dir().join(format!("deskboard-classify-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(dir.join("v1.2")).unwrap();
        std::fs::write(dir.join("tool.exe"), b"").unwrap();
        std::fs::write(dir.join("README"), b"").unwrap();
        std::fs::write(dir.join("memo.txt"), b"").unwrap();

        let kind = |name: &str| classify(&dir.join(name).to_string_lossy());
        // 이름에 점이 있어도 실제로 폴더면 폴더
        assert_eq!(kind("v1.2"), TargetInfo { kind: "folder", exists: true });
        assert_eq!(kind("tool.exe"), TargetInfo { kind: "program", exists: true });
        // 확장자가 없어도 실제로 파일이면 파일
        assert_eq!(kind("README"), TargetInfo { kind: "file", exists: true });
        assert_eq!(kind("memo.txt"), TargetInfo { kind: "file", exists: true });
    }

    #[test]
    fn missing_paths_are_guessed_from_the_name_and_flagged() {
        assert_eq!(classify(r"Z:\nope\app.EXE"), TargetInfo { kind: "program", exists: false });
        assert_eq!(classify("/definitely/not/here"), TargetInfo { kind: "folder", exists: false });
        assert_eq!(classify("/definitely/not/here.pdf"), TargetInfo { kind: "file", exists: false });
    }

    #[test]
    fn links_are_not_checked_on_disk() {
        assert_eq!(classify("  https://example.com  "), TargetInfo { kind: "link", exists: true });
    }
}
