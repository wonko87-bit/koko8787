//! 탐색기가 보여 주는 것과 같은 아이콘을 구해 PNG 로 넘겨준다.

/// 한 변이 `size` 픽셀인 RGBA 그림
pub struct Image {
    pub size: u32,
    pub rgba: Vec<u8>,
}

/// 타일에 쓸 아이콘 크기. 가장 큰 타일(50px)을 배율 200% 화면에서도 또렷하게 그릴 수 있는 크기다.
pub const ICON_SIZE: u32 = 96;

/// `data:image/png;base64,…` 꼴로 만든다.
pub fn to_data_url(image: &Image) -> Option<String> {
    use base64::Engine;
    let mut bytes = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut bytes, image.size, image.size);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        let mut writer = encoder.write_header().ok()?;
        writer.write_image_data(&image.rgba).ok()?;
    }
    Some(format!("data:image/png;base64,{}", base64::engine::general_purpose::STANDARD.encode(bytes)))
}

/// 셸이 주는 그림은 BGRA 순서에 색이 투명도만큼 미리 곱해져 있다. PNG 가 쓰는 꼴로 되돌린다.
pub fn bgra_premultiplied_to_rgba(pixels: &mut [u8]) {
    // 투명도 정보가 아예 없는(전부 0인) 오래된 아이콘은 불투명한 그림으로 본다.
    let opaque = pixels.chunks_exact(4).all(|p| p[3] == 0);
    for p in pixels.chunks_exact_mut(4) {
        p.swap(0, 2);
        let alpha = if opaque { 255 } else { p[3] };
        if alpha != 0 && alpha != 255 {
            for channel in &mut p[..3] {
                *channel = ((*channel as u32 * 255 + alpha as u32 / 2) / alpha as u32).min(255) as u8;
            }
        }
        p[3] = alpha;
    }
}

/// 경로가 가리키는 것의 아이콘. `thumbnail` 이 참이면 그림 파일은 내용 미리보기를 준다.
#[cfg(windows)]
pub fn load(path: &str, thumbnail: bool) -> Option<Image> {
    use std::os::windows::ffi::OsStrExt;
    use windows::core::PCWSTR;
    use windows::Win32::Foundation::SIZE;
    use windows::Win32::Graphics::Gdi::{
        CreateCompatibleDC, DeleteDC, DeleteObject, GetDIBits, BITMAPINFO, BITMAPINFOHEADER,
        BI_RGB, DIB_RGB_COLORS,
    };
    use windows::Win32::System::Com::{
        CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE,
    };
    use windows::Win32::UI::Shell::{
        IShellItemImageFactory, SHCreateItemFromParsingName, SIIGBF_BIGGERSIZEOK, SIIGBF_ICONONLY,
        SIIGBF_RESIZETOFIT,
    };

    let wide: Vec<u16> = std::ffi::OsStr::new(path.trim()).encode_wide().chain(Some(0)).collect();
    let size = ICON_SIZE as i32;

    unsafe {
        let com = CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
        let image = (|| {
            let factory: IShellItemImageFactory =
                SHCreateItemFromParsingName(PCWSTR(wide.as_ptr()), None).ok()?;
            let flags = if thumbnail { SIIGBF_RESIZETOFIT } else { SIIGBF_ICONONLY | SIIGBF_BIGGERSIZEOK };
            let bitmap = factory.GetImage(SIZE { cx: size, cy: size }, flags).ok()?;

            // 받은 그림을 위에서 아래로 내려가는 32비트 픽셀로 읽는다.
            let mut info = BITMAPINFO {
                bmiHeader: BITMAPINFOHEADER {
                    biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                    biWidth: size,
                    biHeight: -size,
                    biPlanes: 1,
                    biBitCount: 32,
                    biCompression: BI_RGB.0,
                    ..Default::default()
                },
                ..Default::default()
            };
            let mut pixels = vec![0u8; (size * size * 4) as usize];
            let dc = CreateCompatibleDC(None);
            let lines = GetDIBits(
                dc,
                bitmap,
                0,
                size as u32,
                Some(pixels.as_mut_ptr().cast()),
                &mut info,
                DIB_RGB_COLORS,
            );
            let _ = DeleteDC(dc);
            let _ = DeleteObject(bitmap.into());
            if lines == 0 {
                return None;
            }
            bgra_premultiplied_to_rgba(&mut pixels);
            Some(Image { size: ICON_SIZE, rgba: pixels })
        })();
        if com.is_ok() {
            CoUninitialize();
        }
        image
    }
}

#[cfg(not(windows))]
pub fn load(_path: &str, _thumbnail: bool) -> Option<Image> {
    None
}

/// 설치된 글꼴 이름을 가나다순으로 준다.
#[cfg(windows)]
pub fn font_families() -> Vec<String> {
    use windows::Win32::Foundation::LPARAM;
    use windows::Win32::Graphics::Gdi::{
        CreateCompatibleDC, DeleteDC, EnumFontFamiliesExW, DEFAULT_CHARSET, LOGFONTW, TEXTMETRICW,
    };

    unsafe extern "system" fn collect(font: *const LOGFONTW, _: *const TEXTMETRICW, _: u32, out: LPARAM) -> i32 {
        let names = &mut *(out.0 as *mut Vec<String>);
        let face = &(*font).lfFaceName;
        let len = face.iter().position(|&c| c == 0).unwrap_or(face.len());
        names.push(String::from_utf16_lossy(&face[..len]));
        1
    }

    let mut names: Vec<String> = Vec::new();
    unsafe {
        let dc = CreateCompatibleDC(None);
        let filter = LOGFONTW { lfCharSet: DEFAULT_CHARSET, ..Default::default() };
        EnumFontFamiliesExW(dc, &filter, Some(collect), LPARAM(&mut names as *mut _ as isize), 0);
        let _ = DeleteDC(dc);
    }
    tidy_font_names(names)
}

#[cfg(not(windows))]
pub fn font_families() -> Vec<String> {
    Vec::new()
}

/// 세로쓰기용 변형(@로 시작)과 중복을 빼고 정렬한다.
pub fn tidy_font_names(mut names: Vec<String>) -> Vec<String> {
    names.retain(|name| !name.is_empty() && !name.starts_with('@'));
    names.sort_by_key(|name| name.to_lowercase());
    names.dedup();
    names
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn premultiplied_bgra_becomes_straight_rgba() {
        // 반투명한 순수 파랑 (B=128 은 255 에 alpha 128 을 곱한 값), 완전 불투명 빨강, 완전 투명
        let mut pixels = vec![128, 0, 0, 128, 0, 0, 255, 255, 0, 0, 0, 0];
        bgra_premultiplied_to_rgba(&mut pixels);
        assert_eq!(pixels, vec![0, 0, 255, 128, 255, 0, 0, 255, 0, 0, 0, 0]);
    }

    #[test]
    fn image_without_any_alpha_is_treated_as_opaque() {
        let mut pixels = vec![10, 20, 30, 0, 40, 50, 60, 0];
        bgra_premultiplied_to_rgba(&mut pixels);
        assert_eq!(pixels, vec![30, 20, 10, 255, 60, 50, 40, 255]);
    }

    #[test]
    fn data_url_holds_a_png_of_the_right_size() {
        use base64::Engine;
        let image = Image { size: 2, rgba: vec![255; 16] };
        let url = to_data_url(&image).unwrap();
        let encoded = url.strip_prefix("data:image/png;base64,").unwrap();
        let bytes = base64::engine::general_purpose::STANDARD.decode(encoded).unwrap();
        let reader = png::Decoder::new(bytes.as_slice()).read_info().unwrap();
        assert_eq!((reader.info().width, reader.info().height), (2, 2));
    }

    #[test]
    fn font_names_are_deduplicated_sorted_and_without_vertical_variants() {
        let names = vec!["굴림".to_string(), "@굴림".into(), "Arial".into(), "arial black".into(), "Arial".into(), "".into()];
        assert_eq!(tidy_font_names(names), vec!["Arial", "arial black", "굴림"]);
    }

    #[cfg(windows)]
    #[test]
    fn real_programs_and_folders_have_icons() {
        let windows = std::env::var("SystemRoot").unwrap();
        for path in [format!(r"{windows}\System32\cmd.exe"), windows.clone()] {
            let image = load(&path, false).unwrap_or_else(|| panic!("no icon for {path}"));
            assert_eq!(image.rgba.len(), (ICON_SIZE * ICON_SIZE * 4) as usize);
            let visible = image.rgba.chunks_exact(4).filter(|p| p[3] > 0).count();
            assert!(visible > 500, "{path}: only {visible} visible pixels");
            assert!(to_data_url(&image).is_some());
        }
    }

    #[cfg(windows)]
    #[test]
    fn missing_path_has_no_icon() {
        assert!(load(r"Z:\definitely\not\here.exe", false).is_none());
    }

    #[cfg(windows)]
    #[test]
    fn common_fonts_are_listed() {
        let fonts = font_families();
        assert!(fonts.iter().any(|f| f == "Segoe UI"), "{} fonts, no Segoe UI", fonts.len());
        assert!(fonts.iter().all(|f| !f.starts_with('@')));
    }
}
