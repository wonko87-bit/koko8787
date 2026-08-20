use crate::model::{ext_of, stem_of, Rule};

/// 확장자 기반 기본 카테고리
pub fn default_category(ext: &str) -> &'static str {
    match ext {
        "pdf" | "doc" | "docx" | "xls" | "xlsx" | "ppt" | "pptx" | "hwp" | "hwpx" | "txt"
        | "md" | "csv" | "odt" => "문서",
        "jpg" | "jpeg" | "png" | "gif" | "bmp" | "webp" | "svg" | "heic" | "tiff" => "이미지",
        "mp4" | "mkv" | "avi" | "mov" | "wmv" | "webm" => "영상",
        "mp3" | "wav" | "flac" | "m4a" | "ogg" | "wma" => "음악",
        "zip" | "7z" | "rar" | "tar" | "gz" | "xz" | "egg" | "alz" => "압축",
        "exe" | "msi" | "msix" | "appx" => "설치파일",
        "vbs" | "py" | "js" | "ts" | "rs" | "c" | "cpp" | "h" | "json" | "xml" | "yaml"
        | "yml" | "bat" | "ps1" | "sh" => "코드/스크립트",
        _ => "기타",
    }
}

fn rule_matches(rule: &Rule, file_name: &str) -> bool {
    let ext = ext_of(file_name);
    let lower_name = file_name.to_lowercase();

    let ext_ok = if rule.extensions.is_empty() {
        true
    } else {
        rule.extensions.iter().any(|e| e == &ext)
    };
    let kw_ok = if rule.keywords.is_empty() {
        true
    } else {
        rule.keywords.iter().any(|k| lower_name.contains(&k.to_lowercase()))
    };
    // 조건이 하나도 없는 규칙은 아무것도 매칭하지 않음
    let has_condition = !rule.extensions.is_empty() || !rule.keywords.is_empty();
    has_condition && ext_ok && kw_ok
}

/// 규칙을 순서대로 평가해 카테고리를 결정. 매칭 규칙이 없으면 확장자 기본 카테고리.
pub fn categorize(rules: &[Rule], file_name: &str) -> String {
    for rule in rules {
        if rule_matches(rule, file_name) {
            if let Some(cat) = &rule.category {
                if !cat.is_empty() {
                    return cat.clone();
                }
            }
        }
    }
    default_category(&ext_of(file_name)).to_string()
}

/// 파일명에 매칭되는 규칙들이 지정한 즐겨찾기 id 목록 (우선순위 순)
pub fn rule_favorites(rules: &[Rule], file_name: &str) -> Vec<String> {
    rules
        .iter()
        .filter(|r| rule_matches(r, file_name))
        .filter_map(|r| r.favorite_id.clone())
        .collect()
}

/// 브라우저가 다운로드 중에 쓰는 임시 파일인지
pub fn is_temp_download(file_name: &str) -> bool {
    let lower = file_name.to_lowercase();
    let ext = ext_of(&lower);
    matches!(ext.as_str(), "crdownload" | "part" | "partial" | "download" | "tmp" | "opdownload")
        || lower.starts_with("~$")
        || lower.starts_with(".")
        || stem_of(&lower).is_empty()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rule(exts: &[&str], kws: &[&str], cat: Option<&str>, fav: Option<&str>) -> Rule {
        Rule {
            id: "r1".into(),
            name: "테스트".into(),
            extensions: exts.iter().map(|s| s.to_string()).collect(),
            keywords: kws.iter().map(|s| s.to_string()).collect(),
            category: cat.map(|s| s.to_string()),
            favorite_id: fav.map(|s| s.to_string()),
        }
    }

    #[test]
    fn default_categories() {
        assert_eq!(default_category("pdf"), "문서");
        assert_eq!(default_category("hwp"), "문서");
        assert_eq!(default_category("png"), "이미지");
        assert_eq!(default_category("xyz"), "기타");
    }

    #[test]
    fn rule_overrides_default() {
        let rules = vec![rule(&["pdf"], &["청구서"], Some("경리"), None)];
        assert_eq!(categorize(&rules, "8월_청구서.pdf"), "경리");
        assert_eq!(categorize(&rules, "여행사진.pdf"), "문서"); // 키워드 불일치 → 기본
    }

    #[test]
    fn keyword_only_rule() {
        let rules = vec![rule(&[], &["simcenter"], Some("시뮬레이션"), Some("f1"))];
        assert_eq!(categorize(&rules, "Simcenter_model.zip"), "시뮬레이션");
        assert_eq!(rule_favorites(&rules, "simcenter_run.log"), vec!["f1".to_string()]);
    }

    #[test]
    fn empty_rule_matches_nothing() {
        let rules = vec![rule(&[], &[], Some("전부"), None)];
        assert_eq!(categorize(&rules, "anything.bin"), "기타");
    }

    #[test]
    fn temp_downloads_detected() {
        assert!(is_temp_download("movie.mp4.crdownload"));
        assert!(is_temp_download("setup.exe.part"));
        assert!(is_temp_download("~$report.docx"));
        assert!(!is_temp_download("report.docx"));
    }
}
