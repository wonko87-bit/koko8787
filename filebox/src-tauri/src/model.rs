use serde::{Deserialize, Serialize};
use std::path::PathBuf;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "snake_case")]
pub enum EntryStatus {
    /// 관리함 안에 있고 아직 최종 목적지가 정해지지 않은 상태
    Inbox,
    /// 즐겨찾기 폴더 등 최종 목적지로 이동 완료
    Filed,
    /// OS 휴지통으로 보냄 (복구는 윈도우 휴지통에서)
    Trashed,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FileEntry {
    pub id: String,
    pub file_name: String,
    /// 현재 실제 경로 (Inbox면 관리함 내부, Filed면 최종 목적지)
    pub path: PathBuf,
    /// 수집 전 원래 있던 경로
    pub origin: PathBuf,
    pub size: u64,
    pub added_at: u64,
    pub category: String,
    pub tags: Vec<String>,
    pub status: EntryStatus,
    pub filed_to: Option<PathBuf>,
    pub filed_at: Option<u64>,
    /// 이동 시 남긴 학습 기록의 id. 되돌리기 때 그 기록만 정확히 제거한다.
    #[serde(default)]
    pub record_id: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Favorite {
    pub id: String,
    pub name: String,
    pub path: PathBuf,
}

/// 사용자 정의 분류/추천 규칙
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Rule {
    pub id: String,
    pub name: String,
    /// 소문자 확장자 목록 (점 없이, 예: "pdf")
    pub extensions: Vec<String>,
    /// 파일명에 포함되면 매칭되는 키워드 (소문자 비교)
    pub keywords: Vec<String>,
    /// 매칭 시 부여할 가상 카테고리
    pub category: Option<String>,
    /// 매칭 시 우선 추천할 즐겨찾기
    pub favorite_id: Option<String>,
}

/// 학습용: 사용자가 파일을 즐겨찾기로 보낸 기록
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MoveRecord {
    #[serde(default)]
    pub id: String,
    pub ext: String,
    pub tokens: Vec<String>,
    pub favorite_id: String,
    pub at: u64,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Settings {
    /// 감시할 폴더 목록 (기본: 다운로드 폴더)
    pub watch_dirs: Vec<PathBuf>,
    /// 관리함 폴더 (수집된 파일이 실제로 놓이는 곳)
    pub inbox_dir: PathBuf,
    /// 새 파일을 자동으로 관리함으로 이동할지
    pub auto_collect: bool,
    /// 새 파일 도착 시 토스트 표시 여부
    pub toast_enabled: bool,
    /// 감시 일시정지
    pub paused: bool,
    /// 앱을 켤 때 새 버전을 자동으로 확인할지. 꺼도 설정 탭에서 수동 확인은 가능하다.
    #[serde(default = "default_true")]
    pub auto_update_check: bool,
}

fn default_true() -> bool {
    true
}

impl Settings {
    pub fn new(watch_dir: PathBuf, inbox_dir: PathBuf) -> Self {
        Self {
            watch_dirs: vec![watch_dir],
            inbox_dir,
            auto_collect: true,
            toast_enabled: true,
            paused: false,
            auto_update_check: true,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct StoreData {
    pub entries: Vec<FileEntry>,
    pub favorites: Vec<Favorite>,
    pub rules: Vec<Rule>,
    pub records: Vec<MoveRecord>,
    pub settings: Option<Settings>,
}

#[derive(Debug, Clone, Serialize)]
pub struct Suggestion {
    pub favorite: Favorite,
    pub score: i64,
}

pub fn now_millis() -> u64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_millis() as u64)
        .unwrap_or(0)
}

pub fn new_id() -> String {
    uuid::Uuid::new_v4().to_string()
}

/// 파일명(확장자 제외)을 학습/매칭용 토큰으로 분해
pub fn tokenize(file_stem: &str) -> Vec<String> {
    file_stem
        .split(|c: char| !c.is_alphanumeric())
        .filter(|t| t.chars().count() >= 2)
        .map(|t| t.to_lowercase())
        .collect()
}

pub fn ext_of(name: &str) -> String {
    std::path::Path::new(name)
        .extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default()
}

pub fn stem_of(name: &str) -> String {
    std::path::Path::new(name)
        .file_stem()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| name.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tokenize_splits_on_non_alphanumeric() {
        let tokens = tokenize("2026-08_회의록 final(v2)");
        assert!(tokens.contains(&"2026".to_string()));
        assert!(tokens.contains(&"회의록".to_string()));
        assert!(tokens.contains(&"final".to_string()));
        assert!(tokens.contains(&"v2".to_string()));
        // 한 글자 토큰은 제외
        assert!(!tokens.contains(&"0".to_string()));
    }

    #[test]
    fn ext_and_stem() {
        assert_eq!(ext_of("보고서.PDF"), "pdf");
        assert_eq!(stem_of("보고서.PDF"), "보고서");
        assert_eq!(ext_of("noext"), "");
    }
}
