//! Flowdeck 으로 할일을 넘기는 쪽.
//!
//! 규격 전체는 저장소의 `docs/flowdeck-bridge.md` 에 있다. 여기서 지켜야 하는 것만 추리면:
//! 파일은 Flowdeck 이 이미 쓰는 `TransferArchive` 모양이어야 하고, 키는 PascalCase 여야
//! 하고(대소문자가 틀리면 그 값은 조용히 사라진다), `.tmp` 로 쓴 뒤 `.txt` 로 이름을
//! 바꿔야 한다. 마지막 것이 특히 중요한데, Flowdeck 의 감시자는 `.txt` 가 보이는 순간
//! 그 파일이 완성됐다고 믿기 때문이다.

use crate::model::{ext_of, now_millis, stem_of, FlowdeckSpec, Rule};
use chrono::{DateTime, Datelike, Duration, Local, NaiveDateTime};
use serde::Serialize;
use std::path::{Path, PathBuf};

/// 사람이 읽는 부분과 앱이 읽는 부분의 경계. Flowdeck 이 이 문자열을 그대로 찾으므로
/// 한 글자도 바꾸면 안 된다.
pub(crate) const MARKER: &str = "--- 여기서부터는 앱이 읽는 부분입니다. 지우지 마세요 ---";
const FORMAT: &str = "flowdeck.transfer";
const VERSION: u32 = 1;

/// 같은 파일을 두 번 보내도 할일이 두 개가 되지 않게 하는 장치.
///
/// Flowdeck 의 가져오기는 `TodoItem.Id` 가 이미 있으면 건너뛴다. 그래서 id 만 항상
/// 같으면 재발송이 저절로 무해해진다. 발송 성공 여부를 어딘가에 적어 두고 그걸 믿는
/// 대신, (항목, 규칙) 에서 id 를 계산해 버리면 기록이 날아가도 성질이 유지된다.
const NAMESPACE: uuid::Uuid = uuid::uuid!("6f3a1d92-0c47-4b8e-9a15-2d7c8e4f0b63");

/// 이 항목을 이 규칙으로 보냈을 때의 할일 id. 몇 번을 계산해도 같은 값이 나온다.
pub fn todo_id(entry_id: &str, rule_id: &str) -> String {
    uuid::Uuid::new_v5(&NAMESPACE, format!("{entry_id}:{rule_id}").as_bytes())
        .simple()
        .to_string()
}

/// Flowdeck 이 기본으로 감시하는 폴더. 설정에서 덮어쓰지 않았다면 여기로 보낸다.
pub fn default_inbox() -> Option<PathBuf> {
    #[cfg(target_os = "windows")]
    let base = std::env::var_os("APPDATA").map(PathBuf::from);

    // 개발과 테스트는 윈도우가 아닌 곳에서도 돌아가야 한다. Flowdeck 자체는 윈도우
    // 전용이므로 이 경로로 실제 전달이 일어나는 일은 없고, 경로 계산만 성립시킨다.
    #[cfg(not(target_os = "windows"))]
    let base = std::env::var_os("HOME").map(|h| PathBuf::from(h).join(".config"));

    base.map(|b| b.join("Flowdeck").join("inbox"))
}

// ---------- 전송 파일 만들기 ----------

/// Flowdeck 의 `TransferArchive`. 필드 이름이 곧 규격이다.
#[derive(Serialize)]
#[serde(rename_all = "PascalCase")]
struct Archive {
    format: &'static str,
    version: u32,
    exported_at: String,
    todos: Vec<Todo>,
    /// 브릿지는 일정을 만들지 않는다. 그래도 키 자체는 있어야 한다.
    events: Vec<serde_json::Value>,
}

/// Flowdeck 의 `TodoItem` 중 브릿지가 채우는 것들.
///
/// 여기 없는 필드(Recurrence, ExternalLink, LinkedEventId)는 일부러 뺐다. C# 쪽 속성
/// 초기화가 각각 `Recurrence.None` 과 `null` 을 넣어 주므로, 보내지 않는 편이 낫다.
#[derive(Serialize)]
#[serde(rename_all = "PascalCase")]
pub struct Todo {
    id: String,
    title: String,
    notes: String,
    /// 시간대 표기를 붙이지 않는다. 붙이면 .NET 이 그걸 시간대 정보로 읽고 값을
    /// 옮겨 버리는데, 여기 담긴 건 시간대가 아니라 "그날 몇 시"라는 벽시계 값이다.
    #[serde(skip_serializing_if = "Option::is_none")]
    due_at: Option<String>,
    has_time: bool,
    priority: String,
    tags: Vec<String>,
    is_done: bool,
    created_at: String,
    updated_at: String,
    source_text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    reminder_minutes_before: Option<i64>,
}

fn stamp_naive(dt: &NaiveDateTime) -> String {
    dt.format("%Y-%m-%dT%H:%M:%S").to_string()
}

/// "HH:MM" 을 시/분으로. 형식이 아니면 None 이라 날짜만 지정한 것과 같아진다.
fn parse_hhmm(text: &str) -> Option<(u32, u32)> {
    let (h, m) = text.trim().split_once(':')?;
    let (h, m) = (h.trim().parse::<u32>().ok()?, m.trim().parse::<u32>().ok()?);
    (h < 24 && m < 60).then_some((h, m))
}

/// 기한과 "시각까지 정해졌는지" 여부.
///
/// 시각을 안 정했으면 자정으로 둔다. Flowdeck 이 자기 입력창에서 날짜만 받았을 때
/// 하는 것과 같아서, 브릿지로 만든 할일이 손으로 적은 할일과 다르게 굴지 않는다.
fn due_of(spec: &FlowdeckSpec, now: &DateTime<Local>) -> (Option<String>, bool) {
    let Some(days) = spec.due_in_days else {
        return (None, false);
    };
    let date = (now.date_naive()) + Duration::days(days);
    match spec.due_time.as_deref().and_then(parse_hhmm) {
        Some((h, m)) => (date.and_hms_opt(h, m, 0).as_ref().map(stamp_naive), true),
        None => (date.and_hms_opt(0, 0, 0).as_ref().map(stamp_naive), false),
    }
}

/// Flowdeck 의 Priority 로 받아들여지는 이름만 통과시킨다. 모르는 값은 None.
fn priority_of(raw: &str) -> String {
    match raw.trim() {
        "Low" => "Low",
        "Normal" => "Normal",
        "High" => "High",
        "Urgent" => "Urgent",
        _ => "None",
    }
    .to_string()
}

/// 제목 템플릿 치환.
fn render_title(template: &str, file_name: &str, category: &str) -> String {
    let title = template
        .replace("{파일명}", &stem_of(file_name))
        .replace("{확장자}", &ext_of(file_name))
        .replace("{카테고리}", category);
    let trimmed = title.trim();
    // 제목 없는 할일은 목록에서 빈 줄로 보인다. 그럴 바엔 파일명이 낫다.
    if trimmed.is_empty() {
        file_name.to_string()
    } else {
        trimmed.to_string()
    }
}

/// 할일 하나를 조립한다.
///
/// `path` 는 반드시 절대 경로여야 한다. Flowdeck 의 가져오기가 `ExternalLink` 를
/// 무조건 지우기 때문에, 나중에 원본 파일로 돌아갈 실마리는 Notes 본문뿐이다.
pub fn build_todo(
    spec: &FlowdeckSpec,
    id: String,
    file_name: &str,
    path: &Path,
    category: &str,
    rule_name: &str,
    now: &DateTime<Local>,
) -> Todo {
    let (due_at, has_time) = due_of(spec, now);
    let stamp = stamp_naive(&now.naive_local());
    let source = if rule_name.is_empty() {
        "출처: FileBox 에서 직접 보냄".to_string()
    } else {
        format!("출처: FileBox 특별규칙 \"{rule_name}\"")
    };

    Todo {
        id,
        title: render_title(&spec.title, file_name, category),
        notes: format!("파일: {}\n{source}", path.display()),
        due_at,
        has_time,
        priority: priority_of(&spec.priority),
        tags: spec.tags.iter().map(|t| t.trim().to_string()).filter(|t| !t.is_empty()).collect(),
        is_done: false,
        created_at: stamp.clone(),
        updated_at: stamp,
        source_text: file_name.to_string(),
        reminder_minutes_before: spec.reminder_minutes.filter(|m| *m > 0),
    }
}

/// 파일 전체 내용. 마커 위쪽은 사람이 폴더를 열어 봤을 때 읽으라고 있는 것이고,
/// Flowdeck 은 마커 뒤의 JSON 만 본다.
pub fn build_file(todos: Vec<Todo>, now: &DateTime<Local>) -> String {
    let mut text = format!(
        "FileBox → Flowdeck · {} · {}건\n\n",
        now.format("%Y-%m-%d %H:%M"),
        todos.len()
    );
    for todo in &todos {
        text.push_str(&format!("- [할일] {}", todo.title));
        if let Some(due) = &todo.due_at {
            if let Some(when) = NaiveDateTime::parse_from_str(due, "%Y-%m-%dT%H:%M:%S").ok() {
                text.push_str(&format!(" · {}월 {}일", when.month(), when.day()));
                if todo.has_time {
                    text.push_str(&format!(" {}", when.format("%H:%M")));
                }
            }
        }
        for tag in &todo.tags {
            text.push_str(&format!(" #{tag}"));
        }
        text.push('\n');
    }

    let archive = Archive {
        format: FORMAT,
        version: VERSION,
        exported_at: now.format("%Y-%m-%dT%H:%M:%S%:z").to_string(),
        todos,
        events: Vec::new(),
    };

    text.push('\n');
    text.push_str(MARKER);
    text.push('\n');
    text.push_str(&serde_json::to_string_pretty(&archive).unwrap_or_default());
    text
}

/// `.tmp` 로 쓰고 `.txt` 로 이름을 바꾼다.
///
/// 같은 폴더 안의 이름 바꾸기는 원자적이라, `.txt` 가 존재하는 순간 그 파일은 이미
/// 완성돼 있다. 곧바로 `.txt` 로 쓰면 Flowdeck 의 감시자가 아직 0바이트인 파일을
/// 열게 된다.
pub fn write_transfer(inbox: &Path, contents: &str, suffix: &str) -> std::io::Result<PathBuf> {
    std::fs::create_dir_all(inbox)?;
    let base = format!(
        "filebox-{}-{}",
        Local::now().format("%Y%m%d-%H%M%S"),
        &suffix[..suffix.len().min(8)]
    );
    let tmp = inbox.join(format!("{base}.tmp"));
    let txt = inbox.join(format!("{base}.txt"));
    std::fs::write(&tmp, contents)?;
    std::fs::rename(&tmp, &txt)?;
    Ok(txt)
}

/// 파일 하나를 규칙 하나에 따라 Flowdeck 으로 보낸다. 성공하면 남길 흔적을 돌려준다.
#[allow(clippy::too_many_arguments)]
pub fn send(
    inbox: &Path,
    spec: &FlowdeckSpec,
    entry_id: &str,
    rule_id: &str,
    rule_name: &str,
    file_name: &str,
    path: &Path,
    category: &str,
) -> std::io::Result<crate::model::FlowdeckTodo> {
    let now = Local::now();
    let id = todo_id(entry_id, rule_id);
    let todo = build_todo(spec, id.clone(), file_name, path, category, rule_name, &now);
    write_transfer(inbox, &build_file(vec![todo], &now), &id)?;
    Ok(crate::model::FlowdeckTodo {
        rule_id: rule_id.to_string(),
        todo_id: id,
        at: now_millis(),
    })
}

/// 이 파일명에 걸리는 규칙 중 특별규칙이 붙은 것들.
pub fn matching_specs<'a>(rules: &'a [Rule], file_name: &str) -> Vec<&'a Rule> {
    rules
        .iter()
        .filter(|r| r.flowdeck.is_some() && crate::rules::rule_matches(r, file_name))
        .collect()
}

/// 지금 쓰는 감시 폴더. 설정에 적힌 것이 우선이고, 없으면 기본 경로.
pub fn inbox_of(store: &crate::store::Store) -> Option<PathBuf> {
    let (enabled, configured) = store.read(|d| {
        d.settings
            .as_ref()
            .map(|s| (s.flowdeck_enabled, s.flowdeck_inbox.clone()))
            .unwrap_or((false, None))
    });
    if !enabled {
        return None;
    }
    configured.filter(|p| !p.as_os_str().is_empty()).or_else(default_inbox)
}

/// 수집된 파일 하나를 걸리는 특별규칙마다 Flowdeck 으로 보낸다.
///
/// 규칙 하나가 실패해도 나머지는 계속 보낸다. 여기서 할 수 있는 복구가 없어서
/// 실패는 로그로만 남기는데, id 가 결정적이라 사용자가 나중에 손으로 다시 보내면
/// 같은 할일이 한 번만 생긴다.
pub fn dispatch(
    store: &crate::store::Store,
    entry: &crate::model::FileEntry,
) -> Vec<crate::model::FlowdeckTodo> {
    let Some(inbox) = inbox_of(store) else {
        return Vec::new();
    };
    let rules = store.read(|d| d.rules.clone());

    let matched = matching_specs(&rules, &entry.file_name);

    let mut sent = Vec::new();
    for rule in &matched {
        let Some(spec) = &rule.flowdeck else { continue };
        match send(
            &inbox,
            spec,
            &entry.id,
            &rule.id,
            &rule.name,
            &entry.file_name,
            &entry.path,
            &entry.category,
        ) {
            Ok(todo) => sent.push(todo),
            Err(e) => eprintln!("Flowdeck 전달 실패 ({}): {e}", entry.file_name),
        }
    }

    // 걸리는 규칙이 없는데 사용자가 직접 표시해 둔 파일. 규칙이 이미 처리했다면
    // 예약분까지 또 보내지 않는다 — 같은 파일로 할일이 두 개 생긴다.
    if entry.flowdeck_pending && sent.is_empty() {
        let spec = crate::model::FlowdeckSpec {
            due_in_days: None,
            ..Default::default()
        };
        match send(
            &inbox,
            &spec,
            &entry.id,
            "",
            "",
            &entry.file_name,
            &entry.path,
            &entry.category,
        ) {
            Ok(todo) => sent.push(todo),
            Err(e) => eprintln!("Flowdeck 전달 실패 ({}): {e}", entry.file_name),
        }
    }

    sent
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::TimeZone;
    use serde_json::Value;

    fn at(y: i32, mo: u32, d: u32, h: u32, mi: u32) -> DateTime<Local> {
        Local.with_ymd_and_hms(y, mo, d, h, mi, 0).unwrap()
    }

    fn spec() -> FlowdeckSpec {
        FlowdeckSpec {
            title: "[읽기] {파일명}".into(),
            due_in_days: Some(7),
            due_time: Some("10:00".into()),
            priority: "High".into(),
            tags: vec!["리포트".into()],
            reminder_minutes: Some(30),
        }
    }

    fn json_of(text: &str) -> Value {
        let start = text.find(MARKER).expect("마커가 있어야 한다") + MARKER.len();
        serde_json::from_str(&text[start..]).expect("마커 뒤는 JSON 이어야 한다")
    }

    #[test]
    fn keys_are_pascal_case() {
        // 대소문자가 틀리면 C# 쪽에서 값이 조용히 사라진다. 이 테스트가 그 방어선이다.
        let now = at(2026, 8, 27, 14, 32);
        let todo = build_todo(&spec(), "abc".into(), "보고서.pdf", Path::new("/tmp/보고서.pdf"), "문서", "리포트 읽기", &now);
        let v = json_of(&build_file(vec![todo], &now));

        assert_eq!(v["Format"], "flowdeck.transfer");
        assert_eq!(v["Version"], 1);
        let t = &v["Todos"][0];
        for key in ["Id", "Title", "Notes", "DueAt", "HasTime", "Priority", "Tags", "IsDone", "CreatedAt", "UpdatedAt", "SourceText", "ReminderMinutesBefore"] {
            assert!(!t[key].is_null(), "{key} 가 비었다");
        }
        assert!(v["Events"].as_array().unwrap().is_empty());
    }

    #[test]
    fn due_date_carries_no_timezone() {
        // 오프셋이나 Z 가 붙으면 .NET 이 값을 옮긴다.
        let now = at(2026, 8, 27, 14, 32);
        let todo = build_todo(&spec(), "abc".into(), "a.pdf", Path::new("/tmp/a.pdf"), "문서", "r", &now);
        let v = json_of(&build_file(vec![todo], &now));
        assert_eq!(v["Todos"][0]["DueAt"], "2026-09-03T10:00:00");
        assert_eq!(v["Todos"][0]["HasTime"], true);
        // 반대로 ExportedAt 은 DateTimeOffset 이라 오프셋이 있어야 한다.
        assert!(v["ExportedAt"].as_str().unwrap().len() > 19);
    }

    #[test]
    fn date_only_due_sits_at_midnight() {
        // Flowdeck 이 날짜만 입력받았을 때와 같은 모양이어야 한다.
        let mut s = spec();
        s.due_time = None;
        let now = at(2026, 8, 27, 14, 32);
        let todo = build_todo(&s, "abc".into(), "a.pdf", Path::new("/tmp/a.pdf"), "문서", "r", &now);
        let v = json_of(&build_file(vec![todo], &now));
        assert_eq!(v["Todos"][0]["DueAt"], "2026-09-03T00:00:00");
        assert_eq!(v["Todos"][0]["HasTime"], false);
    }

    #[test]
    fn no_due_date_omits_the_field() {
        let mut s = spec();
        s.due_in_days = None;
        let now = at(2026, 8, 27, 14, 32);
        let todo = build_todo(&s, "abc".into(), "a.pdf", Path::new("/tmp/a.pdf"), "문서", "r", &now);
        let v = json_of(&build_file(vec![todo], &now));
        assert!(v["Todos"][0].get("DueAt").is_none());
    }

    #[test]
    fn notes_keep_the_absolute_path() {
        // 가져오기가 ExternalLink 를 지우므로 경로가 살아남을 곳은 여기뿐이다.
        let now = at(2026, 8, 27, 14, 32);
        let todo = build_todo(&spec(), "abc".into(), "a.pdf", Path::new("/home/u/관리함/a.pdf"), "문서", "r", &now);
        let v = json_of(&build_file(vec![todo], &now));
        assert!(v["Todos"][0]["Notes"].as_str().unwrap().contains("/home/u/관리함/a.pdf"));
    }

    #[test]
    fn title_template_is_substituted() {
        let now = at(2026, 8, 27, 14, 32);
        let todo = build_todo(&spec(), "abc".into(), "2026_시장분석.pdf", Path::new("/tmp/x.pdf"), "문서", "r", &now);
        let v = json_of(&build_file(vec![todo], &now));
        assert_eq!(v["Todos"][0]["Title"], "[읽기] 2026_시장분석");
    }

    #[test]
    fn unknown_priority_falls_back_to_none() {
        assert_eq!(priority_of("보통"), "None");
        assert_eq!(priority_of("High"), "High");
    }

    #[test]
    fn todo_id_is_stable_and_distinct() {
        // 같은 (항목, 규칙) 은 몇 번을 계산해도 같은 id → 재발송이 무해해진다.
        assert_eq!(todo_id("e1", "r1"), todo_id("e1", "r1"));
        assert_ne!(todo_id("e1", "r1"), todo_id("e1", "r2"));
        assert_ne!(todo_id("e1", "r1"), todo_id("e2", "r1"));
        assert_eq!(todo_id("e1", "r1").len(), 32);
        assert!(todo_id("e1", "r1").chars().all(|c| c.is_ascii_hexdigit()));
    }

    fn store_with(enabled: bool, custom: Option<&str>) -> crate::store::Store {
        let dir = std::env::temp_dir().join(format!("fbinbox-{}", crate::model::new_id()));
        let store = crate::store::Store::load(dir.join("store.json"));
        store.update(|d| {
            let mut s = crate::model::Settings::new("/dl".into(), "/inbox".into());
            s.flowdeck_enabled = enabled;
            s.flowdeck_inbox = custom.map(Into::into);
            d.settings = Some(s);
        });
        store
    }

    #[test]
    fn inbox_follows_the_setting() {
        let custom = store_with(true, Some("/somewhere/else"));
        assert_eq!(inbox_of(&custom), Some(PathBuf::from("/somewhere/else")));

        // 연동을 끄면 경로 자체가 없다 → 보내는 코드가 전부 조용히 멈춘다.
        assert_eq!(inbox_of(&store_with(false, Some("/somewhere/else"))), None);

        // 비워 둔 경로는 설정하지 않은 것과 같게 다뤄야 한다. 빈 문자열을 그대로
        // 쓰면 현재 디렉터리에 파일을 흘리게 된다.
        assert_eq!(inbox_of(&store_with(true, Some(""))), default_inbox());
        assert_eq!(inbox_of(&store_with(true, None)), default_inbox());
    }

    #[test]
    fn notes_carry_whatever_path_it_is_given() {
        // 0.2.4 는 수집 시점에 보내서 메모에 관리함 경로가 박혔고, 파일을 옮기는
        // 순간 그 경로가 죽었다. 이제 이동이 끝난 뒤에 보내므로 최종 경로가 들어간다.
        // build_todo 는 받은 경로를 그대로 적는다 — 어느 경로를 넘기느냐가 전부다.
        let now = at(2026, 8, 27, 14, 32);
        let filed = Path::new("/home/u/프로젝트/시장분석/a.pdf");
        let todo = build_todo(&spec(), "abc".into(), "a.pdf", filed, "문서", "r", &now);
        let v = json_of(&build_file(vec![todo], &now));
        let notes = v["Todos"][0]["Notes"].as_str().unwrap();
        assert!(notes.contains("/home/u/프로젝트/시장분석/a.pdf"));
        assert!(!notes.contains("관리함"), "관리함 경로가 새어 나갔다: {notes}");
    }

    /// dispatch 를 실제로 돌리고, 감시 폴더에 떨어진 전송 파일들을 읽어 온다.
    fn dispatch_into_temp(
        rules: Vec<Rule>,
        entry: &crate::model::FileEntry,
    ) -> (Vec<crate::model::FlowdeckTodo>, Vec<String>) {
        let dir = std::env::temp_dir().join(format!("fbdisp-{}", crate::model::new_id()));
        let store = store_with(true, Some(dir.to_str().unwrap()));
        store.update(|d| d.rules = rules);

        let sent = dispatch(&store, entry);
        let mut files: Vec<String> = std::fs::read_dir(&dir)
            .map(|rd| {
                rd.filter_map(|e| e.ok())
                    .filter(|e| e.path().extension().is_some_and(|x| x == "txt"))
                    .filter_map(|e| std::fs::read_to_string(e.path()).ok())
                    .collect()
            })
            .unwrap_or_default();
        files.sort();
        std::fs::remove_dir_all(&dir).ok();
        (sent, files)
    }

    fn flowdeck_rule() -> Rule {
        Rule {
            id: "r1".into(),
            name: "리포트".into(),
            extensions: vec!["pdf".into()],
            keywords: vec![],
            category: None,
            favorite_id: None,
            flowdeck: Some(spec()),
        }
    }

    #[test]
    fn nothing_is_sent_without_a_rule_or_a_mark() {
        let (sent, files) = dispatch_into_temp(vec![], &sample_entry());
        assert!(sent.is_empty());
        assert!(files.is_empty(), "보낼 이유가 없는데 파일이 생겼다");
    }

    #[test]
    fn a_marked_entry_goes_even_with_no_rule() {
        // 규칙을 만들 만큼 반복되지 않는 파일. 관리함에서 표시해 두면 이동할 때 나간다.
        let mut entry = sample_entry();
        entry.flowdeck_pending = true;

        let (sent, files) = dispatch_into_temp(vec![], &entry);
        assert_eq!(sent.len(), 1);
        assert_eq!(files.len(), 1);
        // 표시만으로 보낸 것은 기한을 붙이지 않는다.
        assert!(!files[0].contains("\"DueAt\""), "기한 없이 나가야 한다");
    }

    #[test]
    fn a_marked_entry_matching_a_rule_makes_one_todo_not_two() {
        // 규칙이 이미 보냈는데 표시분까지 또 보내면 같은 파일로 할일이 두 개가 된다.
        let mut entry = sample_entry();
        entry.flowdeck_pending = true;

        let (sent, files) = dispatch_into_temp(vec![flowdeck_rule()], &entry);
        assert_eq!(sent.len(), 1, "할일이 두 개 만들어졌다");
        assert_eq!(files.len(), 1);
        assert!(files[0].contains("\"DueAt\""), "규칙의 기한이 쓰여야 한다");
    }

    #[test]
    fn the_filed_path_is_what_reaches_flowdeck() {
        // 이 릴리스의 요점. 메모에 최종 폴더가 적혀야 하고 관리함이 아니어야 한다.
        let mut entry = sample_entry();
        entry.path = "/home/u/프로젝트/시장분석/a.pdf".into();

        let (_, files) = dispatch_into_temp(vec![flowdeck_rule()], &entry);
        assert_eq!(files.len(), 1);
        // 원문이 아니라 파싱된 값을 본다 — windows_paths_survive_json_escaping 참고.
        let notes = json_of(&files[0])["Todos"][0]["Notes"].as_str().unwrap().to_string();
        assert!(notes.contains("/home/u/프로젝트/시장분석/a.pdf"), "{notes}");
    }

    fn sample_entry() -> crate::model::FileEntry {
        crate::model::FileEntry {
            id: "e1".into(),
            file_name: "a.pdf".into(),
            path: "/dest/a.pdf".into(),
            origin: "/dl/a.pdf".into(),
            size: 1,
            added_at: 0,
            category: "문서".into(),
            tags: vec![],
            status: crate::model::EntryStatus::Filed,
            filed_to: Some("/dest".into()),
            filed_at: Some(1),
            record_id: None,
            flowdeck_todos: vec![],
            flowdeck_pending: false,
            recent_cleared: false,
            pinned: false,
        }
    }

    #[test]
    fn windows_paths_survive_json_escaping() {
        // 역슬래시는 JSON 에서 \\ 로 이스케이프된다. 그래서 파일 원문을 문자열
        // 포함으로 검사하면 윈도우에서만 어긋난다 — 정작 Flowdeck 은 파싱해서
        // 원래 경로를 되찾으므로 아무 문제가 없는데도 그렇다.
        let now = at(2026, 8, 27, 14, 32);
        let win = Path::new(r"D:\작업\최종폴더\2026_상반기_시장분석.pdf");
        let todo = build_todo(&spec(), "abc".into(), "a.pdf", win, "문서", "r", &now);
        let text = build_file(vec![todo], &now);

        // 원문에는 이스케이프된 형태로만 들어 있다.
        assert!(!text.contains(r"D:\작업\최종폴더"), "원문 매칭은 성립하지 않아야 한다");
        // 파싱하면 정확히 원래 경로가 나온다. Flowdeck 이 보는 것이 이쪽이다.
        let notes = json_of(&text)["Todos"][0]["Notes"].as_str().unwrap().to_string();
        assert!(
            notes.contains(&win.display().to_string()),
            "파싱된 메모에 경로가 없다: {notes}"
        );
    }

    #[test]
    fn writes_tmp_then_renames_to_txt() {
        let dir = std::env::temp_dir().join(format!("fbtest-{}", crate::model::new_id()));
        let path = write_transfer(&dir, "payload", "abcdef123456").unwrap();
        assert_eq!(path.extension().unwrap(), "txt");
        assert_eq!(std::fs::read_to_string(&path).unwrap(), "payload");
        // .tmp 가 남아 있으면 Flowdeck 이 무시하긴 하지만 폴더가 지저분해진다.
        let leftovers: Vec<_> = std::fs::read_dir(&dir)
            .unwrap()
            .filter_map(|e| e.ok())
            .filter(|e| e.path().extension().is_some_and(|x| x == "tmp"))
            .collect();
        assert!(leftovers.is_empty());
        std::fs::remove_dir_all(&dir).ok();
    }
}
