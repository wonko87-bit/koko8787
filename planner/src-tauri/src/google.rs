//! Google 로그인(OAuth 2.0 + PKCE)과 드라이브 앱 전용 저장공간(appDataFolder) 읽기/쓰기.
//!
//! 로그인 흐름
//!   1. 127.0.0.1의 빈 포트에 잠깐 수신기를 연다.
//!   2. 브라우저로 Google 로그인 페이지를 연다. redirect_uri는 GitHub Pages 중계 페이지이고,
//!      state = "d.<포트>.<임의값>" 으로 돌아올 곳을 알려 준다.
//!   3. 중계 페이지가 http://127.0.0.1:<포트>/callback?code=… 로 넘겨 주면 code를 받아
//!      PKCE verifier와 함께 토큰으로 바꾼다.
//!   4. refresh token은 Windows 자격 증명 관리자(그 밖의 OS는 앱 로컬 데이터 폴더)에 저장한다.
//!
//! OAuth 클라이언트 ID·시크릿은 저장소에 없고, 빌드할 때 환경 변수
//! PLANNER_GOOGLE_CLIENT_ID / PLANNER_GOOGLE_CLIENT_SECRET 로만 들어온다.
//! 둘 중 하나라도 없으면 동기화 기능은 꺼진 빌드가 된다(회사 PC용).

use std::{
    io::{Read, Write},
    net::{TcpListener, TcpStream},
    sync::Mutex,
    time::{Duration, Instant, SystemTime, UNIX_EPOCH},
};

use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine};
use serde::Deserialize;
use sha2::{Digest, Sha256};
use tauri::AppHandle;
use tauri_plugin_opener::OpenerExt;

const AUTH_URL: &str = "https://accounts.google.com/o/oauth2/v2/auth";
const TOKEN_URL: &str = "https://oauth2.googleapis.com/token";
const REVOKE_URL: &str = "https://oauth2.googleapis.com/revoke";
const DRIVE_FILES: &str = "https://www.googleapis.com/drive/v3/files";
const DRIVE_UPLOAD: &str = "https://www.googleapis.com/upload/drive/v3/files";
pub const RELAY_URL: &str = "https://wonko87-bit.github.io/koko8787/auth/planner/callback/";
pub const SCOPE: &str = "https://www.googleapis.com/auth/drive.appdata";
const SYNC_FILE: &str = "planner-sync.json";
const LOGIN_TIMEOUT: Duration = Duration::from_secs(300);
pub const LOGIN_REQUIRED: &str = "LOGIN_REQUIRED";

pub fn client() -> Option<(&'static str, &'static str)> {
    match (option_env!("PLANNER_GOOGLE_CLIENT_ID"), option_env!("PLANNER_GOOGLE_CLIENT_SECRET")) {
        (Some(id), Some(secret)) if !id.trim().is_empty() && !secret.trim().is_empty() => {
            Some((id.trim(), secret.trim()))
        }
        _ => None,
    }
}

#[derive(Default)]
pub struct SyncState {
    access_token: Option<String>,
    expires_at: Option<Instant>,
    file_ids: Vec<String>,
}

pub type Shared = Mutex<SyncState>;

/* ───────────────────────── 토큰 저장 ───────────────────────── */

mod store {
    #[cfg(windows)]
    fn entry() -> Result<keyring::Entry, String> {
        keyring::Entry::new("Planner", "google-refresh-token").map_err(|e| e.to_string())
    }

    #[cfg(windows)]
    pub fn load(_app: &tauri::AppHandle) -> Option<String> {
        entry().ok()?.get_password().ok()
    }
    #[cfg(windows)]
    pub fn save(_app: &tauri::AppHandle, token: &str) -> Result<(), String> {
        entry()?.set_password(token).map_err(|e| e.to_string())
    }
    #[cfg(windows)]
    pub fn clear(_app: &tauri::AppHandle) {
        if let Ok(e) = entry() {
            let _ = e.delete_credential();
        }
    }

    #[cfg(not(windows))]
    fn path(app: &tauri::AppHandle) -> Option<std::path::PathBuf> {
        use tauri::Manager;
        let dir = app.path().app_local_data_dir().ok()?;
        std::fs::create_dir_all(&dir).ok()?;
        Some(dir.join("google-auth.txt"))
    }
    #[cfg(not(windows))]
    pub fn load(app: &tauri::AppHandle) -> Option<String> {
        std::fs::read_to_string(path(app)?).ok().map(|s| s.trim().to_string()).filter(|s| !s.is_empty())
    }
    #[cfg(not(windows))]
    pub fn save(app: &tauri::AppHandle, token: &str) -> Result<(), String> {
        let p = path(app).ok_or("저장 위치를 찾을 수 없음")?;
        std::fs::write(p, token).map_err(|e| e.to_string())
    }
    #[cfg(not(windows))]
    pub fn clear(app: &tauri::AppHandle) {
        if let Some(p) = path(app) {
            let _ = std::fs::remove_file(p);
        }
    }
}

/* ───────────────────────── PKCE / URL ───────────────────────── */

fn random_token(bytes: usize) -> Result<String, String> {
    let mut buf = vec![0u8; bytes];
    getrandom::getrandom(&mut buf).map_err(|e| e.to_string())?;
    Ok(URL_SAFE_NO_PAD.encode(buf))
}

pub fn pkce_challenge(verifier: &str) -> String {
    URL_SAFE_NO_PAD.encode(Sha256::digest(verifier.as_bytes()))
}

pub fn auth_url(client_id: &str, state: &str, challenge: &str) -> String {
    let mut u = url::Url::parse(AUTH_URL).expect("static url");
    u.query_pairs_mut()
        .append_pair("client_id", client_id)
        .append_pair("redirect_uri", RELAY_URL)
        .append_pair("response_type", "code")
        .append_pair("scope", SCOPE)
        .append_pair("state", state)
        .append_pair("code_challenge", challenge)
        .append_pair("code_challenge_method", "S256")
        .append_pair("access_type", "offline")
        .append_pair("prompt", "consent");
    u.into()
}

/// "GET /callback?code=..&state=.. HTTP/1.1" → (path, query 쌍)
pub fn parse_request_line(line: &str) -> Option<(String, Vec<(String, String)>)> {
    let target = line.split_whitespace().nth(1)?;
    let u = url::Url::parse(&format!("http://127.0.0.1{target}")).ok()?;
    let q = u.query_pairs().map(|(k, v)| (k.into_owned(), v.into_owned())).collect();
    Some((u.path().to_string(), q))
}

/* ───────────────────────── HTTP 도우미 ───────────────────────── */

fn agent() -> ureq::Agent {
    ureq::AgentBuilder::new()
        .timeout_connect(Duration::from_secs(10))
        .timeout(Duration::from_secs(30))
        .build()
}

fn http_err(e: ureq::Error) -> String {
    match e {
        ureq::Error::Status(code, resp) => {
            let body = resp.into_string().unwrap_or_default();
            format!("HTTP {code}: {}", body.chars().take(300).collect::<String>())
        }
        ureq::Error::Transport(t) => format!("네트워크 오류: {t}"),
    }
}

#[derive(Deserialize)]
struct TokenResp {
    access_token: String,
    expires_in: Option<u64>,
    refresh_token: Option<String>,
}

fn now_secs() -> u64 {
    SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.as_secs()).unwrap_or(0)
}

/* ───────────────────────── 로그인 ───────────────────────── */

fn respond(stream: &mut TcpStream, ok: bool) {
    let (title, msg) = if ok {
        ("로그인 완료", "Planner와 Google 드라이브가 연결됐어요. 이 창을 닫고 Planner로 돌아가세요.")
    } else {
        ("로그인 실패", "로그인을 마치지 못했어요. Planner에서 다시 시도해 주세요.")
    };
    let html = format!(
        "<!DOCTYPE html><html lang=\"ko\"><head><meta charset=\"UTF-8\"><title>{title}</title></head>\
         <body style=\"font-family:'Malgun Gothic',system-ui,sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:#F2F4F8;color:#1B2433\">\
         <div style=\"background:#fff;border:1px solid #E3E8EF;border-radius:16px;padding:28px 36px;text-align:center\">\
         <h2 style=\"margin:0 0 8px;color:#1F3A5F\">{title}</h2><p style=\"margin:0\">{msg}</p></div></body></html>"
    );
    let resp = format!(
        "HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{html}",
        html.len()
    );
    let _ = stream.write_all(resp.as_bytes());
    let _ = stream.flush();
}

/// 브라우저 로그인 → 토큰 저장. 블로킹 (별도 스레드에서 호출)
pub fn login(app: &AppHandle, shared: &Shared) -> Result<(), String> {
    let (client_id, secret) = client().ok_or("이 빌드에는 동기화 기능이 없어요")?;
    let listener = TcpListener::bind("127.0.0.1:0").map_err(|e| e.to_string())?;
    let port = listener.local_addr().map_err(|e| e.to_string())?.port();
    listener.set_nonblocking(true).map_err(|e| e.to_string())?;

    let verifier = random_token(48)?;
    let state = format!("d.{port}.{}", random_token(12)?);
    let url = auth_url(client_id, &state, &pkce_challenge(&verifier));
    app.opener().open_url(url, None::<&str>).map_err(|e| format!("브라우저를 열 수 없음: {e}"))?;

    // 콜백 기다리기 (최대 5분)
    let started = Instant::now();
    let code = loop {
        if started.elapsed() > LOGIN_TIMEOUT {
            return Err("로그인 시간이 초과됐어요 (5분)".into());
        }
        match listener.accept() {
            Ok((mut stream, _)) => {
                let _ = stream.set_nonblocking(false);
                let _ = stream.set_read_timeout(Some(Duration::from_secs(5)));
                let mut buf = [0u8; 8192];
                let n = stream.read(&mut buf).unwrap_or(0);
                let req = String::from_utf8_lossy(&buf[..n]);
                let Some((path, q)) = req.lines().next().and_then(parse_request_line) else { continue };
                if path != "/callback" {
                    let _ = stream.write_all(b"HTTP/1.1 404 Not Found\r\nContent-Length: 0\r\nConnection: close\r\n\r\n");
                    continue;
                }
                let get = |k: &str| q.iter().find(|(a, _)| a == k).map(|(_, v)| v.clone());
                if get("state").as_deref() != Some(state.as_str()) {
                    respond(&mut stream, false);
                    return Err("로그인 응답이 올바르지 않아요 (state 불일치)".into());
                }
                if let Some(err) = get("error") {
                    respond(&mut stream, false);
                    return Err(format!("Google 로그인이 취소됐어요 ({err})"));
                }
                match get("code") {
                    Some(c) if !c.is_empty() => {
                        respond(&mut stream, true);
                        break c;
                    }
                    _ => {
                        respond(&mut stream, false);
                        return Err("로그인 코드가 없어요".into());
                    }
                }
            }
            Err(e) if e.kind() == std::io::ErrorKind::WouldBlock => std::thread::sleep(Duration::from_millis(150)),
            Err(e) => return Err(e.to_string()),
        }
    };

    let resp: TokenResp = agent()
        .post(TOKEN_URL)
        .send_form(&[
            ("client_id", client_id),
            ("client_secret", secret),
            ("code", &code),
            ("code_verifier", &verifier),
            ("grant_type", "authorization_code"),
            ("redirect_uri", RELAY_URL),
        ])
        .map_err(http_err)?
        .into_json()
        .map_err(|e| e.to_string())?;
    let refresh = resp.refresh_token.ok_or("refresh token을 받지 못했어요. 다시 로그인해 주세요")?;
    store::save(app, &refresh)?;
    let mut st = shared.lock().map_err(|e| e.to_string())?;
    st.access_token = Some(resp.access_token);
    st.expires_at = Some(Instant::now() + Duration::from_secs(resp.expires_in.unwrap_or(3600)));
    st.file_ids.clear();
    Ok(())
}

pub fn is_connected(app: &AppHandle) -> bool {
    client().is_some() && store::load(app).is_some()
}

pub fn logout(app: &AppHandle, shared: &Shared) {
    if let Some(tok) = store::load(app) {
        let _ = agent().post(REVOKE_URL).send_form(&[("token", tok.as_str())]);
    }
    store::clear(app);
    if let Ok(mut st) = shared.lock() {
        *st = SyncState::default();
    }
}

/// 유효한 access token (만료 1분 전이면 refresh)
fn access_token(app: &AppHandle, shared: &Shared, force: bool) -> Result<String, String> {
    {
        let st = shared.lock().map_err(|e| e.to_string())?;
        if !force {
            if let (Some(t), Some(exp)) = (&st.access_token, st.expires_at) {
                if Instant::now() + Duration::from_secs(60) < exp {
                    return Ok(t.clone());
                }
            }
        }
    }
    let (client_id, secret) = client().ok_or("이 빌드에는 동기화 기능이 없어요")?;
    let refresh = store::load(app).ok_or(LOGIN_REQUIRED)?;
    let res = agent().post(TOKEN_URL).send_form(&[
        ("client_id", client_id),
        ("client_secret", secret),
        ("refresh_token", refresh.as_str()),
        ("grant_type", "refresh_token"),
    ]);
    let resp: TokenResp = match res {
        Ok(r) => r.into_json().map_err(|e| e.to_string())?,
        Err(ureq::Error::Status(400 | 401, r)) => {
            let body = r.into_string().unwrap_or_default();
            if body.contains("invalid_grant") {
                // 권한 취소·만료 → 다시 로그인 필요
                store::clear(app);
                return Err(LOGIN_REQUIRED.into());
            }
            return Err(format!("토큰 갱신 실패: {body}"));
        }
        Err(e) => return Err(http_err(e)),
    };
    let mut st = shared.lock().map_err(|e| e.to_string())?;
    st.access_token = Some(resp.access_token.clone());
    st.expires_at = Some(Instant::now() + Duration::from_secs(resp.expires_in.unwrap_or(3600)));
    Ok(resp.access_token)
}

/// 401이면 토큰을 새로 받아 한 번 더 시도
fn with_token<T>(app: &AppHandle, shared: &Shared, f: impl Fn(&str) -> Result<T, ureq::Error>) -> Result<T, String> {
    let tok = access_token(app, shared, false)?;
    match f(&tok) {
        Ok(v) => Ok(v),
        Err(ureq::Error::Status(401, _)) => {
            let tok = access_token(app, shared, true)?;
            f(&tok).map_err(http_err)
        }
        Err(e) => Err(http_err(e)),
    }
}

/* ───────────────────────── 드라이브 ───────────────────────── */

#[derive(Deserialize)]
struct FileList {
    files: Vec<FileMeta>,
}
#[derive(Deserialize)]
struct FileMeta {
    id: String,
}

/// appDataFolder의 planner-sync.json 내용들 (최신순). 여러 개면 모두 돌려주고 다음 업로드 때 하나로 정리
pub fn pull(app: &AppHandle, shared: &Shared) -> Result<Vec<String>, String> {
    let ag = agent();
    let list: FileList = with_token(app, shared, |t| {
        ag.get(DRIVE_FILES)
            .set("Authorization", &format!("Bearer {t}"))
            .query("spaces", "appDataFolder")
            .query("q", &format!("name = '{SYNC_FILE}' and trashed = false"))
            .query("orderBy", "modifiedTime desc")
            .query("fields", "files(id)")
            .query("pageSize", "10")
            .call()
    })?
    .into_json()
    .map_err(|e| e.to_string())?;

    let mut out = Vec::new();
    for f in &list.files {
        let body = with_token(app, shared, |t| {
            ag.get(&format!("{DRIVE_FILES}/{}", f.id))
                .set("Authorization", &format!("Bearer {t}"))
                .query("alt", "media")
                .call()
        })?
        .into_string()
        .map_err(|e| e.to_string())?;
        out.push(body);
    }
    let mut st = shared.lock().map_err(|e| e.to_string())?;
    st.file_ids = list.files.into_iter().map(|f| f.id).collect();
    Ok(out)
}

/// 병합본을 올림. 파일이 없으면 만들고, 중복 파일은 지움
pub fn push(app: &AppHandle, shared: &Shared, content: &str) -> Result<(), String> {
    let ids = shared.lock().map_err(|e| e.to_string())?.file_ids.clone();
    let ag = agent();
    if let Some(first) = ids.first() {
        with_token(app, shared, |t| {
            ag.request("PATCH", &format!("{DRIVE_UPLOAD}/{first}"))
                .set("Authorization", &format!("Bearer {t}"))
                .query("uploadType", "media")
                .set("Content-Type", "application/json; charset=utf-8")
                .send_string(content)
        })?;
        for extra in ids.iter().skip(1) {
            let _ = with_token(app, shared, |t| {
                ag.delete(&format!("{DRIVE_FILES}/{extra}"))
                    .set("Authorization", &format!("Bearer {t}"))
                    .call()
            });
        }
        shared.lock().map_err(|e| e.to_string())?.file_ids.truncate(1);
    } else {
        let boundary = format!("planner-{}", now_secs());
        let body = format!(
            "--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n\
             {{\"name\":\"{SYNC_FILE}\",\"parents\":[\"appDataFolder\"]}}\r\n\
             --{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{content}\r\n--{boundary}--"
        );
        let created: FileMeta = with_token(app, shared, |t| {
            ag.post(DRIVE_UPLOAD)
                .set("Authorization", &format!("Bearer {t}"))
                .query("uploadType", "multipart")
                .query("fields", "id")
                .set("Content-Type", &format!("multipart/related; boundary={boundary}"))
                .send_string(&body)
        })?
        .into_json()
        .map_err(|e| e.to_string())?;
        shared.lock().map_err(|e| e.to_string())?.file_ids = vec![created.id];
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pkce_matches_rfc7636_example() {
        // RFC 7636 Appendix B
        assert_eq!(
            pkce_challenge("dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk"),
            "E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM"
        );
    }

    #[test]
    fn auth_url_has_required_params() {
        let u = url::Url::parse(&auth_url("cid.apps.googleusercontent.com", "d.51234.abc", "chal")).unwrap();
        let q: std::collections::HashMap<_, _> = u.query_pairs().into_owned().collect();
        assert_eq!(q["redirect_uri"], RELAY_URL);
        assert_eq!(q["scope"], SCOPE);
        assert_eq!(q["state"], "d.51234.abc");
        assert_eq!(q["code_challenge_method"], "S256");
        assert_eq!(q["access_type"], "offline");
    }

    #[test]
    fn parses_callback_request() {
        let (p, q) = parse_request_line("GET /callback?code=4%2F0Ab&state=d.5.x HTTP/1.1").unwrap();
        assert_eq!(p, "/callback");
        assert!(q.contains(&("code".into(), "4/0Ab".into())));
        assert!(q.contains(&("state".into(), "d.5.x".into())));
        assert!(parse_request_line("garbage").is_none());
    }

    #[test]
    fn random_tokens_are_unique_and_url_safe() {
        let a = random_token(48).unwrap();
        let b = random_token(48).unwrap();
        assert_ne!(a, b);
        assert!(a.chars().all(|c| c.is_ascii_alphanumeric() || c == '-' || c == '_'));
        assert!(a.len() >= 43, "PKCE verifier 최소 길이");
    }
}
