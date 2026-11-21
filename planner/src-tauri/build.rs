fn main() {
    // 동기화용 OAuth 클라이언트는 빌드할 때 환경 변수로만 주입 (저장소에 넣지 않음)
    println!("cargo:rerun-if-env-changed=PLANNER_GOOGLE_CLIENT_ID");
    println!("cargo:rerun-if-env-changed=PLANNER_GOOGLE_CLIENT_SECRET");
    tauri_build::build()
}
