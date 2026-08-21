import { getVersion } from "@tauri-apps/api/app";
import { ask, message } from "@tauri-apps/plugin-dialog";
import { relaunch } from "@tauri-apps/plugin-process";
import { check } from "@tauri-apps/plugin-updater";

/// 업데이트를 확인하고, 있으면 사용자 동의를 받아 설치 후 재시작한다.
/// `manual`이 false면(시작 시 자동 확인) 최신 상태이거나 실패해도 조용히 넘어간다.
export async function checkForUpdate(manual: boolean) {
  try {
    const update = await check();
    if (!update) {
      if (manual) {
        await message(`이미 최신 버전이에요 (v${await getVersion()}).`, {
          title: "업데이트 확인",
        });
      }
      return;
    }

    const notes = update.body ? `\n\n${update.body}` : "";
    const ok = await ask(
      `새 버전 v${update.version}이 나왔어요.${notes}\n\n지금 설치할까요? 설치가 끝나면 FileBox가 다시 시작됩니다.`,
      { title: "업데이트 있음", kind: "info" },
    );
    if (!ok) return;

    await update.downloadAndInstall();
    await relaunch();
  } catch (e) {
    // 오프라인이거나 릴리스가 아직 없을 수 있으므로 자동 확인 실패는 알리지 않는다.
    if (manual) {
      await message(`업데이트를 확인하지 못했어요.\n${String(e)}`, {
        title: "업데이트 확인 실패",
        kind: "error",
      });
    }
  }
}
