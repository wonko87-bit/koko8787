import { api } from "./api";

/// 폴더를 탐색기로 연다. 폴더가 삭제·이동된 경우 조용히 실패하지 않고 알린다.
export function openFolder(path: string) {
  api.openFolder(path).catch((e) => alert(String(e)));
}
