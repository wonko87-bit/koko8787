# FileBox → Flowdeck 브릿지 규격 v1

FileBox(파일 관리자)가 "나중에 읽어야 할 파일"을 Flowdeck(일정/할일)의 **할일**로
넘기기 위한 교환 규격이다. 두 앱은 서로를 실행하지도, 서로의 코드를 참조하지도
않는다. 오가는 것은 **감시 폴더에 놓인 텍스트 파일 하나**뿐이다.

이 문서는 양쪽 구현자가 상대편 코드를 보지 않고도 구현할 수 있도록 쓰였다.
`## Flowdeck 구현` 장만 읽으면 Flowdeck 쪽 작업은 전부 커버된다.

---

## 1. 전체 그림

```
 FileBox (Rust/Tauri)                    Flowdeck (C#/WPF)
 ─────────────────────                   ──────────────────
 특별규칙 매칭
      │
      ├─ 관리함의 파일에서 할일 데이터 생성
      │
      ├─ %APPDATA%\Flowdeck\inbox\xxxx.tmp  (쓰기)
      └─ %APPDATA%\Flowdeck\inbox\xxxx.txt  (원자적 rename)
                                              │
                            FileSystemWatcher ┤ Renamed 이벤트
                                              │
                                    TransferFile.Read(내용)
                                    Repository.ImportAsync(archive)
                                              │
                            성공 → inbox\처리됨\xxxx.txt 로 이동
                            실패 → inbox\실패\xxxx.txt 로 이동
```

핵심 원칙 네 가지:

1. **Flowdeck의 `workspace.json`은 Flowdeck만 쓴다.** FileBox는 절대 건드리지
   않는다. 외부에서 쓰면 실행 중인 Flowdeck이 메모리 상태를 그대로 덮어써서
   그 사이 추가된 항목이 사라진다.
2. **교환은 기존 `TransferArchive` 포맷을 그대로 재사용한다.** 새 포맷을 만들지
   않는다. `TransferFile.Read` + `WorkspaceRepository.ImportAsync`가 이미 파싱과
   중복 제거를 다 하고 있다.
3. **실시간성은 요구하지 않는다.** Flowdeck이 꺼져 있으면 파일은 그냥 쌓이고,
   다음 실행 시 시작 스캔이 주워 간다.
4. **파일은 지우지 않고 옮긴다.** 사용자가 나중에 뭐가 들어왔는지 확인할 수 있어야
   하고, 삭제는 되돌릴 수 없다.

---

## 2. 폴더 규약

| 경로 | 누가 만드나 | 용도 |
|---|---|---|
| `%APPDATA%\Flowdeck\inbox\` | 양쪽 다 (없으면 생성) | FileBox가 쓰는 곳, Flowdeck이 감시하는 곳 |
| `%APPDATA%\Flowdeck\inbox\처리됨\` | Flowdeck | 가져오기 성공한 파일 |
| `%APPDATA%\Flowdeck\inbox\실패\` | Flowdeck | 형식이 잘못돼 못 읽은 파일 |

`%APPDATA%\Flowdeck`을 고른 이유는 Flowdeck이 이미 `App.DataFolder`로 만들어
쓰고 있는 폴더라서다. 양쪽이 설정 파일을 주고받지 않고도 같은 경로를 독립적으로
계산할 수 있다.

두 앱 모두 이 경로를 **설정으로 덮어쓸 수 있어야 한다**(기본값은 위 경로).
회사 PC에서 `%APPDATA%`가 로밍 프로필이라 느린 경우 등을 위한 탈출구다.

`처리됨`/`실패`가 `inbox` **안에** 있는 것이 의도적이다. 감시자는
`IncludeSubdirectories = false`이므로 이 하위 폴더로 옮기는 동작이 감시자를
다시 깨우지 않는다. 폴더 하나로 브릿지 전체가 끝나서 사용자가 지우거나 옮기기도
쉽다.

---

## 3. 파일 규약

### 3.1 이름과 원자성

FileBox는 이렇게 쓴다.

```
1) %APPDATA%\Flowdeck\inbox\filebox-20260827-143210-3f9c1a52.tmp   ← 내용 전체 기록 + flush + close
2) 같은 폴더 안에서 .tmp → .txt 로 rename                            ← 원자적
```

- **Flowdeck은 `*.txt`만 본다. `.tmp`는 절대 열지 않는다.**
- 같은 볼륨·같은 디렉터리 안의 rename은 NTFS에서 원자적이다. 따라서 `.txt`가
  보이는 순간 그 파일은 **항상 완성된 파일**이다.
- 이 방식이 없으면 `FileSystemWatcher`의 `Created`가 파일 핸들 생성 시점에 즉시
  떠서, 아직 0바이트인 파일을 파싱하게 된다.
- 파일명 자체에는 의미가 없다. `filebox-` 접두사는 사람이 폴더를 열었을 때 알아보기
  위한 것이지 판별 기준이 아니다. 판별은 3.2의 `Format` 필드로 한다.

### 3.2 내용

`TransferFile.Write`가 만드는 것과 **같은 구조**다.

```
FileBox → Flowdeck · 2026-08-27 14:32 · 1건

- [할일] 2026 상반기 시장분석 리포트 · 9월 3일 10:00 #리포트 #읽기

--- 여기서부터는 앱이 읽는 부분입니다. 지우지 마세요 ---
{ ...JSON... }
```

- 마커 줄 위쪽(사람용 요약)은 **자유 형식**이다. `TransferFile.Read`는 마커의
  첫 등장 위치만 찾고 그 뒤를 JSON으로 파싱한다. FileBox가 헤더 문구를
  "Flowdeck 내보내기"가 아니라 "FileBox → Flowdeck"로 쓰는 건 그래서 안전하다.
- 마커 문자열은 **한 글자도 다르면 안 된다**:
  `--- 여기서부터는 앱이 읽는 부분입니다. 지우지 마세요 ---`
- 인코딩은 **UTF-8 (BOM 없음)**. 줄바꿈은 `\r\n`, `\n` 아무거나 무방.

### 3.3 JSON 스키마

`System.Text.Json`이 `PropertyNamingPolicy` 없이, `PropertyNameCaseInsensitive`
기본값(= false)으로 역직렬화한다. 따라서 **키는 C# 속성명 그대로 PascalCase**여야
하고 대소문자가 틀리면 그 필드는 조용히 무시된다.

```json
{
  "Format": "flowdeck.transfer",
  "Version": 1,
  "ExportedAt": "2026-08-27T14:32:10+09:00",
  "Todos": [
    {
      "Id": "3f9c1a52b7d4462e8a01c6d5e9f27b41",
      "Title": "[읽기] 2026 상반기 시장분석 리포트",
      "Notes": "파일: C:\\Users\\andrew\\Documents\\관리함\\2026_상반기_시장분석.pdf\n출처: FileBox 특별규칙 \"리포트 읽기\"",
      "DueAt": "2026-09-03T10:00:00",
      "HasTime": true,
      "Priority": "Normal",
      "Tags": ["리포트", "읽기"],
      "IsDone": false,
      "CreatedAt": "2026-08-27T14:32:10",
      "UpdatedAt": "2026-08-27T14:32:10",
      "SourceText": "2026_상반기_시장분석.pdf",
      "ReminderMinutesBefore": 30
    }
  ],
  "Events": []
}
```

필드별 규칙:

| 필드 | 타입 | 규칙 |
|---|---|---|
| `Format` | string | 반드시 `"flowdeck.transfer"`. 다르면 `Read`가 `FormatException`을 던진다 |
| `Version` | int | `1`. `TransferArchive.CurrentVersion`보다 크면 거부된다 |
| `ExportedAt` | DateTimeOffset | **오프셋 포함** ISO 8601 (`+09:00`) |
| `Todos[].Id` | string | 32자 소문자 hex, 하이픈 없음 (`Guid.ToString("N")` 형태). **중복 제거 키** — 4.3 참조 |
| `Todos[].Title` | string | 위젯에 보이는 한 줄. 비우지 말 것 |
| `Todos[].Notes` | string | 원본 파일의 **절대 경로를 반드시 포함**한다. 5.2 참조 |
| `Todos[].DueAt` | DateTime? | **오프셋도 `Z`도 붙이지 않는다.** 로컬 벽시계 시각 그대로 (`2026-09-03T10:00:00`). 시간대 표기를 붙이면 .NET이 그걸 시간대 정보로 해석해 값을 변환하는데, 규칙에서 지정한 건 시간대가 아니라 "그날 10시"라는 벽시계 값이다. 기한 없음이면 `null` 또는 필드 생략 |
| `Todos[].HasTime` | bool | 날짜만 지정했으면 `false` → UI가 시계를 숨긴다 |
| `Todos[].Priority` | string enum | `"None" \| "Low" \| "Normal" \| "High" \| "Urgent"` |
| `Todos[].Tags` | string[] | `#` 없이 순수 문자열 |
| `Todos[].SourceText` | string | 원본 파일명. 감사·재파싱용 |
| `Todos[].ReminderMinutesBefore` | int? | `DueAt` 기준 N분 전 알림. 없으면 생략 |
| `Todos[].Recurrence` | object | **생략한다.** 생략하면 C# 속성 초기화로 `Recurrence.None`이 된다 |
| `Todos[].ExternalLink` | object | **생략한다.** `ImportAsync`가 어차피 `null`로 지운다 |
| `Todos[].LinkedEventId` | string? | **생략한다.** 브릿지는 일정을 만들지 않는다 |
| `Events` | array | 항상 `[]`. 브릿지 v1은 할일만 만든다 |

`CreatedAt` / `UpdatedAt`을 생략해도 C# 쪽 속성 초기화가 `DateTime.Now`를 넣어주므로
동작에는 문제가 없다. 다만 "언제 만들어진 할일인가"가 가져오기 시점으로 바뀌므로
FileBox는 명시적으로 써 준다.

---

## 4. Flowdeck 구현

### 4.1 붙일 자리

새 서비스 하나를 추가하고 `App.StartAsync`에서 기동한다.

```
flowdeck/src/Flowdeck.Windows/Services/InboxWatcher.cs   (신규)
flowdeck/src/Flowdeck.Core/Settings/AppSettings.cs       (설정 2개 추가)
flowdeck/src/Flowdeck.Windows/Views/SettingsWindow.xaml  (설정 UI)
```

`Flowdeck.Core`가 아니라 `Flowdeck.Windows`에 두는 이유: `FileSystemWatcher`와
WPF `Dispatcher`에 의존하므로 Core의 "플랫폼 의존성 0" 원칙을 깬다. 모바일에는
이 기능이 없다(사용자가 모바일에서 파일을 열 일이 없다고 명시).

### 4.2 설정

`AppSettings`에 두 개를 더한다.

```csharp
/// <summary>
/// FileBox 같은 다른 앱이 던져 놓은 내보내기 파일을 자동으로 가져올지.
/// 감시 대상은 Flowdeck 자신의 데이터 폴더 안이라 바깥에서 손댈 수 없다.
/// </summary>
public bool EnableInboxWatch { get; set; } = true;

/// <summary>
/// 감시할 폴더. null이면 %APPDATA%\Flowdeck\inbox.
/// </summary>
public string? InboxFolder { get; set; }
```

기본값을 `true`로 두는 근거: 감시 대상이 자기 데이터 폴더 안이라 사용자 본인
외에는 아무것도 넣을 수 없고, 비용은 감시자 하나뿐이며, 무엇보다 FileBox 쪽에서
규칙을 만든 사용자가 Flowdeck에서 스위치를 하나 더 찾아 켜야 한다면 브릿지가
"안 되는 기능"으로 보인다.

### 4.3 처리 절차

```
시작 시:
  1. inbox, inbox\처리됨, inbox\실패 를 생성 (Directory.CreateDirectory)
  2. inbox\*.txt 를 전부 열거해 큐에 넣는다      ← 꺼져 있는 동안 쌓인 것
  3. FileSystemWatcher 기동

파일 하나 처리:
  1. 크기 확인. 4MB 초과면 실패\ 로 이동하고 끝  ← 폴더에 엉뚱한 게 떨어진 경우
  2. File.ReadAllText(path)
       - IOException(잠김) → 200/400/800/1600/3200ms 백오프로 최대 5회 재시도
       - 5회 다 실패 → 그대로 두고 로그만 남긴다 (다음 시작 스캔이 재시도)
  3. TransferFile.Read(내용)
       - FormatException → 실패\ 로 이동, 트레이 알림, 끝
  4. Repository.ImportAsync(archive)   ← UI 스레드에서
  5. 처리됨\ 으로 이동
  6. result.Added > 0 이면 트레이 알림: _tray?.Notify("Flowdeck", result.Describe())
```

`FileSystemWatcher` 설정:

```csharp
_watcher = new FileSystemWatcher(folder, "*.txt")
{
    NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
    IncludeSubdirectories = false,
};
_watcher.Created += OnAppeared;
_watcher.Renamed += OnAppeared;   // ← .tmp → .txt 는 Created가 아니라 Renamed로 온다
_watcher.Error   += OnWatcherError;
_watcher.EnableRaisingEvents = true;
```

**꼭 지켜야 할 것 넷:**

- **`Renamed`를 반드시 구독한다.** FileBox의 쓰기 방식(3.1) 때문에 정상 경로에서
  오는 이벤트는 `Created`가 아니라 `Renamed`다. **`Created`만 걸면 브릿지가 통째로
  동작하지 않는다** — 실제로 돌려 확인한 결과(부록 A) `.tmp → .txt` 리네임에서는
  `Created`가 단 한 번도 뜨지 않는다. 시작 스캔·수동 복사 대비로 `Created`도 같이 건다.
- **`Deleted`는 구독하지 않는다.** 처리 후 `처리됨\`으로 옮기는 동작이 감시 폴더
  기준으로는 삭제로 보여 `Deleted`가 뜬다. 여기에 로직을 걸면 자기 동작에 자기가 반응한다.
- **UI 스레드로 마샬링한다.** `FileSystemWatcher` 콜백은 스레드 풀에서 온다.
  `WorkspaceRepository`는 스레드 안전하지 않고 `Changed` 이벤트가 WPF 바인딩을
  건드리므로 `Application.Current.Dispatcher.InvokeAsync`로 넘긴다.
- **처리를 직렬화한다.** 파일 여러 개가 동시에 떨어질 수 있다. `SemaphoreSlim(1,1)`
  또는 단일 소비자 큐로 한 번에 하나씩 `ImportAsync` → `SaveAsync`가 끝나게 한다.
  겹치면 `workspace.json` 저장이 서로를 덮어쓴다.
- **같은 파일의 중복 이벤트를 무시한다.** `FileSystemWatcher`는 한 동작에 대해
  이벤트를 두 번 이상 올릴 수 있다. "처리 중인 경로" 집합으로 걸러 낸다.
  (걸러지지 않아도 4.4의 Id 중복 제거가 최종 방어선이라 데이터가 깨지진 않는다.)

이동 시 이름 충돌은 ` (2)`, ` (3)` … 을 붙여 피한다. 덮어쓰지 않는다.

### 4.4 중복 제거는 이미 되어 있다

`WorkspaceRepository.ImportAsync`는 이미 이렇게 동작한다.

```csharp
if (string.IsNullOrEmpty(todo.Id) || !todoIds.Add(todo.Id)) { skipped++; continue; }
```

즉 **같은 `Id`가 이미 있으면 건너뛰고 세기만 한다.** 따라서 같은 파일을 두 번
가져와도, 처리됨 폴더에서 다시 꺼내 놔도, FileBox가 실패로 판단해 다시 써도
할일이 두 개가 되지 않는다. 이게 성립하려면 **FileBox가 `Id`를 안정적으로**
발급해야 한다 — 5.3 참조.

또한 `ImportAsync`는 **추가만 하고 덮어쓰지 않는다.** 사용자가 가져온 할일의
제목을 고치거나 완료 처리한 뒤에 같은 파일이 다시 들어와도 그 작업이 되돌려지지
않는다는 뜻이다. 이 성질은 유지해야 한다.

### 4.5 Flowdeck이 하지 말아야 할 것

- `inbox\`에 **쓰지 않는다.** 그 폴더는 FileBox의 것이다. (`처리됨`/`실패` 하위
  폴더로 옮기는 것만 예외)
- 파일을 **지우지 않는다.** 이동만 한다.
- `.tmp`를 **열지 않는다.**
- 하위 폴더로 **재귀하지 않는다.** `처리됨`을 다시 읽으면 무한 루프가 된다.
- 가져온 할일에 `ExternalLink`를 **남기지 않는다.** (`ImportAsync`가 이미 처리)

---

## 5. FileBox 구현 (구현 완료 — v0.2.4)

Flowdeck 구현자가 상대편 동작을 이해하는 데 필요한 만큼만 적는다.

### 5.1 특별규칙

기존 규칙(확장자 / 파일명 키워드 / 카테고리 / 즐겨찾기)에 "Flowdeck 할일로 등록"
스위치와 아래 설정이 붙는다.

| 설정 | 값 | JSON 어디로 가나 |
|---|---|---|
| 제목 템플릿 | `[읽기] {파일명}` — `{파일명}` `{확장자}` `{카테고리}` 치환 | `Title` |
| 기한 | `없음` / `당일` / `1·3·7·14·30일 뒤` (등록일 기준) | `DueAt` |
| 시각 | `10:00` 형식, 비우면 날짜만 | `DueAt`의 시간부, `HasTime` |
| 우선순위 | 없음 / 낮음 / 보통 / 높음 / 긴급 | `Priority` |
| 태그 | `리포트, 읽기` | `Tags` |
| 알림 | 없음 / 10분 / 30분 / 1시간 / 하루 전 | `ReminderMinutesBefore` |

기한은 **고정 일수 오프셋만** 지원한다. "이번 주 금요일" 같은 상대 표현이나 파일명에서
날짜를 뽑아내는 것(계약서 파일명의 만기일 등)은 v1 범위 밖이다 — 그 일은 Flowdeck 의
자연어 파서가 훨씬 잘하므로, 옮겨 오는 대신 나중에 그쪽을 부르는 편이 낫다.

### 5.1.1 언제 보내는가 — 수집할 때가 아니라 옮길 때 (v0.2.5)

**할일은 파일이 최종 폴더로 옮겨진 뒤에 만든다.** 수집 시점이 아니다.

v0.2.4 는 수집하는 순간 보냈고, 그래서 메모에 적히는 경로가 언제나 관리함이었다.
관리함은 거쳐 가는 곳이므로, 사용자가 파일을 어디로 둘지 정하는 순간 그 경로는 죽는다.
정리를 마친 파일은 100% 링크가 끊긴 채로 남았다 — 예외가 아니라 기본 동작이었다.

메모를 나중에 고치는 길은 막혀 있다. `ImportAsync` 가 추가만 하고 덮어쓰지 않기
때문인데(4.4), 그건 재전송을 안전하게 만드는 성질이라 포기할 수 없다. 고칠 수 없다면
처음부터 맞게 쓰는 수밖에 없고, 경로가 확정되는 시점은 이동이 끝난 뒤다.

원칙으로 봐도 이쪽이 맞다. **관리함 자체가 "아직 안 치운 것" 목록이다.** 거기 있는
파일은 이미 눈에 보이니 할일이 필요 없다. 할일이 값어치를 갖는 건 파일이 관리함을
떠나 어느 폴더로 사라져서 더 이상 보이지 않게 되는 바로 그 순간이다.

규칙과 별개로 두 개의 수동 경로가 있다.

- **관리함의 📋 버튼**은 보내지 않고 **표시만** 한다("옮길 때 등록"). 실제 발송은
  파일을 옮길 때 일어나며, 규칙이 없어도 기한 없는 할일이 만들어진다. 규칙에도
  걸리는 파일이면 규칙 쪽만 나간다 — 둘 다 보내면 할일이 두 개가 된다.
- **기록 탭의 📋 버튼**은 이미 정리가 끝난 파일을 지금 보낸다. 경로가 이미 최종이라
  바로 보내도 안전하다.

5.3 때문에 어느 쪽이든 여러 번 눌러도 할일은 하나다.

남는 예외 하나: 정리까지 끝낸 파일을 나중에 **탐색기로 또 옮기는** 경우. FileBox 는
목적지 폴더를 감시하지 않아 알 방법이 없다. 이건 Flowdeck 쪽에서 파일이 없을 때
"파일을 찾을 수 없습니다" 로 알려 주는 것으로 충분하다.

### 5.2 `Notes`에 경로를 넣는 이유

`ImportAsync`가 `ExternalLink`를 무조건 `null`로 지우기 때문에, 할일에서 원본
파일을 다시 찾아갈 길은 `Notes` 본문밖에 없다. FileBox는 `Notes` 첫 줄을
`파일: <절대경로>` 로 쓴다.

> Flowdeck 쪽 선택 과제(필수 아님): 상세 화면에서 `Notes`의 `파일: ` 줄을
> 인식해 "폴더에서 보기" 버튼을 띄우면 체감이 훨씬 좋아진다. 규격상 요구사항은
> 아니고, `Notes`는 어디까지나 사람이 읽는 텍스트다.

### 5.3 `Id`의 안정성

FileBox는 `Id`를 **(항목 id, 규칙 id)에서 계산한다** — 고정 네임스페이스를 쓴 UUIDv5를
하이픈 없는 32자 hex로 만든다. 같은 파일을 같은 규칙으로 다시 보내면 몇 번을 계산해도
같은 `Id`가 나오므로 4.4의 중복 제거가 그대로 걸린다.

발송 성공 여부를 어딘가에 적어 두고 그걸 믿는 방법도 있었지만, 그러면 그 기록이
날아가는 순간 중복이 생긴다. 계산해서 얻으면 기록과 무관하게 성질이 유지된다.
(FileBox 도 보낸 흔적을 항목에 남기긴 하는데, 그건 화면에 "등록됨"을 표시하기 위한
것이지 중복 방지가 거기 기대고 있지는 않다.)

한 파일이 서로 다른 두 규칙에 걸리면 `Id`도 두 개, 할일도 두 개다. 의도된 동작이다.

### 5.4 배치

FileBox가 파일 여러 개를 한 번에 처리하면 `Todos` 배열에 여러 건을 담은
**파일 하나**를 쓸 수도 있고, 건별로 나눠 쓸 수도 있다. Flowdeck은 `Todos.Count`가
1이든 N이든 똑같이 처리해야 한다.

---

## 6. 검증 체크리스트 (Flowdeck 쪽)

- [ ] `inbox\`에 유효한 `.txt`를 복사 → 할일이 추가되고 파일이 `처리됨\`으로 이동
- [ ] `.tmp`로 만들었다가 `.txt`로 rename → `Renamed` 경로로 정상 처리
- [ ] `.tmp`인 상태로 놔둠 → **아무 일도 일어나지 않음**
- [ ] 같은 파일을 `처리됨\`에서 `inbox\`로 다시 복사 → `Skipped`로 집계, 할일 중복 없음
- [ ] JSON이 깨진 `.txt` → `실패\`로 이동, 앱은 살아 있음
- [ ] `Format`이 다른 JSON → `실패\`로 이동
- [ ] Flowdeck 종료 상태에서 3개 투입 후 실행 → 시작 스캔이 3개 모두 처리
- [ ] 10개를 동시에 투입 → 전부 처리되고 `workspace.json`에 유실 없음
- [ ] `DueAt: "2026-09-03T10:00:00"` → 위젯에 **10:00**으로 표시 (19:00 아님)
- [ ] `HasTime: false` → 시각 없이 날짜만 표시
- [ ] `EnableInboxWatch = false` → 파일을 넣어도 아무 일 없음
- [ ] 감시 폴더가 없는 상태로 실행 → 예외 없이 폴더를 만들고 기동

부록 A 에 이 항목들 중 상당수를 실제로 돌린 기록이 있다.

---

## 7. 범위 밖 (v1에서 하지 않는 것)

- Flowdeck → FileBox 역방향 통신
- 일정(`CalendarEvent`) 생성 — 할일만
- 파일명에서 날짜·금액 등을 추출하는 자연어 파싱
- 모바일(MAUI) 쪽 브릿지
- 할일 완료 시 원본 파일을 자동 정리하는 동작

---

## 부록 A. 검증 기록

이 문서의 규격은 추측이 아니라 실제 `Flowdeck.Core`(브랜치
`claude/windows-calendar-todo-app-1vd441`, `83c4d42`)를 .NET 8로 빌드해 돌려서
확인했다.

**3.3의 JSON 예시를 문서에서 그대로 뽑아** `TransferFile.Read` →
`WorkspaceRepository.ImportAsync`에 통과시킨 결과 24개 항목 전부 통과:

- `Read`가 마커 뒤 JSON을 파싱하고 `Format`/`Version` 검사를 통과
- `DueAt: "2026-09-03T10:00:00"` → `2026-09-03 10:00`, `Kind = Unspecified`
  (시간대 변환이 일어나지 않음을 확인)
- `Priority: "Normal"` 문자열이 enum으로 복원됨
- `Recurrence` / `ExternalLink` / `LinkedEventId`를 생략했을 때 각각
  `Recurrence.None` / `null` / `null`로 채워짐
- **`"Title"`을 `"title"`로 바꾸면 값이 조용히 빈 문자열이 된다** — 3.3의
  "PascalCase 필수"는 실제로 강제된다
- 같은 파일을 두 번 `ImportAsync` → `1건 추가` 후 `0건 추가, 1건 건너뜀`,
  할일은 1건 유지 (4.4의 중복 제거가 실동작)
- 깨진 JSON / 다른 `Format` / 미래 `Version` 셋 다 `FormatException`
  (→ `실패\`로 보낼 근거)

`FileSystemWatcher`(필터 `*.txt`, `IncludeSubdirectories = false`) 관측 결과:

| 동작 | 뜬 이벤트 |
|---|---|
| `filebox-1.tmp` 생성 + 기록 | **없음** |
| `.tmp` → `.txt` 리네임 | `Renamed  filebox-1.tmp -> filebox-1.txt` |
| `.txt` → `처리됨\` 이동 | `Deleted  filebox-1.txt` 뿐, 재진입 없음 |

`Created`는 이 흐름 전체에서 한 번도 뜨지 않았다.

(측정은 .NET 8 / Linux inotify 환경에서 했다. 필터·이벤트 매핑은 `FileSystemWatcher`
공통 계층의 동작이라 Windows에서도 같지만, 실제 배포 전에 Windows에서 위 표를
한 번 재현해 두는 편이 좋다.)

---

## 부록 B. 양쪽 구현 후 통합 검증

FileBox v0.2.4 와 Flowdeck `e2b2d16` 을 붙여서 실제로 돌린 결과다. 두 앱을 모두
빌드하고, FileBox 를 실행해 감시 폴더에 파일을 떨어뜨린 뒤, 거기서 나온 전송 파일을
Flowdeck 의 `InboxImporter` 에 그대로 먹였다. 재구현이 아니라 양쪽 다 진짜 코드다.

**보내는 쪽** — FileBox 를 Xvfb 에서 실행하고 특별규칙(`pdf` + `시장분석` →
7일 뒤 10:00, 높음, `#리포트`, 30분 전 알림)을 심은 상태로:

- 규칙에 걸리는 `2026_상반기_시장분석.pdf` → 전송 파일 생성됨
- 규칙에 안 걸리는 `영수증.pdf` → 수집만 되고 전송 파일 없음
- 관리함에 `.txt` 가 들어오지 않음 → 자기가 쓴 전송 파일을 도로 수집하지 않는다

**받는 쪽** — 그 파일을 `InboxImporter.DrainAsync()` 로:

| 확인 | 결과 |
|---|---|
| 대기 파일 인식 → `Imported` | 통과 |
| 제목 / 우선순위 / 태그 / 알림 | 통과 |
| `DueAt` = `2026-09-03 10:00`, `Kind = Unspecified` | 통과 (시간대 변환 없음) |
| `Notes` 첫 줄이 원본 절대경로 | 통과 |
| `Recurrence` 없음, `ExternalLink` `null` | 통과 |
| 처리 후 `처리됨\` 으로 이동, `실패\` 는 빈 채 | 통과 |
| 같은 파일 재투입 → `Skipped 1`, 할일은 여전히 1건 | 통과 |

20개 항목 전부 통과. FileBox 쪽에는 이 규격을 지키는지 확인하는 단위 테스트 10개가
따로 있다(PascalCase 키, 시간대 없는 `DueAt`, 자정 기본값, `Id` 안정성, `.tmp`→`.txt`).
