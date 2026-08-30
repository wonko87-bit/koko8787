# Sentinel

Galaxy Tab S10 Ultra 전용 **태블릿 사용통제 & 사후 모니터링** 앱.
개인 활용 목적. 기획서는 [artifact 링크](https://claude.ai/code/artifact/be4ec29d-88e7-4ad4-a9e5-afac3cabfb72) 참고.

기본 개념: 태블릿은 항상 **잠긴 벽돌(LOCKED)**. 관리자가 지문으로 열고 사용 시간을
부여할 때만 **ACTIVE**가 되며, 그동안의 활동을 기록한다. 콘텐츠 차단은 하지 않고
**사후 로그 기반 통제**를 지향한다.

## 모듈 구성

```
sentinel/
├── core/         순수 Kotlin. 상태머신·세션·정책(단위 테스트 포함, 안드로이드 의존 0)
├── app-tablet/   Master. 잠금·인증·활성화·기록. (com.sentinel)
└── app-phone/    Monitor(읽기 전용). 현재 스텁, M3에서 Firebase 연동. (com.sentinel.monitor)
```

## 현재 구현 범위 (M0 + M1)

- **M0** 멀티모듈 Gradle 스캐폴딩 (AGP 8.2.2 / Kotlin 1.9.22 / compileSdk 34 / Compose)
- **M1 코어 잠금 루프**
  - `core`: `TabletState`(LOCKED/ACTIVE/EXPIRED), `Session`(남은시간·만료 계산),
    `AccessPolicy`+`PolicyEvaluator`(자동 스케줄/일일예산 — 기본 비활성, D3) + 단위 테스트
  - Device Owner 컨트롤러(`DeviceOwnerController`) + 관리자 리시버(`SentinelAdminReceiver`)
  - 인증: 지문 `BiometricGate`(**BIOMETRIC_STRONG 전용 → 패턴/PIN 폴백 없음**),
    별도 PW `PasswordVault`(PBKDF2 + EncryptedSharedPreferences)
  - 상태 엔진 `SessionEngine`(StateFlow) + 포그라운드 감시 `LockService`(1초 틱·만료 재잠금) + `BootReceiver`
  - Compose UI: PW 설정 → 잠금 → 관리자 인증 → 시간 부여/승인 → ACTIVE HUD → 종료

> 아직 없음(다음 단계): 로그 수집(M2, UsageStats/화면이벤트), 폰 페어링·동기화(M3),
> 콘텐츠 로깅·자동 스케줄 UI(M4).

## 빌드

이 저장소 환경엔 Android SDK가 없어 **Android Studio** 또는 SDK가 있는 CI에서 빌드한다.

```bash
# 최초 1회: 래퍼 생성 (이 폴더에 gradle-wrapper.jar가 없음)
cd sentinel
gradle wrapper --gradle-version 8.4

# core 순수 로직 테스트 (SDK 불필요)
./gradlew :core:test

# APK (SDK 필요)
./gradlew :app-tablet:assembleDebug
./gradlew :app-phone:assembleDebug
```

`local.properties`에 `sdk.dir=/path/to/Android/sdk` 지정 필요.

## Device Owner 프로비저닝 (실제 강제력)

키오스크 잠금·삭제방지 등 실제 통제력은 **Device Owner**로만 나온다.
개발 중 일반 설치 시엔 `DeviceOwnerController`가 no-op이 되어 상태·UI 흐름은 그대로 시험 가능.

실기 세팅(최초 1회, **계정이 하나도 없는 상태** 필요):

1. 태블릿 **공장초기화** → 초기설정에서 계정 추가 없이 통과
2. 앱 설치: `adb install app-tablet/build/outputs/apk/debug/app-tablet-debug.apk`
3. Device Owner 지정:
   ```bash
   adb shell dpm set-device-owner com.sentinel/com.sentinel.tablet.admin.SentinelAdminReceiver
   ```
   > debug 빌드는 applicationId가 `com.sentinel.debug`이므로 그때는
   > `com.sentinel.debug/com.sentinel.tablet.admin.SentinelAdminReceiver` 로 지정.
4. 이후 Google 계정 로그인(폰 바인딩용, M3). 소유자 지정은 이미 끝나 영향 없음.

해제(초기화 전 정리): `adb shell dpm remove-active-admin com.sentinel/com.sentinel.tablet.admin.SentinelAdminReceiver`

## 알려진 실기 튜닝 포인트 (M1.5)

- **ACTIVE 중 홈 동작**: 현재 HOME 런처로 고정하지 않고 lock task로만 잠근다.
  Device Owner 배포 시 홈 고정/전용 런처 방식은 실기에서 결정.
- **지문 vs 얼굴**: STRONG만 허용 → 얼굴(대개 Class 2)은 배제되나, 기기 펌웨어 등급
  분류에 의존하므로 실기 1회 검증.
- **Firebase**: `google-services.json`은 커밋 대상 아님(.gitignore). 콘솔에서 생성해 배치.
