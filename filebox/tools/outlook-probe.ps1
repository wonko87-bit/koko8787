# Outlook 에서 메일 제목/보낸사람을 가져올 수 있는지 확인하는 스크립트.
#
# 읽기만 한다. 메일을 보내거나 고치거나 지우지 않고, 아무것도 저장하지 않는다.
# Outlook 을 켜 둔 상태에서 PowerShell 창에 이 파일 경로를 넣어 실행하면 된다.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File outlook-probe.ps1

$ErrorActionPreference = 'Stop'

Write-Host "=== FileBox · Outlook 연동 가능 여부 확인 ===" -ForegroundColor Cyan
Write-Host ""

# 1) 클래식 Outlook 인가. '새 Outlook' 에는 COM 객체 모델이 아예 없다.
if ($null -eq [Type]::GetTypeFromProgID('Outlook.Application')) {
    Write-Host "[실패] Outlook.Application 을 찾을 수 없습니다." -ForegroundColor Red
    Write-Host "       '새 Outlook' 을 쓰고 계시거나 클래식 Outlook 이 설치돼 있지 않습니다."
    Write-Host "       이 경우 이 방식은 쓸 수 없습니다."
    exit 1
}
Write-Host "[통과] 클래식 Outlook 이 등록돼 있습니다." -ForegroundColor Green

# 2) 이미 떠 있는가. 안 떠 있으면 아래 연결이 Outlook 을 실행시킨다.
$running = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
if ($running) {
    Write-Host "[통과] Outlook 이 실행 중입니다." -ForegroundColor Green
} else {
    Write-Host "[주의] Outlook 이 꺼져 있습니다. 지금 실행됩니다." -ForegroundColor Yellow
    Write-Host "       (실제 기능에서는 꺼져 있으면 그냥 건너뛰게 만들 예정입니다.)"
}

# 3) 실제로 받은편지함을 읽어 본다. 보안 가드가 걸리면 여기서 걸린다.
try {
    $outlook = New-Object -ComObject Outlook.Application
    $inbox = $outlook.GetNamespace('MAPI').GetDefaultFolder(6)  # 6 = 받은편지함
    $items = $inbox.Items
    # Restrict 의 날짜 형식은 지역 설정을 타서 잘 깨진다. 확인용으로는 정렬만 한다.
    $items.Sort('[ReceivedTime]', $true)
} catch {
    Write-Host "[실패] Outlook 에 연결하지 못했습니다: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "       회사 정책이나 백신이 자동화를 막고 있을 수 있습니다."
    exit 1
}
Write-Host "[통과] 받은편지함을 열었습니다." -ForegroundColor Green
Write-Host ""

# 4) 최근 메일에서 첨부파일 → (제목, 보낸사람) 을 뽑아 본다.
Write-Host "최근 메일 200통에서 첨부파일이 있는 것들:" -ForegroundColor Cyan
Write-Host ""

$found = 0
$blocked = 0
$scanned = 0

foreach ($mail in $items) {
    $scanned++
    if ($scanned -gt 200 -or $found -ge 15) { break }

    try {
        if ($mail.Attachments.Count -eq 0) { continue }
    } catch { continue }

    foreach ($attachment in $mail.Attachments) {
        try {
            $name = $attachment.FileName
            # 본문에 박힌 이미지 등은 첨부로 세면 안 된다.
            if ([string]::IsNullOrWhiteSpace($name)) { continue }

            $subject = $mail.Subject
            $sender  = $mail.SenderName
            $when    = $mail.ReceivedTime

            # 주소 계열은 보안 가드가 막는 경우가 있어 따로 시험한다.
            $address = '(못 읽음)'
            try { $address = $mail.SenderEmailAddress } catch { $blocked++ }

            Write-Host ("  파일    : {0}" -f $name)
            Write-Host ("  제목    : {0}" -f $subject)
            Write-Host ("  보낸이  : {0}  <{1}>" -f $sender, $address)
            Write-Host ("  받은날짜: {0}" -f $when)
            Write-Host ""
            $found++
            if ($found -ge 15) { break }
        } catch {
            Write-Host "  (이 첨부파일은 읽지 못했습니다: $($_.Exception.Message))" -ForegroundColor Yellow
        }
    }
}

Write-Host "=== 결과 ===" -ForegroundColor Cyan
if ($found -eq 0) {
    Write-Host "첨부파일이 있는 메일을 못 찾았습니다. 최근 메일에 첨부가 없거나," -ForegroundColor Yellow
    Write-Host "받은편지함이 아닌 다른 폴더로 분류되고 있을 수 있습니다."
} else {
    Write-Host "제목과 보낸사람을 정상적으로 읽었습니다. ($found 건)" -ForegroundColor Green
    if ($blocked -gt 0) {
        Write-Host "다만 메일 주소는 $blocked 건에서 막혔습니다 — 이름만 쓰면 됩니다." -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "위 목록에서 '파일' 이름이 실제로 저장했을 때의 파일명과 같은지 봐 주세요."
    Write-Host "그게 FileBox 가 메일을 찾아낼 유일한 단서입니다."
}
