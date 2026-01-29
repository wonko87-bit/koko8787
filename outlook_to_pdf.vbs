' ============================================================
' Outlook 메일 파일(.msg)을 PDF로 변환하는 VBScript
' 사용법: 스크립트 실행 후 폴더 선택 또는 스크립트 수정하여 폴더 경로 지정
' 요구사항: Microsoft Outlook, Microsoft Word 설치 필요
' ============================================================

Option Explicit

Dim objFSO, objShell, objOutlook, objWord
Dim strFolderPath, strOutputFolder
Dim objFolder, objFile
Dim convertedCount, errorCount

' 객체 생성
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' 폴더 선택 다이얼로그
Function SelectFolder(strPrompt)
    Dim objFolder
    Set objFolder = objShell.BrowseForFolder(0, strPrompt, 0, "")
    If objFolder Is Nothing Then
        SelectFolder = ""
    Else
        SelectFolder = objFolder.Self.Path
    End If
End Function

' 메인 실행
Sub Main()
    ' 입력 폴더 선택
    strFolderPath = SelectFolder("Outlook 메일 파일(.msg)이 있는 폴더를 선택하세요:")
    If strFolderPath = "" Then
        MsgBox "폴더가 선택되지 않았습니다. 프로그램을 종료합니다.", vbExclamation, "알림"
        WScript.Quit
    End If

    ' 출력 폴더 선택
    strOutputFolder = SelectFolder("PDF 파일을 저장할 폴더를 선택하세요:")
    If strOutputFolder = "" Then
        ' 출력 폴더 미선택 시 입력 폴더에 PDF 하위 폴더 생성
        strOutputFolder = strFolderPath & "\PDF_Output"
        If Not objFSO.FolderExists(strOutputFolder) Then
            objFSO.CreateFolder(strOutputFolder)
        End If
    End If

    ' Outlook 및 Word 애플리케이션 시작
    On Error Resume Next
    Set objOutlook = CreateObject("Outlook.Application")
    If Err.Number <> 0 Then
        MsgBox "Outlook을 시작할 수 없습니다. Outlook이 설치되어 있는지 확인하세요.", vbCritical, "오류"
        WScript.Quit
    End If

    Set objWord = CreateObject("Word.Application")
    If Err.Number <> 0 Then
        MsgBox "Word를 시작할 수 없습니다. Word가 설치되어 있는지 확인하세요.", vbCritical, "오류"
        WScript.Quit
    End If
    On Error GoTo 0

    objWord.Visible = False

    ' 카운터 초기화
    convertedCount = 0
    errorCount = 0

    ' 폴더 내 .msg 파일 처리
    Set objFolder = objFSO.GetFolder(strFolderPath)
    ProcessFolder objFolder

    ' Word 종료
    objWord.Quit
    Set objWord = Nothing
    Set objOutlook = Nothing

    ' 결과 메시지
    MsgBox "변환 완료!" & vbCrLf & vbCrLf & _
           "성공: " & convertedCount & " 개" & vbCrLf & _
           "실패: " & errorCount & " 개" & vbCrLf & vbCrLf & _
           "저장 위치: " & strOutputFolder, vbInformation, "완료"
End Sub

' 폴더 처리 (하위 폴더 포함)
Sub ProcessFolder(folder)
    Dim file, subfolder

    ' 현재 폴더의 .msg 파일 처리
    For Each file In folder.Files
        If LCase(objFSO.GetExtensionName(file.Name)) = "msg" Then
            ConvertMsgToPdf file.Path
        End If
    Next

    ' 하위 폴더 처리 (원하지 않으면 이 부분 주석 처리)
    For Each subfolder In folder.SubFolders
        ProcessFolder subfolder
    Next
End Sub

' MSG 파일을 PDF로 변환
Sub ConvertMsgToPdf(strMsgPath)
    Dim objMail, objDoc
    Dim strPdfPath, strBaseName
    Dim objNamespace

    On Error Resume Next

    ' 파일명 생성 (특수문자 제거)
    strBaseName = objFSO.GetBaseName(strMsgPath)
    strBaseName = CleanFileName(strBaseName)
    strPdfPath = strOutputFolder & "\" & strBaseName & ".pdf"

    ' 동일 파일명 존재 시 번호 추가
    Dim counter
    counter = 1
    Do While objFSO.FileExists(strPdfPath)
        strPdfPath = strOutputFolder & "\" & strBaseName & "_" & counter & ".pdf"
        counter = counter + 1
    Loop

    ' Outlook에서 .msg 파일 열기
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objMail = objNamespace.OpenSharedItem(strMsgPath)

    If Err.Number <> 0 Then
        WScript.Echo "오류 - 파일 열기 실패: " & strMsgPath
        errorCount = errorCount + 1
        Err.Clear
        Exit Sub
    End If

    ' Word 문서 생성 및 메일 내용 복사
    Set objDoc = objWord.Documents.Add

    ' 메일 정보 추가
    objDoc.Content.InsertAfter "보낸 사람: " & objMail.SenderName & " <" & objMail.SenderEmailAddress & ">" & vbCrLf
    objDoc.Content.InsertAfter "받는 사람: " & objMail.To & vbCrLf
    If objMail.CC <> "" Then
        objDoc.Content.InsertAfter "참조: " & objMail.CC & vbCrLf
    End If
    objDoc.Content.InsertAfter "날짜: " & objMail.ReceivedTime & vbCrLf
    objDoc.Content.InsertAfter "제목: " & objMail.Subject & vbCrLf
    objDoc.Content.InsertAfter String(50, "-") & vbCrLf & vbCrLf

    ' 메일 본문 추가 (HTML 또는 텍스트)
    If objMail.BodyFormat = 2 Then ' olFormatHTML
        ' HTML 본문인 경우 - HTMLBody를 Word로 붙여넣기
        Dim objTempDoc
        Set objTempDoc = objWord.Documents.Add
        objTempDoc.Content.InsertAfter objMail.Body
        objTempDoc.Content.Copy
        objDoc.Content.InsertAfter vbCrLf
        objDoc.Bookmarks("\EndOfDoc").Range.Paste
        objTempDoc.Close False
    Else
        ' 일반 텍스트
        objDoc.Content.InsertAfter objMail.Body
    End If

    ' PDF로 저장 (wdFormatPDF = 17)
    objDoc.SaveAs2 strPdfPath, 17

    If Err.Number <> 0 Then
        WScript.Echo "오류 - PDF 저장 실패: " & strMsgPath
        errorCount = errorCount + 1
        Err.Clear
    Else
        WScript.Echo "변환 완료: " & strBaseName & ".pdf"
        convertedCount = convertedCount + 1
    End If

    ' 정리
    objDoc.Close False
    objMail.Close 0 ' olDiscard

    Set objDoc = Nothing
    Set objMail = Nothing

    On Error GoTo 0
End Sub

' 파일명에서 사용 불가능한 문자 제거
Function CleanFileName(strName)
    Dim strResult, i, char
    Dim invalidChars
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    strResult = strName
    For i = 0 To UBound(invalidChars)
        strResult = Replace(strResult, invalidChars(i), "_")
    Next

    ' 앞뒤 공백 제거
    strResult = Trim(strResult)

    ' 파일명 길이 제한 (200자)
    If Len(strResult) > 200 Then
        strResult = Left(strResult, 200)
    End If

    CleanFileName = strResult
End Function

' 스크립트 실행
Main
