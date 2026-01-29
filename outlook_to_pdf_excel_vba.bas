' ============================================================
' Outlook 메일 파일(.msg)을 PDF로 변환하는 Excel VBA 매크로
'
' 사용법:
'   1. Excel에서 Alt + F11 → VBA 편집기 열기
'   2. 삽입 → 모듈 → 이 코드 붙여넣기
'   3. 도구 → 참조에서 다음 항목 체크:
'      - Microsoft Outlook XX.0 Object Library
'      - Microsoft Word XX.0 Object Library
'   4. 매크로 실행 (Alt + F8 → ConvertOutlookMailToPDF)
' ============================================================

Option Explicit

Sub ConvertOutlookMailToPDF()

    Dim fso As Object
    Dim outlookApp As Object
    Dim wordApp As Object
    Dim ns As Object

    Dim inputFolder As String
    Dim outputFolder As String
    Dim file As Object
    Dim mail As Object
    Dim doc As Object
    Dim pdfPath As String

    Dim successCount As Long
    Dim failCount As Long

    ' 폴더 선택
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Outlook 메일 파일(.msg)이 있는 폴더 선택"
        If .Show = -1 Then
            inputFolder = .SelectedItems(1)
        Else
            MsgBox "폴더가 선택되지 않았습니다.", vbExclamation
            Exit Sub
        End If
    End With

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "PDF 저장할 폴더 선택 (취소 시 자동 생성)"
        If .Show = -1 Then
            outputFolder = .SelectedItems(1)
        Else
            outputFolder = inputFolder & "\PDF_Output"
        End If
    End With

    ' 객체 생성
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 출력 폴더 생성
    If Not fso.FolderExists(outputFolder) Then
        fso.CreateFolder outputFolder
    End If

    ' Outlook, Word 시작
    On Error Resume Next
    Set outlookApp = CreateObject("Outlook.Application")
    Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0

    If outlookApp Is Nothing Or wordApp Is Nothing Then
        MsgBox "Outlook 또는 Word를 시작할 수 없습니다.", vbCritical
        Exit Sub
    End If

    wordApp.Visible = False
    Set ns = outlookApp.GetNamespace("MAPI")

    ' 상태 표시
    Application.StatusBar = "변환 중..."
    Application.ScreenUpdating = False

    successCount = 0
    failCount = 0

    ' .msg 파일 처리
    For Each file In fso.GetFolder(inputFolder).Files
        If LCase(fso.GetExtensionName(file.Name)) = "msg" Then

            On Error Resume Next

            ' 메일 열기
            Set mail = ns.OpenSharedItem(file.Path)

            If Err.Number <> 0 Then
                failCount = failCount + 1
                Err.Clear
                GoTo NextFile
            End If

            ' PDF 경로 생성
            pdfPath = outputFolder & "\" & CleanFileName(fso.GetBaseName(file.Name)) & ".pdf"
            pdfPath = GetUniquePath(pdfPath, fso)

            ' Word 문서 생성
            Set doc = wordApp.Documents.Add

            ' 메일 정보 입력
            With doc.Content
                .InsertAfter "보낸 사람: " & mail.SenderName & " <" & mail.SenderEmailAddress & ">" & vbCrLf
                .InsertAfter "받는 사람: " & mail.To & vbCrLf
                If mail.CC <> "" Then
                    .InsertAfter "참조: " & mail.CC & vbCrLf
                End If
                .InsertAfter "날짜: " & mail.ReceivedTime & vbCrLf
                .InsertAfter "제목: " & mail.Subject & vbCrLf
                .InsertAfter String(50, "-") & vbCrLf & vbCrLf
                .InsertAfter mail.Body
            End With

            ' PDF 저장
            doc.SaveAs2 pdfPath, 17  ' 17 = wdFormatPDF
            doc.Close False
            mail.Close 0  ' olDiscard

            If Err.Number = 0 Then
                successCount = successCount + 1
                Application.StatusBar = "변환 완료: " & fso.GetBaseName(file.Name)
            Else
                failCount = failCount + 1
                Err.Clear
            End If

            On Error GoTo 0

NextFile:
        End If
    Next file

    ' 정리
    wordApp.Quit
    Set wordApp = Nothing
    Set outlookApp = Nothing
    Set fso = Nothing

    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' 결과 메시지
    MsgBox "변환 완료!" & vbCrLf & vbCrLf & _
           "성공: " & successCount & "개" & vbCrLf & _
           "실패: " & failCount & "개" & vbCrLf & vbCrLf & _
           "저장 위치: " & outputFolder, vbInformation, "완료"

End Sub

' 파일명에서 사용 불가능한 문자 제거
Private Function CleanFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim i As Long

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(invalidChars) To UBound(invalidChars)
        fileName = Replace(fileName, invalidChars(i), "_")
    Next i

    fileName = Trim(fileName)

    If Len(fileName) > 200 Then
        fileName = Left(fileName, 200)
    End If

    CleanFileName = fileName
End Function

' 중복 파일명 처리
Private Function GetUniquePath(ByVal filePath As String, ByVal fso As Object) As String
    Dim basePath As String
    Dim ext As String
    Dim counter As Long

    If Not fso.FileExists(filePath) Then
        GetUniquePath = filePath
        Exit Function
    End If

    basePath = fso.GetParentFolderName(filePath) & "\" & fso.GetBaseName(filePath)
    ext = "." & fso.GetExtensionName(filePath)

    counter = 1
    Do While fso.FileExists(basePath & "_" & counter & ext)
        counter = counter + 1
    Loop

    GetUniquePath = basePath & "_" & counter & ext
End Function


' ============================================================
' 간단 버전 - 폴더 경로 직접 지정
' ============================================================

Sub ConvertOutlookMailToPDF_Simple()

    ' ★★★ 여기에 폴더 경로 입력 ★★★
    Const INPUT_FOLDER As String = "C:\Mail"
    Const OUTPUT_FOLDER As String = "C:\Mail\PDF"

    Dim fso As Object, outlookApp As Object, wordApp As Object, ns As Object
    Dim file As Object, mail As Object, doc As Object
    Dim pdfPath As String, cnt As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set outlookApp = CreateObject("Outlook.Application")
    Set wordApp = CreateObject("Word.Application")
    Set ns = outlookApp.GetNamespace("MAPI")

    wordApp.Visible = False

    If Not fso.FolderExists(OUTPUT_FOLDER) Then fso.CreateFolder OUTPUT_FOLDER

    cnt = 0
    For Each file In fso.GetFolder(INPUT_FOLDER).Files
        If LCase(fso.GetExtensionName(file.Name)) = "msg" Then
            On Error Resume Next

            Set mail = ns.OpenSharedItem(file.Path)
            Set doc = wordApp.Documents.Add

            doc.Content.InsertAfter "보낸 사람: " & mail.SenderName & vbCrLf
            doc.Content.InsertAfter "받는 사람: " & mail.To & vbCrLf
            doc.Content.InsertAfter "날짜: " & mail.ReceivedTime & vbCrLf
            doc.Content.InsertAfter "제목: " & mail.Subject & vbCrLf
            doc.Content.InsertAfter String(50, "-") & vbCrLf & vbCrLf
            doc.Content.InsertAfter mail.Body

            pdfPath = OUTPUT_FOLDER & "\" & fso.GetBaseName(file.Name) & ".pdf"
            doc.SaveAs2 pdfPath, 17
            doc.Close False
            mail.Close 0

            cnt = cnt + 1
            On Error GoTo 0
        End If
    Next file

    wordApp.Quit
    MsgBox cnt & "개 파일 변환 완료!", vbInformation

End Sub
