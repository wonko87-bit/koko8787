' ============================================================
' EML 메일 파일을 PDF로 변환하는 Excel VBA 매크로
'
' 사용법:
'   1. Excel에서 Alt + F11 → VBA 편집기 열기
'   2. 삽입 → 모듈 → 이 코드 붙여넣기
'   3. 도구 → 참조에서 다음 항목 체크:
'      - Microsoft Word XX.0 Object Library
'      - Microsoft CDO for Windows 2000 Library (또는 CDO 1.21)
'   4. 매크로 실행 (Alt + F8 → ConvertEmlToPDF)
' ============================================================

Option Explicit

Sub ConvertEmlToPDF()

    Dim fso As Object
    Dim wordApp As Object
    Dim cdoMsg As Object

    Dim inputFolder As String
    Dim outputFolder As String
    Dim file As Object
    Dim doc As Object
    Dim pdfPath As String

    Dim successCount As Long
    Dim failCount As Long

    ' 폴더 선택
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "EML 메일 파일이 있는 폴더 선택"
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

    ' Word 시작
    On Error Resume Next
    Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0

    If wordApp Is Nothing Then
        MsgBox "Word를 시작할 수 없습니다.", vbCritical
        Exit Sub
    End If

    wordApp.Visible = False

    ' 상태 표시
    Application.StatusBar = "변환 중..."
    Application.ScreenUpdating = False

    successCount = 0
    failCount = 0

    ' .eml 파일 처리
    For Each file In fso.GetFolder(inputFolder).Files
        If LCase(fso.GetExtensionName(file.Name)) = "eml" Then

            On Error Resume Next

            ' CDO Message 객체로 EML 읽기
            Set cdoMsg = CreateObject("CDO.Message")

            ' EML 파일 스트림으로 로드
            Dim adoStream As Object
            Set adoStream = CreateObject("ADODB.Stream")
            adoStream.Open
            adoStream.LoadFromFile file.Path
            cdoMsg.DataSource.OpenObject adoStream, "_Stream"
            adoStream.Close

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
                .InsertAfter "보낸 사람: " & cdoMsg.From & vbCrLf
                .InsertAfter "받는 사람: " & cdoMsg.To & vbCrLf
                If cdoMsg.CC <> "" Then
                    .InsertAfter "참조: " & cdoMsg.CC & vbCrLf
                End If
                .InsertAfter "날짜: " & cdoMsg.SentOn & vbCrLf
                .InsertAfter "제목: " & cdoMsg.Subject & vbCrLf
                .InsertAfter String(50, "-") & vbCrLf & vbCrLf
                .InsertAfter cdoMsg.TextBody
            End With

            ' PDF 저장
            doc.SaveAs2 pdfPath, 17  ' 17 = wdFormatPDF
            doc.Close False

            If Err.Number = 0 Then
                successCount = successCount + 1
                Application.StatusBar = "변환 완료: " & fso.GetBaseName(file.Name)
            Else
                failCount = failCount + 1
                Err.Clear
            End If

            Set cdoMsg = Nothing
            On Error GoTo 0

NextFile:
        End If
    Next file

    ' 정리
    wordApp.Quit
    Set wordApp = Nothing
    Set fso = Nothing

    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' 결과 메시지
    MsgBox "변환 완료!" & vbCrLf & vbCrLf & _
           "성공: " & successCount & "개" & vbCrLf & _
           "실패: " & failCount & "개" & vbCrLf & vbCrLf & _
           "저장 위치: " & outputFolder, vbInformation, "완료"

End Sub


' ============================================================
' 간단 버전 - 폴더 경로 직접 지정
' ============================================================

Sub ConvertEmlToPDF_Simple()

    ' ★★★ 여기에 폴더 경로 입력 ★★★
    Const INPUT_FOLDER As String = "C:\Mail"
    Const OUTPUT_FOLDER As String = "C:\Mail\PDF"

    Dim fso As Object, wordApp As Object, cdoMsg As Object, adoStream As Object
    Dim file As Object, doc As Object
    Dim pdfPath As String, cnt As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    If Not fso.FolderExists(OUTPUT_FOLDER) Then fso.CreateFolder OUTPUT_FOLDER

    cnt = 0
    For Each file In fso.GetFolder(INPUT_FOLDER).Files
        If LCase(fso.GetExtensionName(file.Name)) = "eml" Then
            On Error Resume Next

            ' EML 파일 로드
            Set cdoMsg = CreateObject("CDO.Message")
            Set adoStream = CreateObject("ADODB.Stream")
            adoStream.Open
            adoStream.LoadFromFile file.Path
            cdoMsg.DataSource.OpenObject adoStream, "_Stream"
            adoStream.Close

            ' Word 문서 생성
            Set doc = wordApp.Documents.Add
            doc.Content.InsertAfter "보낸 사람: " & cdoMsg.From & vbCrLf
            doc.Content.InsertAfter "받는 사람: " & cdoMsg.To & vbCrLf
            doc.Content.InsertAfter "날짜: " & cdoMsg.SentOn & vbCrLf
            doc.Content.InsertAfter "제목: " & cdoMsg.Subject & vbCrLf
            doc.Content.InsertAfter String(50, "-") & vbCrLf & vbCrLf
            doc.Content.InsertAfter cdoMsg.TextBody

            ' PDF 저장
            pdfPath = OUTPUT_FOLDER & "\" & fso.GetBaseName(file.Name) & ".pdf"
            doc.SaveAs2 pdfPath, 17
            doc.Close False

            cnt = cnt + 1
            Set cdoMsg = Nothing
            On Error GoTo 0
        End If
    Next file

    wordApp.Quit
    MsgBox cnt & "개 파일 변환 완료!", vbInformation

End Sub


' ============================================================
' EML 파일 직접 파싱 버전 (CDO 없이 동작)
' ============================================================

Sub ConvertEmlToPDF_NoCDO()

    Dim fso As Object
    Dim wordApp As Object
    Dim inputFolder As String
    Dim outputFolder As String
    Dim file As Object
    Dim doc As Object
    Dim pdfPath As String
    Dim successCount As Long, failCount As Long

    ' 폴더 선택
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "EML 메일 파일이 있는 폴더 선택"
        If .Show = -1 Then
            inputFolder = .SelectedItems(1)
        Else
            MsgBox "폴더가 선택되지 않았습니다.", vbExclamation
            Exit Sub
        End If
    End With

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "PDF 저장할 폴더 선택"
        If .Show = -1 Then
            outputFolder = .SelectedItems(1)
        Else
            outputFolder = inputFolder & "\PDF_Output"
        End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    Application.StatusBar = "변환 중..."
    successCount = 0
    failCount = 0

    For Each file In fso.GetFolder(inputFolder).Files
        If LCase(fso.GetExtensionName(file.Name)) = "eml" Then
            On Error Resume Next

            ' EML 직접 파싱
            Dim emlContent As String
            Dim fromAddr As String, toAddr As String
            Dim subject As String, sentDate As String, body As String

            emlContent = ReadTextFile(file.Path)

            fromAddr = ExtractHeader(emlContent, "From:")
            toAddr = ExtractHeader(emlContent, "To:")
            subject = ExtractHeader(emlContent, "Subject:")
            sentDate = ExtractHeader(emlContent, "Date:")
            body = ExtractBody(emlContent)

            ' PDF 경로
            pdfPath = outputFolder & "\" & CleanFileName(fso.GetBaseName(file.Name)) & ".pdf"
            pdfPath = GetUniquePath(pdfPath, fso)

            ' Word 문서
            Set doc = wordApp.Documents.Add
            With doc.Content
                .InsertAfter "보낸 사람: " & fromAddr & vbCrLf
                .InsertAfter "받는 사람: " & toAddr & vbCrLf
                .InsertAfter "날짜: " & sentDate & vbCrLf
                .InsertAfter "제목: " & subject & vbCrLf
                .InsertAfter String(50, "-") & vbCrLf & vbCrLf
                .InsertAfter body
            End With

            doc.SaveAs2 pdfPath, 17
            doc.Close False

            If Err.Number = 0 Then
                successCount = successCount + 1
            Else
                failCount = failCount + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next file

    wordApp.Quit
    Application.StatusBar = False

    MsgBox "변환 완료!" & vbCrLf & _
           "성공: " & successCount & "개" & vbCrLf & _
           "실패: " & failCount & "개", vbInformation

End Sub


' ============================================================
' 헬퍼 함수들
' ============================================================

Private Function CleanFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim i As Long
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(invalidChars) To UBound(invalidChars)
        fileName = Replace(fileName, invalidChars(i), "_")
    Next i
    fileName = Trim(fileName)
    If Len(fileName) > 200 Then fileName = Left(fileName, 200)
    CleanFileName = fileName
End Function

Private Function GetUniquePath(ByVal filePath As String, ByVal fso As Object) As String
    Dim basePath As String, ext As String, counter As Long
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

Private Function ReadTextFile(ByVal filePath As String) As String
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False)
    ReadTextFile = ts.ReadAll
    ts.Close
End Function

Private Function ExtractHeader(ByVal content As String, ByVal headerName As String) As String
    Dim startPos As Long, endPos As Long
    Dim headerValue As String

    startPos = InStr(1, content, headerName, vbTextCompare)
    If startPos = 0 Then
        ExtractHeader = ""
        Exit Function
    End If

    startPos = startPos + Len(headerName)
    endPos = InStr(startPos, content, vbCrLf)
    If endPos = 0 Then endPos = InStr(startPos, content, vbLf)
    If endPos = 0 Then endPos = Len(content)

    headerValue = Mid(content, startPos, endPos - startPos)
    ExtractHeader = Trim(headerValue)
End Function

Private Function ExtractBody(ByVal content As String) As String
    Dim bodyStart As Long

    ' 헤더와 본문 구분 (빈 줄)
    bodyStart = InStr(content, vbCrLf & vbCrLf)
    If bodyStart = 0 Then bodyStart = InStr(content, vbLf & vbLf)

    If bodyStart > 0 Then
        ExtractBody = Mid(content, bodyStart + 4)
    Else
        ExtractBody = content
    End If

    ExtractBody = Trim(ExtractBody)
End Function
