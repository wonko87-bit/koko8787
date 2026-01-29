' ============================================================
' Outlook 메일 PDF 변환 - 간단 버전
' 아래 폴더 경로를 직접 수정하여 사용
' ============================================================

Option Explicit

' ★★★ 여기에 폴더 경로 입력 ★★★
Const INPUT_FOLDER = "C:\Mail"          ' .msg 파일이 있는 폴더
Const OUTPUT_FOLDER = "C:\Mail\PDF"     ' PDF 저장할 폴더

Dim fso, outlook, word, ns
Dim file, mail, doc, pdfPath

Set fso = CreateObject("Scripting.FileSystemObject")
Set outlook = CreateObject("Outlook.Application")
Set word = CreateObject("Word.Application")
Set ns = outlook.GetNamespace("MAPI")

word.Visible = False

' 출력 폴더 생성
If Not fso.FolderExists(OUTPUT_FOLDER) Then
    fso.CreateFolder(OUTPUT_FOLDER)
End If

' .msg 파일 처리
For Each file In fso.GetFolder(INPUT_FOLDER).Files
    If LCase(fso.GetExtensionName(file.Name)) = "msg" Then
        On Error Resume Next

        ' 메일 열기
        Set mail = ns.OpenSharedItem(file.Path)

        ' Word 문서 생성
        Set doc = word.Documents.Add

        ' 메일 정보 입력
        doc.Content.InsertAfter "보낸 사람: " & mail.SenderName & vbCrLf
        doc.Content.InsertAfter "받는 사람: " & mail.To & vbCrLf
        doc.Content.InsertAfter "날짜: " & mail.ReceivedTime & vbCrLf
        doc.Content.InsertAfter "제목: " & mail.Subject & vbCrLf
        doc.Content.InsertAfter String(50, "-") & vbCrLf & vbCrLf
        doc.Content.InsertAfter mail.Body

        ' PDF 저장
        pdfPath = OUTPUT_FOLDER & "\" & fso.GetBaseName(file.Name) & ".pdf"
        doc.SaveAs2 pdfPath, 17  ' 17 = wdFormatPDF

        doc.Close False
        mail.Close 0

        WScript.Echo "완료: " & fso.GetBaseName(file.Name)
        On Error GoTo 0
    End If
Next

word.Quit
MsgBox "변환 완료!", vbInformation

Set doc = Nothing
Set mail = Nothing
Set word = Nothing
Set outlook = Nothing
Set fso = Nothing
