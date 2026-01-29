' ============================================
' 파일명 추출 및 코드 분리 VBA 모듈
' ATL, AL, TL, TP + 숫자코드 자동 분리
' ============================================

Option Explicit

Sub ExtractFilenames()
    Dim folderPath As String
    Dim fileName As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim ws As Worksheet
    Dim row As Long
    Dim extractedCode As String
    Dim codeType As String
    Dim codeNumber As String

    ' 폴더 경로 입력 받기
    folderPath = InputBox("파일명을 추출할 폴더 경로를 입력하세요:" & vbCrLf & vbCrLf & _
                          "예: C:\Users\Documents\MyFolder", "폴더 경로 입력")

    ' 취소 또는 빈 입력 처리
    If folderPath = "" Then
        MsgBox "폴더 경로가 입력되지 않았습니다.", vbExclamation, "알림"
        Exit Sub
    End If

    ' 경로 끝에 백슬래시 추가
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' FileSystemObject 생성
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 폴더 존재 확인
    If Not fso.FolderExists(folderPath) Then
        MsgBox "입력한 폴더가 존재하지 않습니다:" & vbCrLf & folderPath, vbCritical, "오류"
        Exit Sub
    End If

    Set folder = fso.GetFolder(folderPath)

    ' 새 워크시트 생성 또는 기존 시트 사용
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("파일목록")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "파일목록"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ' 헤더 작성
    With ws
        .Range("A1").Value = "번호"
        .Range("B1").Value = "파일명 (전체)"
        .Range("C1").Value = "코드 타입"
        .Range("D1").Value = "코드 번호"
        .Range("E1").Value = "추출된 코드"

        ' 헤더 서식
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(200, 200, 200)
    End With

    row = 2

    ' 폴더 내 모든 파일 순회
    For Each file In folder.Files
        fileName = file.Name

        ' 코드 추출
        Call ExtractCode(fileName, codeType, codeNumber, extractedCode)

        ' 데이터 입력
        ws.Cells(row, 1).Value = row - 1           ' 번호
        ws.Cells(row, 2).Value = fileName          ' 전체 파일명
        ws.Cells(row, 3).Value = codeType          ' 코드 타입 (ATL, AL, TL, TP)
        ws.Cells(row, 4).Value = codeNumber        ' 숫자 코드
        ws.Cells(row, 5).Value = extractedCode     ' 전체 추출 코드

        row = row + 1
    Next file

    ' 열 너비 자동 조정
    ws.Columns("A:E").AutoFit

    ' 결과 메시지
    MsgBox "총 " & (row - 2) & "개의 파일명이 추출되었습니다.", vbInformation, "완료"

    ' 메모리 해제
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    Set ws = Nothing
End Sub

' 코드 추출 함수
Private Sub ExtractCode(ByVal fileName As String, ByRef codeType As String, _
                        ByRef codeNumber As String, ByRef extractedCode As String)
    Dim regex As Object
    Dim matches As Object
    Dim pattern As String

    ' 초기화
    codeType = ""
    codeNumber = ""
    extractedCode = ""

    ' 정규식 객체 생성
    Set regex = CreateObject("VBScript.RegExp")

    ' 패턴: ATL, AL, TL, TP 뒤에 하이픈(선택) + 숫자들
    ' 대소문자 구분 없이 검색
    pattern = "(ATL|AL|TL|TP)[-]?(\d+)"

    With regex
        .Global = False
        .IgnoreCase = True
        .pattern = pattern
    End With

    ' 매칭 확인
    If regex.Test(fileName) Then
        Set matches = regex.Execute(fileName)

        If matches.Count > 0 Then
            codeType = UCase(matches(0).SubMatches(0))      ' ATL, AL, TL, TP
            codeNumber = matches(0).SubMatches(1)            ' 숫자 부분
            extractedCode = matches(0).Value                 ' 전체 코드 (예: ATL1234, TP-4335)
        End If
    End If

    Set matches = Nothing
    Set regex = Nothing
End Sub

' 하위 폴더 포함 버전
Sub ExtractFilenamesWithSubfolders()
    Dim folderPath As String
    Dim fso As Object
    Dim ws As Worksheet
    Dim row As Long

    ' 폴더 경로 입력 받기
    folderPath = InputBox("파일명을 추출할 폴더 경로를 입력하세요:" & vbCrLf & _
                          "(하위 폴더 포함)" & vbCrLf & vbCrLf & _
                          "예: C:\Users\Documents\MyFolder", "폴더 경로 입력")

    If folderPath = "" Then
        MsgBox "폴더 경로가 입력되지 않았습니다.", vbExclamation, "알림"
        Exit Sub
    End If

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        MsgBox "입력한 폴더가 존재하지 않습니다.", vbCritical, "오류"
        Exit Sub
    End If

    ' 워크시트 준비
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("파일목록")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "파일목록"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0

    ' 헤더
    With ws
        .Range("A1").Value = "번호"
        .Range("B1").Value = "파일명 (전체)"
        .Range("C1").Value = "코드 타입"
        .Range("D1").Value = "코드 번호"
        .Range("E1").Value = "추출된 코드"
        .Range("F1").Value = "폴더 경로"
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(200, 200, 200)
    End With

    row = 2

    ' 재귀적으로 파일 추출
    Call ProcessFolder(fso.GetFolder(folderPath), ws, row)

    ws.Columns("A:F").AutoFit

    MsgBox "총 " & (row - 2) & "개의 파일명이 추출되었습니다.", vbInformation, "완료"

    Set fso = Nothing
    Set ws = Nothing
End Sub

' 재귀 폴더 처리 함수
Private Sub ProcessFolder(ByVal folder As Object, ByVal ws As Worksheet, ByRef row As Long)
    Dim file As Object
    Dim subfolder As Object
    Dim codeType As String
    Dim codeNumber As String
    Dim extractedCode As String

    ' 현재 폴더의 파일 처리
    For Each file In folder.Files
        Call ExtractCode(file.Name, codeType, codeNumber, extractedCode)

        ws.Cells(row, 1).Value = row - 1
        ws.Cells(row, 2).Value = file.Name
        ws.Cells(row, 3).Value = codeType
        ws.Cells(row, 4).Value = codeNumber
        ws.Cells(row, 5).Value = extractedCode
        ws.Cells(row, 6).Value = folder.Path

        row = row + 1
    Next file

    ' 하위 폴더 재귀 처리
    For Each subfolder In folder.SubFolders
        Call ProcessFolder(subfolder, ws, row)
    Next subfolder
End Sub
