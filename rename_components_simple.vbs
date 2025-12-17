' ==============================================================================
' Simcenter Magnet - Simple Component Rename Script
' ==============================================================================
'
' Quick script to rename multiple components with a prefix
' Press Cancel or ESC to exit
'
' ==============================================================================

' Get prefix from user
Dim prefix
prefix = InputBox("접두사(prefix) 입력:", "Prefix", "Part")
If prefix = "" Then
    Call MsgBox("취소되었습니다.", vbInformation)
    WScript.Quit
End If

' Initialize
Dim oDoc, oView, selectedComp, componentName, newName, count
Set oDoc = GetDocument()
Set oView = oDoc.GetView()
count = 0

' Instructions
Call MsgBox("컴포넌트를 선택하고 이름을 입력하세요." & vbCrLf & _
            "형식: " & prefix & "_[이름]" & vbCrLf & vbCrLf & _
            "종료하려면 취소 또는 ESC를 누르세요.", vbInformation)

' Main loop
Do
    ' Select component
    On Error Resume Next
    Set selectedComp = oView.SelectComponent("컴포넌트 선택 (취소=종료)")

    If Err.Number <> 0 Or selectedComp Is Nothing Then
        Exit Do
    End If
    On Error GoTo 0

    ' Get name
    componentName = InputBox("이름 입력:", "새 이름", "")
    If componentName = "" Then Exit Do

    ' Rename
    newName = prefix & "_" & componentName
    Call selectedComp.SetName(newName)
    count = count + 1
Loop

' Summary
Call MsgBox("완료! 변경된 컴포넌트: " & count & "개", vbInformation, "완료")
