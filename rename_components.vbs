' ==============================================================================
' Simcenter Magnet - Interactive Component Rename Script
' ==============================================================================
'
' This script allows interactive renaming of multiple components with a prefix
'
' Workflow:
' 1. Enter prefix when script starts
' 2. Select a component in the model
' 3. Enter name for the component
' 4. Component is renamed to "prefix_name"
' 5. Repeat for next component
' 6. Press ESC or Cancel to exit
'
' Author: Claude
' Date: 2025-12-17
' ==============================================================================

Option Explicit

' ==============================================================================
' Main Subroutine - Interactive Component Rename
' ==============================================================================
Sub Main()
    Dim oDoc, oView
    Dim prefix
    Dim selectedComp
    Dim componentName
    Dim newName
    Dim continueLoop
    Dim count

    ' Get the active document
    Set oDoc = GetDocument()
    Set oView = oDoc.GetView()

    ' Get prefix from user
    prefix = InputBox("컴포넌트 이름에 사용할 접두사(prefix)를 입력하세요:" & vbCrLf & vbCrLf & _
                      "예: 'Part'를 입력하고 이름을 'Core'로 지정하면 'Part_Core'가 됩니다.", _
                      "Prefix 입력", "Component")

    ' Check if user cancelled
    If prefix = "" Then
        Call MsgBox("취소되었습니다.", vbInformation, "알림")
        Exit Sub
    End If

    ' Initialize counter
    count = 0
    continueLoop = True

    ' Show instructions
    Call MsgBox("이제 이름을 변경할 컴포넌트를 선택하세요." & vbCrLf & vbCrLf & _
                "각 컴포넌트 선택 후 이름을 입력하면 '" & prefix & "_[이름]' 형식으로 변경됩니다." & vbCrLf & vbCrLf & _
                "작업을 마치려면 컴포넌트 선택 또는 이름 입력 시 '취소'를 누르세요.", _
                vbInformation, "사용 방법")

    ' Main loop - continues until user cancels
    Do While continueLoop
        ' Prompt user to select a component
        On Error Resume Next
        Set selectedComp = oView.SelectComponent("이름을 변경할 컴포넌트를 선택하세요 (ESC 또는 취소하면 종료)")

        ' Check if selection was cancelled or failed
        If Err.Number <> 0 Or selectedComp Is Nothing Then
            continueLoop = False
            On Error GoTo 0
        Else
            On Error GoTo 0

            ' Get current component name
            Dim currentName
            currentName = selectedComp.GetName()

            ' Ask for new name
            componentName = InputBox("컴포넌트의 이름을 입력하세요:" & vbCrLf & vbCrLf & _
                                    "현재 이름: " & currentName & vbCrLf & _
                                    "새 이름: " & prefix & "_[입력한 이름]", _
                                    "이름 입력", "")

            ' Check if user cancelled
            If componentName = "" Then
                continueLoop = False
            Else
                ' Create new name with prefix
                newName = prefix & "_" & componentName

                ' Rename the component
                Call selectedComp.SetName(newName)

                ' Increment counter
                count = count + 1

                ' Show confirmation (optional - can be removed for faster workflow)
                ' Call MsgBox("이름 변경 완료: " & newName, vbInformation, "확인")
            End If
        End If
    Loop

    ' Show summary
    If count > 0 Then
        Call MsgBox("작업 완료!" & vbCrLf & vbCrLf & _
                    "변경된 컴포넌트 수: " & count, _
                    vbInformation, "완료")
    Else
        Call MsgBox("변경된 컴포넌트가 없습니다.", vbInformation, "알림")
    End If

    ' Clean up
    Set selectedComp = Nothing
    Set oView = Nothing
    Set oDoc = Nothing
End Sub

' ==============================================================================
' Alternative Version - Without Confirmation Messages
' ==============================================================================
Sub RenameComponentsFast()
    Dim oDoc, oView
    Dim prefix
    Dim selectedComp
    Dim componentName
    Dim newName
    Dim count

    Set oDoc = GetDocument()
    Set oView = oDoc.GetView()

    ' Get prefix
    prefix = InputBox("접두사(prefix) 입력:", "Prefix", "Part")
    If prefix = "" Then Exit Sub

    count = 0

    ' Loop without confirmation messages for faster workflow
    Do
        On Error Resume Next
        Set selectedComp = oView.SelectComponent("컴포넌트 선택 (취소=종료)")

        If Err.Number <> 0 Or selectedComp Is Nothing Then
            Exit Do
        End If
        On Error GoTo 0

        componentName = InputBox("이름 입력 (" & prefix & "_???):", "이름", "")
        If componentName = "" Then Exit Do

        newName = prefix & "_" & componentName
        Call selectedComp.SetName(newName)
        count = count + 1
    Loop

    Call MsgBox("완료! 변경된 컴포넌트: " & count & "개", vbInformation, "완료")

    Set selectedComp = Nothing
    Set oView = Nothing
    Set oDoc = Nothing
End Sub

' ==============================================================================
' Advanced Version - With Name Validation and Preview
' ==============================================================================
Sub RenameComponentsAdvanced()
    Dim oDoc, oView
    Dim prefix, separator
    Dim selectedComp
    Dim componentName, currentName, newName
    Dim count
    Dim continueRename

    Set oDoc = GetDocument()
    Set oView = oDoc.GetView()

    ' Get prefix and separator
    prefix = InputBox("접두사(prefix)를 입력하세요:", "Prefix 입력", "Component")
    If prefix = "" Then Exit Sub

    separator = InputBox("구분자를 입력하세요:" & vbCrLf & _
                        "(기본값: '_' 또는 '-', '.' 등 사용 가능)", _
                        "구분자 입력", "_")
    If separator = "" Then separator = "_"

    Call MsgBox("컴포넌트 선택을 시작합니다." & vbCrLf & vbCrLf & _
                "형식: " & prefix & separator & "[이름]" & vbCrLf & vbCrLf & _
                "종료하려면 선택 또는 입력 시 취소하세요.", _
                vbInformation, "시작")

    count = 0

    Do
        On Error Resume Next
        Set selectedComp = oView.SelectComponent("이름 변경할 컴포넌트 선택 (취소=종료)")

        If Err.Number <> 0 Or selectedComp Is Nothing Then
            Exit Do
        End If
        On Error GoTo 0

        currentName = selectedComp.GetName()

        ' Show preview in input box
        componentName = InputBox("새 이름을 입력하세요:" & vbCrLf & vbCrLf & _
                                "현재: " & currentName & vbCrLf & _
                                "미리보기: " & prefix & separator & "[입력한 이름]", _
                                "이름 입력", "")

        If componentName = "" Then
            ' Ask if user wants to continue or exit
            continueRename = MsgBox("계속 다른 컴포넌트 이름을 변경하시겠습니까?", _
                                   vbYesNo + vbQuestion, "계속?")
            If continueRename = vbNo Then
                Exit Do
            End If
        Else
            ' Validate name (no special characters)
            If InStr(componentName, "/") > 0 Or InStr(componentName, "\") > 0 Or _
               InStr(componentName, ":") > 0 Or InStr(componentName, "*") > 0 Then
                Call MsgBox("이름에 특수문자(/ \ : *)를 사용할 수 없습니다.", _
                           vbExclamation, "오류")
            Else
                newName = prefix & separator & componentName
                Call selectedComp.SetName(newName)
                count = count + 1
            End If
        End If
    Loop

    Call MsgBox("작업 완료!" & vbCrLf & _
                "변경된 컴포넌트: " & count & "개", _
                vbInformation, "완료")

    Set selectedComp = Nothing
    Set oView = Nothing
    Set oDoc = Nothing
End Sub

' ==============================================================================
' Run the main subroutine
' ==============================================================================
' 아래 중 하나를 선택하여 실행하세요:

Call Main()                      ' 기본 버전 (상세한 안내 메시지)
' Call RenameComponentsFast()    ' 빠른 버전 (확인 메시지 최소화)
' Call RenameComponentsAdvanced() ' 고급 버전 (이름 검증 및 미리보기)
