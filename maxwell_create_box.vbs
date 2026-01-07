' ============================================================================
' Ansys Maxwell Electronics Desktop 2021 R1 - 3D Box Modeling Script (VBScript)
' ============================================================================
'
' 이 스크립트는 Maxwell 3D에서 직육면체를 생성합니다.
' 박스는 원점 (0,0,0)을 한 꼭지점으로 하여 양의 방향으로 확장됩니다.
'
' 작성자: Claude
' 날짜: 2026-01-07
'
' 사용 방법:
' ---------
' 1. Ansys Maxwell Electronics Desktop을 실행합니다
' 2. Tools → Run Script에서 이 파일을 실행합니다
'
' 또는 스크립트 콘솔에서 직접 실행 가능합니다.
' ============================================================================

Option Explicit

' ============================================================================
' 메인 함수: Maxwell 3D 박스 생성
' ============================================================================
Sub CreateMaxwellBox(width, depth, height, boxName, materialName)
    ' 매개변수:
    '   width        - X 방향 너비 (mm)
    '   depth        - Y 방향 깊이 (mm)
    '   height       - Z 방향 높이 (mm)
    '   boxName      - 박스 이름
    '   materialName - 재질 이름

    Dim oAnsoftApp, oDesktop, oProject, oDesign, oEditor

    ' 유효성 검사
    If width <= 0 Or depth <= 0 Or height <= 0 Then
        MsgBox "오류: 모든 치수는 양수여야 합니다!", vbCritical, "입력 오류"
        Exit Sub
    End If

    ' 정보 출력
    Dim msg
    msg = "Maxwell 3D - 박스 생성 중..." & vbCrLf & vbCrLf & _
          "치수:" & vbCrLf & _
          "  너비 (X): " & width & " mm" & vbCrLf & _
          "  깊이 (Y): " & depth & " mm" & vbCrLf & _
          "  높이 (Z): " & height & " mm" & vbCrLf & _
          "  이름: " & boxName & vbCrLf & _
          "  재질: " & materialName

    AddMessage msg

    On Error Resume Next

    ' Desktop 객체 가져오기
    Set oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")
    Set oDesktop = oAnsoftApp.GetAppDesktop()
    oDesktop.RestoreWindow

    ' 활성 프로젝트 가져오기 (없으면 새로 생성)
    Set oProject = oDesktop.GetActiveProject()
    If oProject Is Nothing Then
        Set oProject = oDesktop.NewProject()
        AddMessage "새 프로젝트를 생성했습니다."
    End If

    ' 활성 디자인 가져오기
    Set oDesign = oProject.GetActiveDesign()

    ' Maxwell 3D 디자인이 없으면 새로 생성
    If oDesign Is Nothing Then
        oProject.InsertDesign "Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", ""
        Set oDesign = oProject.GetActiveDesign()
        AddMessage "새 Maxwell 3D 디자인을 생성했습니다."
    End If

    ' 3D Modeler 에디터 가져오기
    Set oEditor = oDesign.SetActiveEditor("3D Modeler")

    ' 박스 생성 (Maxwell 2021 R1 API에 맞춤)
    Dim arrBoxParams(1), arrAttributes(1)

    ' 박스 매개변수 배열
    arrBoxParams(0) = "NAME:BoxParameters"
    arrBoxParams(1) = Array(_
        "XPosition:=", "0mm", _
        "YPosition:=", "0mm", _
        "ZPosition:=", "0mm", _
        "XSize:=", CStr(width) & "mm", _
        "YSize:=", CStr(depth) & "mm", _
        "ZSize:=", CStr(height) & "mm" _
    )

    ' 속성 배열
    arrAttributes(0) = "NAME:Attributes"
    arrAttributes(1) = Array(_
        "Name:=", boxName, _
        "Flags:=", "", _
        "Color:=", "(143 175 143)", _
        "Transparency:=", 0, _
        "PartCoordinateSystem:=", "Global", _
        "UDMId:=", "", _
        "MaterialValue:=", Chr(34) & materialName & Chr(34), _
        "SurfaceMaterialValue:=", Chr(34) & Chr(34), _
        "SolveInside:=", True, _
        "ShellElement:=", False, _
        "ShellElementThickness:=", "0mm", _
        "IsMaterialEditable:=", True, _
        "UseMaterialAppearance:=", False, _
        "IsLightweight:=", False _
    )

    ' 박스 생성 명령 실행
    oEditor.CreateBox arrBoxParams, arrAttributes

    If Err.Number <> 0 Then
        MsgBox "오류 발생: " & Err.Description, vbCritical, "스크립트 오류"
        Exit Sub
    End If

    ' 뷰 맞추기
    oEditor.FitAll

    AddMessage "성공! 박스 '" & boxName & "' 가 생성되었습니다."

    MsgBox "박스가 성공적으로 생성되었습니다!" & vbCrLf & vbCrLf & _
           "이름: " & boxName & vbCrLf & _
           "크기: " & width & " x " & depth & " x " & height & " mm", _
           vbInformation, "완료"

End Sub


' ============================================================================
' 간단한 박스 생성 함수
' ============================================================================
Sub CreateBoxSimple()
    ' 0,0,0 위치에서 20x10x5 크기의 박스 생성
    Call CreateMaxwellBox(20, 10, 5, "Box_20x10x5", "vacuum")
End Sub


' ============================================================================
' 사용자 입력을 받아 박스를 생성하는 함수
' ============================================================================
Sub CreateBoxInteractive()
    Dim width, depth, height, boxName, materialName
    Dim userInput

    ' 너비 입력
    userInput = InputBox("박스의 너비(X)를 입력하세요 (mm):", "너비 입력", "20")
    If userInput = "" Then Exit Sub
    width = CDbl(userInput)

    ' 깊이 입력
    userInput = InputBox("박스의 깊이(Y)를 입력하세요 (mm):", "깊이 입력", "10")
    If userInput = "" Then Exit Sub
    depth = CDbl(userInput)

    ' 높이 입력
    userInput = InputBox("박스의 높이(Z)를 입력하세요 (mm):", "높이 입력", "5")
    If userInput = "" Then Exit Sub
    height = CDbl(userInput)

    ' 이름 입력
    boxName = InputBox("박스의 이름을 입력하세요:", "이름 입력", "MyBox")
    If boxName = "" Then boxName = "Box1"

    ' 재질 입력
    materialName = InputBox("재질 이름을 입력하세요:" & vbCrLf & vbCrLf & _
                           "일반적인 재질:" & vbCrLf & _
                           "  - vacuum (진공)" & vbCrLf & _
                           "  - aluminum (알루미늄)" & vbCrLf & _
                           "  - copper (구리)" & vbCrLf & _
                           "  - steel (강철)" & vbCrLf & _
                           "  - iron (철)", _
                           "재질 입력", "vacuum")
    If materialName = "" Then materialName = "vacuum"

    ' 박스 생성
    Call CreateMaxwellBox(width, depth, height, boxName, materialName)
End Sub


' ============================================================================
' 여러 개의 박스를 생성하는 예제
' ============================================================================
Sub CreateMultipleBoxes()
    AddMessage "여러 박스를 생성합니다..."

    ' 작은 박스
    Call CreateMaxwellBox(20, 10, 5, "SmallBox", "vacuum")

    ' 정육면체
    Call CreateMaxwellBox(50, 50, 50, "Cube", "aluminum")

    ' 큰 박스
    Call CreateMaxwellBox(100, 30, 15, "LargeBox", "copper")

    MsgBox "3개의 박스가 생성되었습니다!", vbInformation, "완료"
End Sub


' ============================================================================
' 헬퍼 함수: 메시지 출력
' ============================================================================
Sub AddMessage(msg)
    ' 디버깅 및 정보 출력용
    ' Maxwell의 메시지 창에 출력됩니다
    On Error Resume Next
    Dim oDesktop
    Set oDesktop = CreateObject("Ansoft.ElectronicsDesktop").GetAppDesktop()
    If Not oDesktop Is Nothing Then
        oDesktop.AddMessage "", "", 0, msg
    End If
End Sub


' ============================================================================
' 메인 실행 부분
' ============================================================================

' 사용하려는 함수의 주석을 해제하세요:

' 방법 1: 기본 박스 생성 (20x10x5)
Call CreateBoxSimple()

' 방법 2: 대화형 입력으로 박스 생성
' Call CreateBoxInteractive()

' 방법 3: 사용자 정의 박스 생성
' Call CreateMaxwellBox(100, 50, 30, "CustomBox", "aluminum")

' 방법 4: 여러 박스 생성
' Call CreateMultipleBoxes()
