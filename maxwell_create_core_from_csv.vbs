' ============================================================================
' Maxwell 3D - CSV 기반 변압기 철심 생성 스크립트 (VBScript, Yoke 포함)
' ============================================================================
'
' CSV 파일에서 철심 치수 데이터를 읽어 완전한 1x2 변압기 철심을 생성합니다.
'
' CSV 구조:
' - A2:A(x): X1 값 (Main leg의 X 크기)
' - B2:B(x): X2 값 (Return leg의 X 크기)
' - C2:C(x): Y 값 (Y축 크기)
' - E1: Return leg 이격거리 (Main leg 중점 ↔ Return leg 중점)
' - E2: 철심 창 높이 (Z축 extrusion 높이)
'
' 생성되는 구조:
' - Main leg: 1개 (중앙)
' - Return legs: 2개 (양쪽)
' - Top yoke: 3개 leg를 상단에서 연결
' - Bottom yoke: 3개 leg를 하단에서 연결
' - 모든 박스는 원점(0,0,0)을 중심으로 배치
' - Unite 연산으로 5개 박스를 하나로 통합
'
' 작성자: Claude
' 날짜: 2026-01-07
' ============================================================================

Option Explicit

' ============================================================================
' CSV 파일에서 변압기 철심 생성 (Legs + Yokes)
' ============================================================================
Sub CreateTransformerCoreFromCSV(csvFilePath, materialName, namePrefix)
    Dim oAnsoftApp, oDesktop, oProject, oDesign, oEditor
    Dim fso, file, line, lines, fields
    Dim i, rowIndex
    Dim x1, x2, y, gap, windowHeight
    Dim dataRows()
    Dim rowCount
    Dim layer_boxes, core_name
    Dim mainName, leftName, rightName, topYokeName, bottomYokeName
    Dim mainXPos, mainYPos, mainZPos
    Dim leftXPos, leftYPos, leftZPos
    Dim rightXPos, rightYPos, rightZPos
    Dim yokeXSize, yokeYSize, yokeZSize
    Dim topYokeXPos, topYokeYPos, topYokeZPos
    Dim bottomYokeXPos, bottomYokeYPos, bottomYokeZPos

    ' Desktop 연결
    Set oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")
    Set oDesktop = oAnsoftApp.GetAppDesktop()
    oDesktop.RestoreWindow

    ' CSV 파일 읽기
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(csvFilePath) Then
        MsgBox "CSV 파일을 찾을 수 없습니다: " & csvFilePath, vbCritical, "오류"
        Exit Sub
    End If

    oDesktop.AddMessage "", "", 0, "CSV 기반 변압기 철심 생성 (Legs + Yokes)"
    oDesktop.AddMessage "", "", 0, "CSV 파일 읽기 중: " & csvFilePath

    Set file = fso.OpenTextFile(csvFilePath, 1)  ' 1 = ForReading

    ' 모든 라인 읽기
    ReDim lines(0)
    i = 0
    Do Until file.AtEndOfStream
        line = file.ReadLine
        ReDim Preserve lines(i)
        lines(i) = line
        i = i + 1
    Loop
    file.Close

    ' E1, E2 값 읽기
    gap = 0
    windowHeight = 0

    If UBound(lines) >= 0 Then
        fields = Split(lines(0), ",")
        If UBound(fields) >= 4 Then
            If IsNumeric(fields(4)) Then
                gap = CDbl(fields(4))
            End If
        End If
    End If

    If UBound(lines) >= 1 Then
        fields = Split(lines(1), ",")
        If UBound(fields) >= 4 Then
            If IsNumeric(fields(4)) Then
                windowHeight = CDbl(fields(4))
            End If
        End If
    End If

    oDesktop.AddMessage "", "", 0, "Return leg 이격거리 (E1): " & gap & " mm"
    oDesktop.AddMessage "", "", 0, "철심 창 높이 (E2): " & windowHeight & " mm"

    ' 데이터 행 읽기 (A2:C(x))
    rowCount = 0
    ReDim dataRows(0)

    For i = 1 To UBound(lines)
        fields = Split(lines(i), ",")

        If UBound(fields) >= 2 Then
            If IsNumeric(fields(0)) And IsNumeric(fields(1)) And IsNumeric(fields(2)) Then
                x1 = CDbl(fields(0))
                x2 = CDbl(fields(1))
                y = CDbl(fields(2))

                If x1 > 0 And x2 > 0 And y > 0 Then
                    ReDim Preserve dataRows(rowCount)
                    dataRows(rowCount) = Array(x1, x2, y)
                    rowCount = rowCount + 1
                End If
            End If
        End If
    Next

    If rowCount = 0 Then
        MsgBox "CSV 파일에 유효한 데이터가 없습니다.", vbCritical, "오류"
        Exit Sub
    End If

    oDesktop.AddMessage "", "", 0, "데이터 행 수: " & rowCount

    ' 프로젝트 및 디자인 설정
    Set oProject = oDesktop.GetActiveProject()
    If oProject Is Nothing Then
        Set oProject = oDesktop.NewProject()
        oDesktop.AddMessage "", "", 0, "새 프로젝트를 생성했습니다."
    End If

    Set oDesign = oProject.GetActiveDesign()
    If oDesign Is Nothing Then
        oProject.InsertDesign "Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", ""
        Set oDesign = oProject.GetActiveDesign()
        oDesktop.AddMessage "", "", 0, "새 Maxwell 3D 디자인을 생성했습니다."
    End If

    Set oEditor = oDesign.SetActiveEditor("3D Modeler")

    ' 각 데이터 행마다 완전한 철심 생성 (3 legs + 2 yokes)
    For i = 0 To rowCount - 1
        x1 = dataRows(i)(0)
        x2 = dataRows(i)(1)
        y = dataRows(i)(2)

        oDesktop.AddMessage "", "", 0, "레이어 " & (i + 1) & ": X1=" & x1 & ", X2=" & x2 & ", Y=" & y

        ' ===== 1. Main leg (중앙) =====
        mainName = namePrefix & "_Layer" & (i + 1) & "_Main"
        mainXPos = -x1 / 2.0
        mainYPos = -y / 2.0
        mainZPos = -windowHeight / 2.0

        Call CreateBox(oEditor, mainName, mainXPos, mainYPos, mainZPos, _
                       x1, y, windowHeight, materialName, "(143 175 143)")

        oDesktop.AddMessage "", "", 0, "  Main leg 생성: " & mainName

        ' ===== 2. Left Return leg =====
        leftName = namePrefix & "_Layer" & (i + 1) & "_LeftReturn"
        leftXPos = -gap - x2 / 2.0
        leftYPos = -y / 2.0
        leftZPos = -windowHeight / 2.0

        Call CreateBox(oEditor, leftName, leftXPos, leftYPos, leftZPos, _
                       x2, y, windowHeight, materialName, "(132 132 193)")

        oDesktop.AddMessage "", "", 0, "  Left Return leg 생성: " & leftName

        ' ===== 3. Right Return leg =====
        rightName = namePrefix & "_Layer" & (i + 1) & "_RightReturn"
        rightXPos = gap - x2 / 2.0
        rightYPos = -y / 2.0
        rightZPos = -windowHeight / 2.0

        Call CreateBox(oEditor, rightName, rightXPos, rightYPos, rightZPos, _
                       x2, y, windowHeight, materialName, "(132 132 193)")

        oDesktop.AddMessage "", "", 0, "  Right Return leg 생성: " & rightName

        ' ===== 4. Top Yoke =====
        ' Yoke 치수 계산
        ' Left Return 외곽(-gap - x2/2) ~ Right Return 외곽(gap + x2/2)
        yokeXSize = 2 * gap + x2  ' 정확한 길이 (튀어나오지 않음)
        yokeYSize = x2
        yokeZSize = y

        topYokeName = namePrefix & "_Layer" & (i + 1) & "_TopYoke"
        topYokeXPos = -yokeXSize / 2.0
        topYokeYPos = -yokeYSize / 2.0
        topYokeZPos = windowHeight / 2.0

        Call CreateBox(oEditor, topYokeName, topYokeXPos, topYokeYPos, topYokeZPos, _
                       yokeXSize, yokeYSize, yokeZSize, materialName, "(193 132 132)")

        oDesktop.AddMessage "", "", 0, "  Top yoke 생성: " & topYokeName

        ' ===== 5. Bottom Yoke =====
        bottomYokeName = namePrefix & "_Layer" & (i + 1) & "_BottomYoke"
        bottomYokeXPos = -yokeXSize / 2.0
        bottomYokeYPos = -yokeYSize / 2.0
        bottomYokeZPos = -windowHeight / 2.0 - yokeZSize

        Call CreateBox(oEditor, bottomYokeName, bottomYokeXPos, bottomYokeYPos, bottomYokeZPos, _
                       yokeXSize, yokeYSize, yokeZSize, materialName, "(193 132 132)")

        oDesktop.AddMessage "", "", 0, "  Bottom yoke 생성: " & bottomYokeName

        ' ===== 6. Unite all parts into one core =====
        core_name = namePrefix & "_Layer" & (i + 1)
        oDesktop.AddMessage "", "", 0, "  Unite 연산 중... (5개 박스 -> 1개 철심)"

        ' 박스 이름 리스트 작성
        layer_boxes = mainName & "," & leftName & "," & rightName & "," & topYokeName & "," & bottomYokeName

        ' Unite 연산
        Dim arrUnite1(1), arrUnite2(1)
        arrUnite1(0) = "NAME:Selections"
        arrUnite1(1) = Array("Selections:=", layer_boxes)
        arrUnite2(0) = "NAME:UniteParameters"
        arrUnite2(1) = Array("KeepOriginals:=", False)

        oEditor.Unite arrUnite1, arrUnite2

        ' Unite 후 첫 번째 이름으로 남으므로, 원하는 이름으로 변경
        Dim arrChange1(1), arrChange2(1), arrChange3(1), arrChange4(1)
        arrChange1(0) = "NAME:AllTabs"
        arrChange2(0) = "NAME:Geometry3DAttributeTab"
        arrChange3(0) = "NAME:PropServers"
        arrChange3(1) = Array(mainName)
        arrChange4(0) = "NAME:ChangedProps"
        arrChange4(1) = Array(Array("NAME:Name", "Value:=", core_name))

        arrChange2(1) = Array(arrChange3, arrChange4)
        arrChange1(1) = Array(arrChange2)

        oEditor.ChangeProperty arrChange1

        oDesktop.AddMessage "", "", 0, "  완료! 통합된 철심: " & core_name
    Next

    ' 뷰 맞추기
    oEditor.FitAll

    MsgBox "완료! 총 " & rowCount & " 개의 철심 레이어가 생성되었습니다.", vbInformation, "완료"
    oDesktop.AddMessage "", "", 0, "완료! 총 " & rowCount & " 개의 철심 레이어가 생성되었습니다."
End Sub


' ============================================================================
' 박스 생성 헬퍼 함수
' ============================================================================
Sub CreateBox(oEditor, boxName, xPos, yPos, zPos, xSize, ySize, zSize, materialName, colorRGB)
    Dim arrBoxParams(1), arrAttributes(1)

    arrBoxParams(0) = "NAME:BoxParameters"
    arrBoxParams(1) = Array(_
        "XPosition:=", CStr(xPos) & "mm", _
        "YPosition:=", CStr(yPos) & "mm", _
        "ZPosition:=", CStr(zPos) & "mm", _
        "XSize:=", CStr(xSize) & "mm", _
        "YSize:=", CStr(ySize) & "mm", _
        "ZSize:=", CStr(zSize) & "mm" _
    )

    arrAttributes(0) = "NAME:Attributes"
    arrAttributes(1) = Array(_
        "Name:=", boxName, _
        "Flags:=", "", _
        "Color:=", colorRGB, _
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

    oEditor.CreateBox arrBoxParams, arrAttributes
End Sub


' ============================================================================
' 메인 실행 부분
' ============================================================================

' CSV 파일 경로 설정 (여기를 수정하세요!)
Dim csvFile
csvFile = "C:\path\to\your\transformer_core_data.csv"

' 재질 설정
Dim coreMaterial
coreMaterial = "steel_1008"  ' Maxwell의 재질 라이브러리에 있는 재질 이름

' 파일 선택 대화상자 표시 (옵션)
' csvFile = SelectFile()  ' 이 줄의 주석을 해제하면 파일 선택 대화상자가 나타남

' 변압기 철심 생성
Call CreateTransformerCoreFromCSV(csvFile, coreMaterial, "Core")


' ============================================================================
' 파일 선택 대화상자 (옵션)
' ============================================================================
Function SelectFile()
    Dim shell, file
    Set shell = CreateObject("WScript.Shell")

    ' 간단한 InputBox로 파일 경로 입력받기
    file = InputBox("CSV 파일 경로를 입력하세요:", "CSV 파일 선택", _
                    "C:\transformer_core_data.csv")

    SelectFile = file
End Function
