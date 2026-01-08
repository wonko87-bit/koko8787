' Maxwell 3D - CSV 기반 변압기 철심 생성 스크립트 V2 (Rectangle + Sweep 방식)
' ============================================================================
'
' Rectangle을 먼저 그리고 Move, Rotate 후 Unite하여 Sweep하는 고속 방식
'
' CSV 구조:
' - A2:A(x): X1 값 (Main leg의 X 크기)
' - B2:B(x): X2 값 (Return leg의 X 크기)
' - C2:C(x): Y 값 (Y축 크기)
' - E1: Return leg 이격거리 (Main leg 중점 ↔ Return leg 중점)
' - E2: 철심 창 높이 (Z축 extrusion 높이)
'
' Author: Claude
' Date: 2026-01-07

Option Explicit

Dim oAnsoftApp
Dim oDesktop
Dim oProject
Dim oDesign
Dim oEditor

' ANSYS Electronics Desktop 연결
Set oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")
Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow

' CSV 파일 경로 (스크립트와 동일한 폴더)
Dim fso, scriptFolder, csvFile
Set fso = CreateObject("Scripting.FileSystemObject")
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
csvFile = fso.BuildPath(scriptFolder, "transformer_core_sample.csv")

' 메인 함수 실행
CreateTransformerCoreFromCSV csvFile, "steel_1008", "Core"

MsgBox "스크립트 실행이 완료되었습니다!", vbInformation


'==============================================================================
' 서브루틴: CSV 파일에서 변압기 철심 생성 (Rectangle + Sweep 방식)
'==============================================================================
Sub CreateTransformerCoreFromCSV(csvFilePath, materialName, namePrefix)
    Dim fso, file, line, lines(), i, j
    Dim parts, x1, x2, y, gap, windowHeight
    Dim dataRows(), rowCount
    Dim arrRects(), rectCount
    Dim mainName, leftSideName, rightSideName, topYokeName, bottomYokeName
    Dim coreName
    Dim yokeXSize, yokeYSize
    Dim zStart
    Dim dx, dy, dz

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' CSV 파일 읽기
    If Not fso.FileExists(csvFilePath) Then
        MsgBox "CSV 파일을 찾을 수 없습니다: " & csvFilePath, vbCritical
        Exit Sub
    End If

    Set file = fso.OpenTextFile(csvFilePath, 1)
    Dim content
    content = file.ReadAll
    file.Close

    lines = Split(content, vbCrLf)
    If UBound(lines) < 0 Then
        lines = Split(content, vbLf)
    End If

    ' E1 (gap) 및 E2 (window_height) 읽기
    If UBound(lines) >= 0 Then
        parts = Split(lines(0), ",")
        If UBound(parts) >= 4 Then
            gap = CDbl(parts(4))
        Else
            MsgBox "E1 (gap) 값이 없습니다.", vbCritical
            Exit Sub
        End If
    End If

    If UBound(lines) >= 1 Then
        parts = Split(lines(1), ",")
        If UBound(parts) >= 4 Then
            windowHeight = CDbl(parts(4))
        Else
            MsgBox "E2 (window_height) 값이 없습니다.", vbCritical
            Exit Sub
        End If
    End If

    ' 데이터 행 읽기
    rowCount = 0
    ReDim dataRows(100, 2)  ' 최대 100행

    For i = 0 To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            parts = Split(lines(i), ",")
            If UBound(parts) >= 2 Then
                On Error Resume Next
                x1 = CDbl(parts(0))
                x2 = CDbl(parts(1))
                y = CDbl(parts(2))
                If Err.Number = 0 Then
                    dataRows(rowCount, 0) = x1
                    dataRows(rowCount, 1) = x2
                    dataRows(rowCount, 2) = y
                    rowCount = rowCount + 1
                End If
                On Error GoTo 0
            End If
        End If
    Next

    If rowCount = 0 Then
        MsgBox "CSV 파일에서 유효한 데이터를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If

    oDesktop.AddMessage "", "", 0, "======================================"
    oDesktop.AddMessage "", "", 0, "변압기 철심 생성 (Rectangle+Sweep 방식)"
    oDesktop.AddMessage "", "", 0, "======================================"
    oDesktop.AddMessage "", "", 0, "Return leg 이격거리 (E1): " & gap & "mm"
    oDesktop.AddMessage "", "", 0, "철심 창 높이 (E2): " & windowHeight & "mm"
    oDesktop.AddMessage "", "", 0, "데이터 행 수: " & rowCount

    ' Maxwell 프로젝트 및 디자인 가져오기
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

    ' 화면 업데이트 일시 중지 (성능 향상)
    oEditor.SuspendUpdate
    oDesktop.AddMessage "", "", 0, "화면 업데이트를 일시 중지했습니다. (성능 향상 모드)"

    ' Z 시작 위치 계산 (원점 중심)
    zStart = -windowHeight / 2.0

    ' Y 누적 위치 추적 (적층용)
    Dim cumulativeY, prevY
    cumulativeY = 0.0
    prevY = 0.0

    ' 각 데이터 행마다 완전한 철심 생성
    For i = 0 To rowCount - 1
        x1 = dataRows(i, 0)
        x2 = dataRows(i, 1)
        y = dataRows(i, 2)

        oDesktop.AddMessage "", "", 0, "레이어 " & (i + 1) & ": X1=" & x1 & ", X2=" & x2 & ", Y=" & y

        ' 적층 Y 좌표 계산 (이전 레이어의 Y/2 + 현재 레이어의 Y/2)
        If i = 0 Then
            cumulativeY = 0.0  ' 첫 번째 레이어는 원점
        Else
            cumulativeY = cumulativeY + prevY / 2.0 + y / 2.0
        End If

        oDesktop.AddMessage "", "", 0, "  적층 Y 위치: " & cumulativeY & "mm"

        ReDim arrRects(4)
        rectCount = 0

        ' ===== 1. Main leg (중앙) =====
        mainName = namePrefix & "_Layer" & (i + 1) & "_Main"

        ' Rectangle 생성
        oEditor.CreateRectangle _
            Array("NAME:RectangleParameters", _
                "IsCovered:=", True, _
                "XStart:=", "0mm", _
                "YStart:=", "0mm", _
                "ZStart:=", zStart & "mm", _
                "Width:=", x1 & "mm", _
                "Height:=", y & "mm", _
                "WhichAxis:=", "Z"), _
            Array("NAME:Attributes", _
                "Name:=", mainName, _
                "Flags:=", "", _
                "Color:=", "(132 132 193)", _
                "Transparency:=", 0.4, _
                "PartCoordinateSystem:=", "Global", _
                "UDMId:=", "", _
                "MaterialValue:=", """vacuum""", _
                "SurfaceMaterialValue:=", """""", _
                "SolveInside:=", True, _
                "ShellElement:=", False, _
                "ShellElementThickness:=", "0mm", _
                "IsMaterialEditable:=", True, _
                "UseMaterialAppearance:=", False, _
                "IsLightweight:=", False)

        ' 중심으로 이동 (적층 위치 고려)
        dx = -x1 / 2.0
        dy = cumulativeY - y / 2.0
        oEditor.Move _
            Array("NAME:Selections", _
                "Selections:=", mainName, _
                "NewPartsModelFlag:=", "Model"), _
            Array("NAME:TranslateParameters", _
                "TranslateVectorX:=", dx & "mm", _
                "TranslateVectorY:=", dy & "mm", _
                "TranslateVectorZ:=", "0mm")

        arrRects(rectCount) = mainName
        rectCount = rectCount + 1

        ' ===== 2. Left Side leg (X,Y 좌표 교환으로 자동 회전) =====
        leftSideName = namePrefix & "_Layer" & (i + 1) & "_LeftSide"

        ' Y x X2 rectangle 생성 (X, Y 좌표 바꿈)
        oEditor.CreateRectangle _
            Array("NAME:RectangleParameters", _
                "IsCovered:=", True, _
                "XStart:=", "0mm", _
                "YStart:=", "0mm", _
                "ZStart:=", zStart & "mm", _
                "Width:=", y & "mm", _
                "Height:=", x2 & "mm", _
                "WhichAxis:=", "Z"), _
            Array("NAME:Attributes", _
                "Name:=", leftSideName, _
                "Flags:=", "", _
                "Color:=", "(132 132 193)", _
                "Transparency:=", 0.4, _
                "PartCoordinateSystem:=", "Global", _
                "UDMId:=", "", _
                "MaterialValue:=", """vacuum""", _
                "SurfaceMaterialValue:=", """""", _
                "SolveInside:=", True, _
                "ShellElement:=", False, _
                "ShellElementThickness:=", "0mm", _
                "IsMaterialEditable:=", True, _
                "UseMaterialAppearance:=", False, _
                "IsLightweight:=", False)

        ' 중심 정렬 후 왼쪽으로 이동 (적층 위치 고려)
        dx = -gap - y / 2.0
        dy = cumulativeY - x2 / 2.0
        oEditor.Move _
            Array("NAME:Selections", _
                "Selections:=", leftSideName, _
                "NewPartsModelFlag:=", "Model"), _
            Array("NAME:TranslateParameters", _
                "TranslateVectorX:=", dx & "mm", _
                "TranslateVectorY:=", dy & "mm", _
                "TranslateVectorZ:=", "0mm")

        arrRects(rectCount) = leftSideName
        rectCount = rectCount + 1

        ' ===== 3. Right Side leg (X,Y 좌표 교환으로 자동 회전) =====
        rightSideName = namePrefix & "_Layer" & (i + 1) & "_RightSide"

        ' Y x X2 rectangle 생성 (X, Y 좌표 바꿈)
        oEditor.CreateRectangle _
            Array("NAME:RectangleParameters", _
                "IsCovered:=", True, _
                "XStart:=", "0mm", _
                "YStart:=", "0mm", _
                "ZStart:=", zStart & "mm", _
                "Width:=", y & "mm", _
                "Height:=", x2 & "mm", _
                "WhichAxis:=", "Z"), _
            Array("NAME:Attributes", _
                "Name:=", rightSideName, _
                "Flags:=", "", _
                "Color:=", "(132 132 193)", _
                "Transparency:=", 0.4, _
                "PartCoordinateSystem:=", "Global", _
                "UDMId:=", "", _
                "MaterialValue:=", """vacuum""", _
                "SurfaceMaterialValue:=", """""", _
                "SolveInside:=", True, _
                "ShellElement:=", False, _
                "ShellElementThickness:=", "0mm", _
                "IsMaterialEditable:=", True, _
                "UseMaterialAppearance:=", False, _
                "IsLightweight:=", False)

        ' 중심 정렬 후 오른쪽으로 이동 (적층 위치 고려)
        dx = gap - y / 2.0
        dy = cumulativeY - x2 / 2.0
        oEditor.Move _
            Array("NAME:Selections", _
                "Selections:=", rightSideName, _
                "NewPartsModelFlag:=", "Model"), _
            Array("NAME:TranslateParameters", _
                "TranslateVectorX:=", dx & "mm", _
                "TranslateVectorY:=", dy & "mm", _
                "TranslateVectorZ:=", "0mm")

        arrRects(rectCount) = rightSideName
        rectCount = rectCount + 1

        ' ===== 4. Top Yoke =====
        yokeXSize = 2 * gap + x2
        yokeYSize = x2
        topYokeName = namePrefix & "_Layer" & (i + 1) & "_TopYoke"

        oEditor.CreateRectangle _
            Array("NAME:RectangleParameters", _
                "IsCovered:=", True, _
                "XStart:=", "0mm", _
                "YStart:=", "0mm", _
                "ZStart:=", zStart & "mm", _
                "Width:=", yokeXSize & "mm", _
                "Height:=", yokeYSize & "mm", _
                "WhichAxis:=", "Z"), _
            Array("NAME:Attributes", _
                "Name:=", topYokeName, _
                "Flags:=", "", _
                "Color:=", "(132 132 193)", _
                "Transparency:=", 0.4, _
                "PartCoordinateSystem:=", "Global", _
                "UDMId:=", "", _
                "MaterialValue:=", """vacuum""", _
                "SurfaceMaterialValue:=", """""", _
                "SolveInside:=", True, _
                "ShellElement:=", False, _
                "ShellElementThickness:=", "0mm", _
                "IsMaterialEditable:=", True, _
                "UseMaterialAppearance:=", False, _
                "IsLightweight:=", False)

        ' 상단으로 이동 (적층 위치 고려)
        dx = -yokeXSize / 2.0
        dy = cumulativeY + y / 2.0
        oEditor.Move _
            Array("NAME:Selections", _
                "Selections:=", topYokeName, _
                "NewPartsModelFlag:=", "Model"), _
            Array("NAME:TranslateParameters", _
                "TranslateVectorX:=", dx & "mm", _
                "TranslateVectorY:=", dy & "mm", _
                "TranslateVectorZ:=", "0mm")

        arrRects(rectCount) = topYokeName
        rectCount = rectCount + 1

        ' ===== 5. Bottom Yoke =====
        bottomYokeName = namePrefix & "_Layer" & (i + 1) & "_BottomYoke"

        oEditor.CreateRectangle _
            Array("NAME:RectangleParameters", _
                "IsCovered:=", True, _
                "XStart:=", "0mm", _
                "YStart:=", "0mm", _
                "ZStart:=", zStart & "mm", _
                "Width:=", yokeXSize & "mm", _
                "Height:=", yokeYSize & "mm", _
                "WhichAxis:=", "Z"), _
            Array("NAME:Attributes", _
                "Name:=", bottomYokeName, _
                "Flags:=", "", _
                "Color:=", "(132 132 193)", _
                "Transparency:=", 0.4, _
                "PartCoordinateSystem:=", "Global", _
                "UDMId:=", "", _
                "MaterialValue:=", """vacuum""", _
                "SurfaceMaterialValue:=", """""", _
                "SolveInside:=", True, _
                "ShellElement:=", False, _
                "ShellElementThickness:=", "0mm", _
                "IsMaterialEditable:=", True, _
                "UseMaterialAppearance:=", False, _
                "IsLightweight:=", False)

        ' 하단으로 이동 (적층 위치 고려)
        dx = -yokeXSize / 2.0
        dy = cumulativeY - y / 2.0 - yokeYSize
        oEditor.Move _
            Array("NAME:Selections", _
                "Selections:=", bottomYokeName, _
                "NewPartsModelFlag:=", "Model"), _
            Array("NAME:TranslateParameters", _
                "TranslateVectorX:=", dx & "mm", _
                "TranslateVectorY:=", dy & "mm", _
                "TranslateVectorZ:=", "0mm")

        arrRects(rectCount) = bottomYokeName
        rectCount = rectCount + 1

        ' 다음 레이어를 위해 현재 Y 저장
        prevY = y

        oDesktop.AddMessage "", "", 0, "  5개의 Rectangle 생성 완료"

        ' ===== 6. Unite - 5개 Rectangle 병합 =====
        Dim selectionsStr
        selectionsStr = Join(arrRects, ",")

        oEditor.Unite _
            Array("NAME:Selections", _
                "Selections:=", selectionsStr), _
            Array("NAME:UniteParameters", _
                "KeepOriginals:=", False)

        oDesktop.AddMessage "", "", 0, "  5개의 Rectangle Unite 완료"

        ' Unite 후 이름 변경
        coreName = namePrefix & "_Layer" & (i + 1)
        oEditor.ChangeProperty _
            Array("NAME:AllTabs", _
                Array("NAME:Geometry3DAttributeTab", _
                    Array("NAME:PropServers", arrRects(0)), _
                    Array("NAME:ChangedProps", _
                        Array("NAME:Name", "Value:=", coreName))))

        ' ===== 7. Sweep - Z축으로 확장 =====
        oEditor.SweepAlongVector _
            Array("NAME:Selections", _
                "Selections:=", coreName, _
                "NewPartsModelFlag:=", "Model"), _
            Array("NAME:VectorSweepParameters", _
                "DraftAngle:=", "0deg", _
                "DraftType:=", "Round", _
                "CheckFaceFaceIntersection:=", False, _
                "SweepVectorX:=", "0mm", _
                "SweepVectorY:=", "0mm", _
                "SweepVectorZ:=", windowHeight & "mm")

        oDesktop.AddMessage "", "", 0, "  Sweep 완료"

        ' ===== 8. 재질 변경 =====
        oEditor.ChangeProperty _
            Array("NAME:AllTabs", _
                Array("NAME:Geometry3DAttributeTab", _
                    Array("NAME:PropServers", coreName), _
                    Array("NAME:ChangedProps", _
                        Array("NAME:Material", "Value:=", """" & materialName & """"))))

        oDesktop.AddMessage "", "", 0, "  완료! 통합된 철심: " & coreName
    Next

    ' 화면 업데이트 재개
    oEditor.ResumeUpdate
    oDesktop.AddMessage "", "", 0, "화면 업데이트를 재개했습니다."

    ' 뷰 맞추기
    oEditor.FitAll

    MsgBox "완료! 총 " & rowCount & " 개의 철심 레이어가 생성되었습니다.", vbInformation, "완료"
    oDesktop.AddMessage "", "", 0, "완료! 총 " & rowCount & " 개의 철심 레이어가 생성되었습니다."
End Sub
