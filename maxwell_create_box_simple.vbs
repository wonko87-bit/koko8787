' ============================================================================
' Maxwell 3D 박스 생성 - 간단한 버전 (VBScript)
' ============================================================================
'
' 실제 Maxwell 2021 R1에서 녹화된 API를 기반으로 한 간단한 스크립트입니다.
' 파일 상단의 파라미터만 수정하면 바로 사용할 수 있습니다.
'
' 사용 방법:
' ---------
' 1. 아래 파라미터 값을 원하는 크기로 수정합니다
' 2. Maxwell에서 Tools → Run Script로 실행합니다
' ============================================================================

Option Explicit

' ============================================================================
' 파라미터 설정 (여기를 수정하세요!)
' ============================================================================
Dim boxWidth, boxDepth, boxHeight, boxName, materialName

boxWidth = 20        ' X 방향 너비 (mm)
boxDepth = 10        ' Y 방향 깊이 (mm)
boxHeight = 5        ' Z 방향 높이 (mm)
boxName = "Box1"     ' 박스 이름
materialName = "vacuum"  ' 재질 (vacuum, aluminum, copper, steel, iron 등)

' ============================================================================
' 박스 생성 실행
' ============================================================================

Dim oAnsoftApp, oDesktop, oProject, oDesign, oEditor

' Desktop 연결
Set oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")
Set oDesktop = oAnsoftApp.GetAppDesktop()
oDesktop.RestoreWindow

' 프로젝트와 디자인 설정
Set oProject = oDesktop.GetActiveProject()
If oProject Is Nothing Then
    Set oProject = oDesktop.NewProject()
    oDesktop.AddMessage "", "", 0, "새 프로젝트를 생성했습니다."
End If

Set oDesign = oProject.GetActiveDesign()
If oDesign Is Nothing Then
    oProject.InsertDesign "Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", ""
    Set oDesign = oProject.SetActiveDesign("Maxwell3DDesign1")
    oDesktop.AddMessage "", "", 0, "새 Maxwell 3D 디자인을 생성했습니다."
End If

' 3D Modeler 에디터 가져오기
Set oEditor = oDesign.SetActiveEditor("3D Modeler")

' 박스 생성 (Maxwell 2021 R1 녹화 API 기반)
Dim arrBoxParams(1), arrAttributes(1)

arrBoxParams(0) = "NAME:BoxParameters"
arrBoxParams(1) = Array(_
    "XPosition:=", "0mm", _
    "YPosition:=", "0mm", _
    "ZPosition:=", "0mm", _
    "XSize:=", CStr(boxWidth) & "mm", _
    "YSize:=", CStr(boxDepth) & "mm", _
    "ZSize:=", CStr(boxHeight) & "mm" _
)

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

oEditor.CreateBox arrBoxParams, arrAttributes

' 뷰 맞추기
oEditor.FitAll

' 완료 메시지
Dim msg
msg = "박스 생성 완료!" & vbCrLf & vbCrLf & _
      "이름: " & boxName & vbCrLf & _
      "크기: " & boxWidth & " x " & boxDepth & " x " & boxHeight & " mm" & vbCrLf & _
      "재질: " & materialName

MsgBox msg, vbInformation, "Maxwell 3D"
oDesktop.AddMessage "", "", 0, msg
