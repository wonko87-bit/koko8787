' ==============================================================================
' Simcenter Magnet - Simple 3D Box Creation Script
' ==============================================================================
'
' Quick script to create a box from origin (0,0,0) in positive direction
' Modify the parameters below and run in Simcenter Magnet
'
' ==============================================================================

' ===== PARAMETERS - MODIFY THESE VALUES =====
Dim boxWidth
Dim boxDepth
Dim boxHeight
Dim boxMaterial

boxWidth = 100.0     ' Width in X direction (mm)
boxDepth = 50.0      ' Depth in Y direction (mm)
boxHeight = 30.0     ' Height in Z direction (mm)
boxMaterial = "Air"  ' Material name
' ============================================

' Create the box
Dim oDoc, oView, oComponent

Set oDoc = GetDocument()
Set oView = oDoc.GetView()
Set oComponent = oView.NewComponent("Box")

' Create box from (0,0,0) to (width, depth, height)
Call oComponent.MakeComponentInABox(0, 0, 0, boxWidth, boxDepth, boxHeight)
Call oComponent.SetMaterial(boxMaterial)
Call oView.ViewAll()

' Show confirmation
Call MsgBox("Box created!" & vbCrLf & _
            "Size: " & boxWidth & " x " & boxDepth & " x " & boxHeight & " mm" & vbCrLf & _
            "Material: " & boxMaterial, vbInformation, "Success")
