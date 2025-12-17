' ==============================================================================
' Simcenter Magnet 3D Box Modeling Script (VBScript)
' ==============================================================================
'
' This script creates a 3D rectangular box (cuboid) in Simcenter Magnet.
' The box is created with one corner at the origin (0,0,0) and extends
' in the positive direction.
'
' Author: Claude
' Date: 2025-12-17
' ==============================================================================

Option Explicit

' ==============================================================================
' Main Subroutine - Entry Point
' ==============================================================================
Sub Main()
    Dim width, depth, height
    Dim materialName

    ' Define box dimensions (in mm)
    width = 100.0    ' X direction
    depth = 50.0     ' Y direction
    height = 30.0    ' Z direction

    ' Define material
    materialName = "Air"

    ' Display information
    Call MsgBox("Creating 3D Box with dimensions:" & vbCrLf & _
                "Width (X):  " & width & " mm" & vbCrLf & _
                "Depth (Y):  " & depth & " mm" & vbCrLf & _
                "Height (Z): " & height & " mm" & vbCrLf & _
                "Material:   " & materialName, vbInformation, "Simcenter Magnet - Box Creation")

    ' Create the box
    Call CreateBox(width, depth, height, materialName)

    Call MsgBox("Box created successfully!", vbInformation, "Success")
End Sub

' ==============================================================================
' CreateBox - Creates a rectangular box component
' ==============================================================================
' Parameters:
'   w - Width in X direction (mm)
'   d - Depth in Y direction (mm)
'   h - Height in Z direction (mm)
'   material - Material name
' ==============================================================================
Sub CreateBox(w, d, h, material)
    Dim oDoc
    Dim oView
    Dim oComponent
    Dim x0, y0, z0
    Dim x1, y1, z1

    ' Validate input
    If w <= 0 Or d <= 0 Or h <= 0 Then
        Call MsgBox("Error: All dimensions must be positive values!", vbCritical, "Error")
        Exit Sub
    End If

    ' Get the active document
    Set oDoc = GetDocument()
    Set oView = oDoc.GetView()

    ' Define corner points
    x0 = 0.0  : y0 = 0.0  : z0 = 0.0     ' Origin corner
    x1 = w    : y1 = d    : z1 = h       ' Opposite corner

    ' Create a new component
    Set oComponent = oView.NewComponent("Box")

    ' Create box using MakeComponentInABox method
    Call oComponent.MakeComponentInABox(x0, y0, z0, x1, y1, z1)

    ' Set material
    Call oComponent.SetMaterial(material)

    ' Refresh view
    Call oView.ViewAll()

    ' Clean up
    Set oComponent = Nothing
    Set oView = Nothing
    Set oDoc = Nothing
End Sub

' ==============================================================================
' CreateBoxAdvanced - Creates a box using edge extrusion method
' ==============================================================================
' This method provides more control over the geometry creation
' ==============================================================================
Sub CreateBoxAdvanced(w, d, h, material, componentName)
    Dim oDoc, oView, oComponent
    Dim oEdge1, oEdge2, oEdge3, oEdge4
    Dim oCurve
    Dim x0, y0, z0

    ' Validate input
    If w <= 0 Or d <= 0 Or h <= 0 Then
        Call MsgBox("Error: All dimensions must be positive values!", vbCritical, "Error")
        Exit Sub
    End If

    ' Get the active document
    Set oDoc = GetDocument()
    Set oView = oDoc.GetView()

    ' Origin point
    x0 = 0.0 : y0 = 0.0 : z0 = 0.0

    ' Create the bottom rectangle edges in XY plane
    Set oEdge1 = oView.NewLine(x0, y0, z0, x0 + w, y0, z0)          ' Bottom edge
    Set oEdge2 = oView.NewLine(x0 + w, y0, z0, x0 + w, y0 + d, z0)  ' Right edge
    Set oEdge3 = oView.NewLine(x0 + w, y0 + d, z0, x0, y0 + d, z0)  ' Top edge
    Set oEdge4 = oView.NewLine(x0, y0 + d, z0, x0, y0, z0)          ' Left edge

    ' Create a curve from edges
    Set oCurve = oView.NewCurve()
    Call oCurve.AddEdge(oEdge1)
    Call oCurve.AddEdge(oEdge2)
    Call oCurve.AddEdge(oEdge3)
    Call oCurve.AddEdge(oEdge4)

    ' Create component and extrude
    Set oComponent = oView.NewComponent(componentName)
    Call oComponent.ExtrudeEdges(oCurve, 0, 0, h)

    ' Set material
    Call oComponent.SetMaterial(material)

    ' Refresh view
    Call oView.ViewAll()

    ' Clean up
    Set oComponent = Nothing
    Set oCurve = Nothing
    Set oEdge1 = Nothing
    Set oEdge2 = Nothing
    Set oEdge3 = Nothing
    Set oEdge4 = Nothing
    Set oView = Nothing
    Set oDoc = Nothing
End Sub

' ==============================================================================
' CreateParametricBox - Interactive box creation with user input
' ==============================================================================
Sub CreateParametricBox()
    Dim width, depth, height, material
    Dim userInput

    ' Get width from user
    userInput = InputBox("Enter box width (X direction) in mm:", "Box Width", "100")
    If userInput = "" Then Exit Sub
    width = CDbl(userInput)

    ' Get depth from user
    userInput = InputBox("Enter box depth (Y direction) in mm:", "Box Depth", "50")
    If userInput = "" Then Exit Sub
    depth = CDbl(userInput)

    ' Get height from user
    userInput = InputBox("Enter box height (Z direction) in mm:", "Box Height", "30")
    If userInput = "" Then Exit Sub
    height = CDbl(userInput)

    ' Get material from user
    material = InputBox("Enter material name:", "Material", "Air")
    If material = "" Then material = "Air"

    ' Create the box
    Call CreateBox(width, depth, height, material)

    Call MsgBox("Parametric box created successfully!" & vbCrLf & _
                "Dimensions: " & width & " x " & depth & " x " & height & " mm" & vbCrLf & _
                "Material: " & material, vbInformation, "Success")
End Sub

' ==============================================================================
' Example: Create multiple boxes
' ==============================================================================
Sub CreateMultipleBoxes()
    ' Create air region
    Call CreateBox(500, 500, 500, "Air")

    ' Create magnet core at offset position
    ' Note: You would need to modify CreateBox to accept offset parameters
    Call CreateBox(100, 100, 50, "Steel")

    ' Create coil
    Call CreateBox(120, 120, 30, "Copper")
End Sub

' ==============================================================================
' Run the main subroutine
' ==============================================================================
Call Main()
