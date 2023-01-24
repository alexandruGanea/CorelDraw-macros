Attribute VB_Name = "Tools"
Sub BoardGenerator()
    Dim s As Shape, t1 As Shape, srSelection As ShapeRange, crv As Curve
    Dim x As Double, y As Double, h As Double, w As Double
    Dim dimension As String
    
    Dim lC As Layer
    Dim lI As Layer
        
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup
    Optimization = True
    
    Set lActive = ActiveLayer
    Set srSelection = ActiveSelectionRange
    
    If srSelection.Count = 0 Then MsgBox "NO CUT LINE SELECTED": Exit Sub
    
    srSelection.GetBoundingBox x, y, w, h
    
    On Error GoTo ErrHandler
            
        Set lC = FindLayer(ActivePage, "C")
        lC.Activate
        
    Set s = ActiveVirtualLayer.CreateRectangle2(x - 5, y - 5, w + 10, h + 10)
        s.Outline.Color.CMYKAssign 0, 60, 100, 0
        
    Set s = Nothing
        
        Set lI = FindLayer(ActivePage, "I")
        lI.Activate
        
    dimension = CInt(h + 10) & "x" & CInt(w + 10)

        Set t1 = ActiveVirtualLayer.CreateArtisticText(x - 5, y + h + 25, dimension)
            t1.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
            t1.SetSize 0, 20
            t1.Outline.SetNoOutline
            
    
                   
ExitSub:
    ActiveDocument.ClearSelection
    ActiveDocument.EndCommandGroup
    Optimization = False
    ActiveWindow.Refresh
    Refresh
    Exit Sub

ErrHandler:
    MsgBox "Unexpected error occured: " & Err.Description & " [" & Err.Number & "]", vbCritical, "Error"
    Resume ExitSub
End Sub
Sub GroupAndCombineShapes()
    Dim cutRange As ShapeRange, creaseRange As ShapeRange, otherRange As ShapeRange
    Dim cut As Shape, crease As Shape, other As Shape, s As Shape
    Dim srSelection As Shapes
    Dim combinedShapes As ShapeRange
    Dim l As Layer
    Dim g As Shape
   
    ActiveDocument.Unit = cdrMillimeter
    Set srSelection = ActiveSelection.Shapes
          
    If srSelection.Count = 0 Then MsgBox "NO SHAPES SELECTED": Exit Sub
        
    Set l = FindLayer(ActivePage, "S")
        l.Activate
        
    Set combinedShapes = ActiveSelectionRange
                               
                          
    Set cutRange = srSelection.FindShapes(, , True, "@outline.color = cmyk(0,0,0,100)")
    If cutRange.Count <= 1 Then
        Set cut = cutRange.FirstShape
    Else:
        cutRange.Combine
        combinedShapes.AddRange cutRange
    End If
                    
    Set creaseRange = srSelection.FindShapes(, , True, "@outline.color = cmyk(0,100,100,0)")
    If creaseRange.Count <= 1 Then
        Set crease = creaseRange.FirstShape
        
    Else:
        Set crease = creaseRange.Combine
        combinedShapes.Add crease
    End If
             
    Set otherRange = srSelection.FindShapes(, , True, "@outline.color = cmyk(100,0,0,0)")
    If otherRange.Count <= 1 Then
        Set other = otherRange.FirstShape
    Else:
        Set other = otherRange.Combine
        combinedShapes.Add other
    End If
    
        Set g = combinedShapes.Group
        ActiveWindow.Refresh
        
        
        
End Sub
Sub ImportFilesForPrint()

    Dim l As Layer

    Set l = FindLayer(ActivePage, "P")
            l.Activate
         
        Dim impopt As StructImportOptions
    Set impopt = CreateStructImportOptions
    With impopt
        .Mode = cdrImportFull
        .MaintainLayers = True
    End With

ExitSub:
    ActiveDocument.ClearSelection
    ActiveWindow.Refresh
    Refresh
    Exit Sub
End Sub
Sub GetArea()
    Dim srSelection As ShapeRange
    Dim obj As Shape
    Dim area As Double

    ActiveDocument.Unit = cdrMillimeter
    Set srSelection = ActiveSelection.Shapes
              
    If srSelection.Count = 0 Then MsgBox "NO SHAPES SELECTED": Exit Sub
    
    Set obj = srSelection.Combine
    Set obj = obj.ConvertToCurves
    Set area = obj.Area(
    MsgBox (area)
        
    
End Sub
Function FindLayer(ByVal p As Page, ByVal Name As String) As Layer
    Dim lLayerFound As Layer
    Dim l As Layer
    
    Set lLayerFound = Nothing
    
    For Each l In p.Layers
        If l.Name = Name Then
            Set lLayerFound = l
            Exit For
        End If
    Next l
    
    If lLayerFound Is Nothing Then
        Set lLayerFound = ActivePage.CreateLayer(Name)
        Select Case lLayerFound.Name
        Case "C"
            lLayerFound.Color = CreateCMYKColor(0, 60, 100, 0)
        Case "P"
            lLayerFound.Color = CreateCMYKColor(100, 0, 0, 0)
        Case "S"
            lLayerFound.Color = CreateCMYKColor(0, 100, 0, 0)
        Case "I"
            lLayerFound.Color = CreateCMYKColor(100, 0, 100, 0)
        Case Else
        End Select
    End If
    
    Set FindLayer = lLayerFound
End Function

