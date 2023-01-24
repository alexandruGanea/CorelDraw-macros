Attribute VB_Name = "RegMarks"
Sub RegMarks()
    Dim s As Shape, srSelection As ShapeRange, srCutMarks As New ShapeRange
    Dim x As Double, y As Double, h As Double, w As Double
    Dim lActive As Layer, lRM As Layer
    
    ActiveDocument.Unit = cdrMillimeter
    
    Set lActive = ActiveLayer
    Set srSelection = ActiveSelectionRange
    
    If srSelection.Count = 0 Then MsgBox "NO CUT LINE SELECTED": Exit Sub
    
    srSelection.GetBoundingBox x, y, w, h
    
    ActiveDocument.BeginCommandGroup "RM"
    On Error GoTo ErrHandler
    Optimization = True
        
        Set lRM = FindLayer(ActivePage, "RM")
        lRM.Activate
        
        Set s = ActiveVirtualLayer.CreateEllipse2(x - 5, y - 5, 3, 3)
        s.Fill.UniformColor.CMYKAssign 0, 0, 0, 98
        s.Outline.SetNoOutline
        
       srCutMarks.Add s
        srCutMarks.Add s.Duplicate(0, h + 10)
        srCutMarks.Add s.Duplicate(w + 10, h + 10)
        srCutMarks.Add s.Duplicate(w + 10, 0)
        
        ActiveDocument.LogCreateShapeRange srCutMarks
        lActive.Activate
    
        Set s = Nothing
        Set srCutMarks = Nothing
    
ExitSub:
    ActiveDocument.EndCommandGroup
    Optimization = False
    ActiveDocument.ClearSelection
    ActiveWindow.Refresh
    Refresh
    Exit Sub

ErrHandler:
    MsgBox "Unexpected error occured: " & Err.Description & " [" & Err.Number & "]", vbCritical, "Error"
    Resume ExitSub
End Sub
Sub CutMarks()
    Dim s As Shape, srSelection As ShapeRange, crv As Curve
    Dim x As Double, y As Double, h As Double, w As Double
    Dim lRM As Layer
    
    ActiveDocument.Unit = cdrMillimeter
    
    Set lActive = ActiveLayer
    Set srSelection = ActiveSelectionRange
    
    If srSelection.Count = 0 Then MsgBox "NO CUT LINE SELECTED": Exit Sub
    
    srSelection.GetBoundingBox x, y, w, h
    
    On Error GoTo ErrHandler
            
        Set lRM = FindLayer(ActivePage, "RM")
        lRM.Activate
        
    Set s = ActiveLayer.CreateLineSegment(x + 5, y - 5, x - 5, y - 5)
    s.Fill.ApplyNoFill
    s.Outline.SetPropertiesEx 0.003, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
    Set crv = ActiveDocument.CreateCurve
    With crv.CreateSubPath(x + 5, y - 5)
        .AppendLineSegment x - 5, y - 5
        .AppendLineSegment x - 5, y + 5
    End With
    s.Curve.CopyAssign crv
       
    Set s = s.Duplicate(0, h)
            s.Flip (cdrFlipVertical)
            s.Outline.Color.CMYKAssign 0, 100, 100, 0
    
    Set s = s.Duplicate(w, 0)
            s.Flip (cdrFlipHorizontal)
            s.Outline.Color.CMYKAssign 0, 0, 0, 100
            
    Set s = s.Duplicate(0, -h)
            s.Flip (cdrFlipVertical)
         
        Set s = Nothing
           
ExitSub:
    ActiveDocument.ClearSelection
    ActiveWindow.Refresh
    Refresh
    Exit Sub

ErrHandler:
    MsgBox "Unexpected error occured: " & Err.Description & " [" & Err.Number & "]", vbCritical, "Error"
    Resume ExitSub
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











