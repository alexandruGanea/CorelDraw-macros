Attribute VB_Name = "LayerManagement"
Sub CreateLayers()
    Dim l As Layer
    Dim emptyLayer As Layer
    
    For Each l In ActivePage.AllLayers
        If InStr(1, l.Name, "Layer") And l.Shapes.Count = 0 Then
            l.Delete
        End If
    Next l
           
    Set l = FindLayer(ActivePage, "C")
    Set l = FindLayer(ActivePage, "P")
    Set l = FindLayer(ActivePage, "S")
    Set l = FindLayer(ActivePage, "RM")
    Set l = FindLayer(ActivePage, "I")
    
    ActivePage.Layers("S").Activate
End Sub
Sub MoveObjToInfo()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.MoveToLayer ActivePage.Layers("I")
End Sub
Sub MoveObjToRM()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.MoveToLayer ActivePage.Layers("RM")
End Sub
Sub MoveObjToPrint()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.MoveToLayer ActivePage.Layers("P")
End Sub
Sub MoveObjToCut()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.MoveToLayer ActivePage.Layers("S")
End Sub
Sub MoveObjToBoard()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.MoveToLayer ActivePage.Layers("C")
End Sub
Sub ToggleVisibleInfoLayer()
    If ActivePage.Layers("I").Visible = False Then ActivePage.Layers("I").Visible = True Else: ActivePage.Layers("I").Visible = False
    ActivePage.Layers("I").Activate
End Sub
Sub ToggleVisibleRMLayer()
    If ActivePage.Layers("RM").Visible = False Then ActivePage.Layers("RM").Visible = True Else: ActivePage.Layers("RM").Visible = False
    ActivePage.Layers("RM").Activate
End Sub
Sub ToggleVisibleCutLayer()
    If ActivePage.Layers("S").Visible = False Then ActivePage.Layers("S").Visible = True Else: ActivePage.Layers("S").Visible = False
    ActivePage.Layers("S").Activate
End Sub
Sub ToggleVisiblePrintLayer()
    If ActivePage.Layers("P").Visible = False Then ActivePage.Layers("P").Visible = True Else: ActivePage.Layers("P").Visible = False
    ActivePage.Layers("P").Activate
End Sub
Sub ToggleVisibleBoardLayer()
    If ActivePage.Layers("C").Visible = False Then ActivePage.Layers("C").Visible = True Else: ActivePage.Layers("C").Visible = False
    ActivePage.Layers("C").Activate
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


