Sub roundmacro()

    
Dim MyRegion As Range
Dim MyNum As Variant
Dim i As Integer

Set MyRegion = Selection

    'Begin Loop
     While i < MyRegion.count
    
        'traverses through each cell in region
        If i = 0 Then
            ActiveCell.Offset(0, 0).Select
        Else
            ActiveCell.Offset(1, 0).Select
        End If
        
        'replaces cell value with ROUND formula
        MyNum = ActiveCell.Value
        ActiveCell.Formula = "=ROUND(" & MyNum & ",0)"
        
        'format to General, Commas, and 0 decimal places
        ActiveCell.NumberFormat = "General"
        ActiveCell.Style = "Comma"
        ActiveCell.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
        i = i + 1
        
     Wend
End Sub
