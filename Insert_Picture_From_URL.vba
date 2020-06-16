Sub Insert_Picture_From_URL()
'Updateby Extendoffice 20161116
    Dim Pshp As Shape
    Dim xRg As Range
    Dim xCol As Long
    On Error Resume Next
    Application.ScreenUpdating = False
    Set Rng = ActiveSheet.Range("A30:A35")
    For Each cell In Rng
        filenam = cell
        ActiveSheet.Pictures.Insert(filenam).Select
        Set Pshp = Selection.ShapeRange.Item(1)
        If Pshp Is Nothing Then GoTo lab
        xCol = cell.column
        Set xRg = Cells(cell.row, xCol)
        With Pshp
            .LockAspectRatio = msoFalse
            .Width = 50
           .Height = 50
            .Top = xRg.Top + (xRg.Height - .Height) / 2
            .Left = xRg.Left + (xRg.Width - .Width) / 2
        End With
lab:
    Set Pshp = Nothing
    Range("A3").Select
    Next
    Application.ScreenUpdating = True
End Sub
