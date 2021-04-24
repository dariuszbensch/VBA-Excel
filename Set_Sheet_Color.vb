'Changing color of sheet tab - Declaration example: Set_Sheet_Color("Sheet1", RGB(255, 0, 0))
Function Set_Sheet_Color(sheetName As String, color As Variant)

    On Error GoTo Error_handler

        With ActiveWorkbook.Sheets(sheetName).Tab
            .color = color
            .TintAndShade = 0
        End With
        
Error_handler:
Err.Clear

End Function
