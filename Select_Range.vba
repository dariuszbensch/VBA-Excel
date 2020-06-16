'Select any choosen range in one move in active sheet
'Columns - Example: "A" or "A:C"
'Rows - Example: "1" or "1:5"
'Range - Example: "A1" or "A1:B2"

'Example
'arr = Array("A", "1", "C:D", "5:10")
'Call Select_Range(arr)

Function Select_Range(arr As Variant)

    On Error GoTo Error_handler
    
        Dim selected As String
    
        For Each e In arr
            selected = selected & e & ":" & e & ","
        Next
    
        'Delete last symbol (,) to get proper string
        selected = Left(selected, Len(selected) - 1)
    
        'Select choosen range
        Range(selected).Select

Error_handler:
Err.Clear

End Function
