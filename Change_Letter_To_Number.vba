'Change letter to number - Letter position in alphabet = number
'test
Function Change_Letter_To_Number(letter As String) As Integer
    
    On Error GoTo Error_handler
    
    Change_Letter_To_Number = Range(letter & 1).column

Error_handler:
Err.Clear

End Function
