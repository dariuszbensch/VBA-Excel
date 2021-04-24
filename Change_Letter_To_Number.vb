'Change letter to number - Letter position in alphabet = number
Function ChangeLetterToNumber(letter As String) As Integer
    
    On Error GoTo Error_handler
    
    ChangeLetterToNumber = Range(letter & 1).column

Error_handler:
Err.Clear

End Function
