'Change number to letter - Number = letter from alphabet on position of number we sent
Function ChangeNumberToLetter(number As Long) As String
    
    On Error GoTo Error_handler
    
    ChangeNumberToLetter = Split(Cells(1, number).Address, "$")(1)
    
Error_handler:
Err.Clear

End Function
