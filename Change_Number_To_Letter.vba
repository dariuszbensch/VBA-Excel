'Change number to letter - Number = letter from alphabet on position of number we sent
Function Change_Number_To_Letter(number As Long) As String
    
    On Error GoTo Error_handler
    
    Change_Number_To_Letter = Split(Cells(1, number).Address, "$")(1)
    
Error_handler:
Err.Clear

End Function
