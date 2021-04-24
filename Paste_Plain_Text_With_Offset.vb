'Wkleja tekst sformatowany z przesumiêciem wierszy (CTRL + V)
Sub Paste_Plain_Text_With_Offset()



'Obs³uga b³êdów
    On Error GoTo Error_handler
      

    
    
'Deklaracja zmiennych
    Dim lastCell As String
    Dim lastRow As Integer
    
    Dim position As Variant: position = ActiveCell.Address
    Dim firstSheetName As String: firstSheetName = ActiveSheet.Name
    Dim data As String: data = "Data"
    
    
    
'Odkrywa arkusz Data i przechodzi do niego
    Sheet_Unhide_Existing (data)
    Worksheets(data).Activate



'Wklejanie wartosci do Data
    Data_Paste_Only_Values



'Pobieanie pozycji ostatnich elemantów z Data
    lastCell = Last_Cell_Address(data)
    lastRow = Last_Row_Address(data)



'Powrót do Exception Helper
    Worksheets(firstSheetName).Activate
    Range(position).Select



'Utworzenie wymaganego miejsca
    Dim i As Integer: i = 0
    If Application.CountA(ActiveCell.EntireRow) = 0 Then i = 1
    
    Do While i < lastRow
        Selection.EntireRow.Insert
        i = i + 1
    Loop
    
    
    
'Kopiowanie danych z Data do EXCEPTIONS HELPER
    Worksheets(data).Range("A1", lastCell).Copy Worksheets(firstSheetName).Range(position)
    
    
    
    
'Czyszczenie bledów
Error_handler:
Err.Clear



'Czyszczenie powierzchni roboczej "Data"
    Sheet_Delete_Existing (data)
    Sheet_Create_New (data)
    Sheet_Hide_Existing (data)
    Worksheets(firstSheetName).Activate



'Powrot na pozycje
    Range(position).Select
    
    
    
End Sub
