'Open forwarded URL address
Option Explicit
Private Declare PtrSafe Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long
                                                        
                                                        
                                                        
Sub Open_Website(strUrl As String)
On Error GoTo wellsrLaunchError
    Dim r As Long
    r = ShellExecute(0, "open", strUrl, 0, 0, 1)
    If r = 5 Then 'if access denied, try this alternative
            r = ShellExecute(0, "open", "rundll32.exe", "url.dll,FileProtocolHandler " & strUrl, 0, 1)
    End If
    Exit Sub
wellsrLaunchError:
MsgBox "Error encountered while trying to launch URL." & vbNewLine & vbNewLine & "Error: " & Err.number & ", " & Err.Description, vbCritical, "Error Encountered"
End Sub
