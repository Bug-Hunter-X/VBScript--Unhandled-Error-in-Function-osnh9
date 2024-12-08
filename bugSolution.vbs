Function MyFunction(param1, param2)
  On Error Resume Next
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise 13, , "Parameters cannot be empty"
  End If
  On Error GoTo 0
  ' ... rest of the function
End Function

Sub CallMyFunction()
  On Error GoTo ErrorHandler
  Call MyFunction(param1:=1, param2:=2)
  ' ...more code...
  Exit Sub
ErrorHandler:
  MsgBox "Error Number: " & Err.Number & ", Description: " & Err.Description
End Sub