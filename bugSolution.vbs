Function MyFunc(param1)
  On Error Resume Next
  If IsEmpty(param1) Then
    Err.Raise 5, , "Parameter cannot be empty"
    If Err.Number <> 0 Then
      ' Handle the error appropriately
      MsgBox "Error: " & Err.Description, vbCritical
      Err.Clear
      Exit Function
    End If
  End If
  ' ... rest of the function
End Function