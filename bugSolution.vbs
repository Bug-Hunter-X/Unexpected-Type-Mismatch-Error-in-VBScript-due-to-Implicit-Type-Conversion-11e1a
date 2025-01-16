Function MyFunction(param1)
  On Error Resume Next
  If IsEmpty(param1) Or VarType(param1) <> vbString Then
    Err.Raise 13, , "Invalid parameter type: Expected String"
  End If
  On Error GoTo 0
  ' ... rest of the function
End Function