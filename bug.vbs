Function MyFunction(param1)
  If IsEmpty(param1) Then
    Err.Raise 13, , "Type mismatch"
  End If
  ' ... rest of the function
End Function