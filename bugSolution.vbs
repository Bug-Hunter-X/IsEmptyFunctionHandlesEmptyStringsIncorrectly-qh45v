Function MyFunction(param1)
  If Len(param1) = 0 Then
    ' Handle empty string parameter
    ' ... specific handling for empty string ...
  ElseIf IsEmpty(param1) Then
    ' Handle truly empty parameter
    ' ... specific handling for uninitialized variable ...
  Else
    ' ... rest of function
  End If
End Function