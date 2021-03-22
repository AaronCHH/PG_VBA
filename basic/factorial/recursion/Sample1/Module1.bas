Attribute VB_Name = "Module1"
Function FactRecurs(n)
  If n > 0 Then
    FactRecurs = n * FactRecurs(n-1)
  Else
    FactRecurs = 1
  End If  
End Function
