Attribute VB_Name = "Module1"
Sub TestStatic()  
  a = 15
  c = Accumulate(a)  

  MsgBox c
  MsgBox Accumulate(28)
End Sub

Function Accumulate(n)  
  Static sum
  sum = sum + n
  Accumulate = sum  
End Function

