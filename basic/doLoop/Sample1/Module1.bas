Attribute VB_Name = "Module1"
'Option Explicit
Sub Add()
  Dim a As Double, b As Double

  Do
    a = InputBox("Enter first number: ")
    b = InputBox("Enter second number: ")
    Response = MsgBox("The sum is " & a + b, vbOKCancel)
    If Response = vbOK Then
      
    Else
      Exit Do
    End If
  Loop

End Sub
