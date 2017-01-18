Attribute VB_Name = "Module1"
Sub Test()
  Dim i As Long , n As Long
  Dim tstart As Double , tfinish As Double
  Application.ScreenUpdating = False
  n = InputBox("Number of values = ", "Bubble Sort")
  Range ("b3") . Select
  ActiveCell.Value = n
  ReDim a(n) As Double , b(n) As Double

  'generate n random numbers
  For i = 1 To n
    a(i) = Rnd
  Next i

  'implement bubble sort
  Range("b4").Select
  tstart = Timer
  Call Bubble(n, a , b)
  tfinish = Timer
  ActiveCell.Value = tfinish - tstart

  'display unsorted and sorted values
  Sheets("sheet1").Select
  Range("a6:c60005").ClearContents
  Range("a6").Select

  For i = 1 To n
    ActiveCell.Value = i
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = a(i)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = b(i)
    ActiveCell.Offset(1, -2).Select
  Next i

  Range("l10").Select
  Application.ScreenUpdating = True

End Sub

' ===== '
Sub Bubble(n, a, b)
  'sorts an array in ascending order
  'using the bubble sort
  Dim m As integer i As integer
  Dim switch As Boolean
  Dim dum As Double

  For i = 1 To n
    b(i) = a(i)
  Next i

  m = n - 1
  Do 
    switch = False 'loop through passes
  
    For i = 1 To m 'loop through array
      If b(i) > b(i + 1) Then
        dum = b(i)
        b(i) = b(i + 1)
        b(i + 1) = dum
        switch = True
      End If
    Next i

    If switch = False Then Exit
    
    m = m - 1    
  Loop
End Sub
