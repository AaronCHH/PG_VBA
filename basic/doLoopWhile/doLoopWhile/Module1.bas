Attribute VB_Name = "Module1"
Sub doLoopDoWhile()

  Debug.Print ("Do Until")
  x = 10

  Do While x < 0
    Debug.Print (x)
    x = x - 5
  Loop

  Debug.Print ("=====")
  Debug.Print ("Do Loop While")
  x = 10

  Do
    Debug.Print (x)
    x = x - 5
  Loop While x < 0

End Sub
