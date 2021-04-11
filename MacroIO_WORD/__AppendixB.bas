Attribute VB_Name = "AppendixB"
Option Explicit

Sub DrawSine2()

' Dampened sine wave of small stars

Const pi = 3.1416

Dim i As Integer
Dim x As Single, y As Single
Dim rng As Range     ' For starting point
Dim n As Single      ' Cycle length in inches
Dim k As Integer     ' k stars
Dim ScaleY As Single ' Vertical scaling
Dim sSize As Single  ' Star size
Dim sDamp1 As Single  ' Dampening factor
Dim sDamp2 As Single  ' Dampening factor
Dim cCycles As Integer  ' Number of cycles
Dim sh As Shape

cCycles = 3
sDamp1 = 1
sDamp2 = 0.2
n = 2
k = 20
ScaleY = 0.5
sSize = InchesToPoints(0.1)

' Start at insertion point
Set rng = Selection.Range

' Loop for first curve with phase shift
For i = 1 To cCycles * k
   x = n * i / k
   y = ScaleY * Sin((2 * pi * i) / k + n) * _
      (sDamp1 / (x + sDamp2))
   y = InchesToPoints(y)
   x = InchesToPoints(x)
   Set sh = ActiveDocument.Shapes.AddShape _
      (msoShape5pointStar, x, y, sSize, sSize, rng)
   sh.Fill.ForeColor.RGB = RGB(192, 192, 192)  ' 25% gray
   sh.Fill.Visible = msoTrue
Next i

End Sub


Sub DrawName()

' Random placement of large stars with name

Const pi = 3.1416

Dim i As Integer
Dim x As Single, y As Single
Dim z As Single
Dim rng As Range     ' For starting point
Dim n As Single      ' Cycle length in inches
Dim k As Integer     ' k stars
Dim sSize As Single  ' Star size
Dim sh As Shape
Dim sName As String  ' Name to display

sName = "Steven Roman"
n = 5
k = Len(sName)
sSize = InchesToPoints(0.5)

' Start at insertion point
Set rng = Selection.Range
Randomize Timer
z = 0#

' Loop for first curve with phase shift
For i = 1 To k
   If Mid(sName, i, 1) <> " " Then
      x = n * i / k
      x = InchesToPoints(x)
      
      ' Get random 0 or 1. Go up or down accordingly.
      If Int(2 * Rnd) = 0 Then
         z = z + 0.2
      Else
         z = z - 0.2
      End If

      y = InchesToPoints(z)
      Set sh = ActiveDocument.Shapes.AddShape _
         (msoShape5pointStar, x, y, sSize, sSize, rng)
      
      ' Add shading
      sh.Fill.ForeColor.RGB = RGB(230, 230, 230)
      sh.Fill.Visible = msoTrue
         
      ' Add text
      sh.TextFrame.TextRange.Text = Mid(sName, i, 1)
      sh.TextFrame.TextRange.Font.Size = 10
      sh.TextFrame.TextRange.Font.Name = "Arial"
      sh.TextFrame.TextRange.Font.Bold = True
      
   End If
Next i

End Sub
Sub DrawRose()

' Draw rose of small stars

Const pi = 3.1416
Dim t As Single
Dim i As Integer
Dim x As Single, y As Single
Dim rng As Range     ' For starting point
Dim n As Single      ' Number of stars per cycle
Dim k As Integer     ' Number of cycles
Dim sSize As Single  ' Star size
Dim r As Integer     ' half the number of petals
Dim sh As Shape

' For a 3-petal rose
r = 3
k = 1
n = 100

' For a 4-petal rose
'r = 2
'k = 2
'n = 150      ' Number of stars

sSize = InchesToPoints(0.03)

' Start curve at insertion point
Set rng = Selection.Range

For i = 1 To n
   t = k * pi * i / n
   x = Sin(r * t) * Sin(t)
   y = Sin(r * t) * Cos(t)
   x = InchesToPoints(x)
   y = InchesToPoints(y)
   Set sh = ActiveDocument.Shapes.AddShape _
      (msoShape5pointStar, x, y, sSize, sSize, rng)
Next i

End Sub
Sub DrawSpiral()

' Draw spiral of small stars

Const pi = 3.1416
Dim t As Single
Dim i As Integer
Dim z As Single
Dim x As Single, y As Single
Dim rng As Range     ' For starting point
Dim n As Single      ' Number of stars per cycle
Dim k As Integer     ' Length of spiral
Dim sSize As Single  ' Star size
Dim sh As Shape

n = 80      ' Number of stars
k = 8       ' Length
sSize = InchesToPoints(0.03)

' Start curve at insertion point
Set rng = Selection.Range

For i = 5 To n
   t = k * pi * i / n
   x = 2 * (1 / t) * Sin(t)
   y = 2 * (1 / t) * Cos(t)
   x = InchesToPoints(x)
   y = InchesToPoints(y)
   
   Set sh = ActiveDocument.Shapes.AddShape _
      (msoShape5pointStar, x, y, sSize, sSize, rng)
   z = 256 * i / n
   sh.Line.ForeColor.RGB = RGB(z, z, z) ' vary line color
   sh.Line.Visible = msoTrue
Next i

End Sub
Sub DrawHypocycloid()

' Draw hypocycloid of small stars

Const pi = 3.1416
Dim t As Single
Dim i As Integer
Dim x As Single, y As Single
Dim rng As Range     ' For starting point
Dim n As Single
Dim k As Integer
Dim sSize As Single  ' Star size
Dim r As Integer
Dim r0 As Integer
Dim R1 As Integer
Dim sh As Shape
Dim sc As Single

r = 1
r0 = 3 * r
R1 = 8 * r

n = 400
k = 4
sc = 0.1
sSize = InchesToPoints(0.03)

' Start curve at insertion point
Set rng = Selection.Range

For i = 1 To n
   t = k * pi * i / n
   x = (R1 - r) * Cos(t) + r0 * Cos(t * (R1 - r) / r)
   y = (R1 - r) * Sin(t) - r0 * Sin(t * (R1 - r) / r)
   x = sc * x
   y = sc * y
   x = InchesToPoints(x)
   y = InchesToPoints(y)
   Set sh = ActiveDocument.Shapes.AddShape _
      (msoShape5pointStar, x, y, sSize, sSize, rng)
Next i

End Sub
Sub ShowWordArtEffects()

Dim sh As Shape
Dim rng As Range
Dim i As Integer

Set rng = Selection.Range
Set sh = ActiveDocument.Shapes.AddTextEffect(msoTextEffect1, _
   "PresetTextEffect xx", "Arial", 24, False, False, _
   0, 0, rng)
For i = msoTextEffect1 To msoTextEffect30
   sh.TextEffect.PresetTextEffect = i
   sh.TextEffect.Text = "PresetTextEffect " & Format(i)
   sh.Visible = True
   Delay 1
Next i

End Sub


