# Functions

Add the following function which returns the area. Note that a value/values can be returned with the function name itself.

__Example__
```{vb}
Function findArea(Length As Double, Optional Width As Variant)
   If IsMissing(Width) Then
      findArea = Length * Length
   Else
      findArea = Length * Width
   End If
End Function
```

## Reference
* Tutorialspoints https://www.tutorialspoint.com/vba/vba_functions.htm