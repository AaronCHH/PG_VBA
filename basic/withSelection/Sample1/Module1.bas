Attribute VB_Name = "Module1"
Sub FormatCell()
  ' FormatCell Macro
  ' Keyboard Shortcut: Ctrl+Shift+F
  Selection.Font.Bold = True
  Selection.Font.ltalic = False

  With Selection.Font
    .Name = "Arial Black"
    .Size = 11
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontNone
  End With
  
  ActiveCell.Columns("A:A").EntireColumn _
    .EntireColumn.AutoFit
  
End Sub
