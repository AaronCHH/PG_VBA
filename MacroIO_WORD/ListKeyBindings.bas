Attribute VB_Name = "ListKeyBindings"

Option Explicit

Public Sub ListKeyBindings()

Dim kb As KeyBinding
Dim s As String
Dim tbl As Table

' Check active document for text
' and warn user.
If ActiveDocument.Content <> vbCr Then
   If MsgBox("Active doc has content. Proceed?", _
      vbQuestion + vbYesNo) = vbNo Then Exit Sub
End If

' Print Heading
Selection.InsertAfter KeyBindings.Count & _
   " key bindings in context: " & CustomizationContext _
   & vbCr & vbCr

' Collapse selection to end of document
Selection.Collapse wdCollapseEnd

' Insert start of table bookmark
ActiveDocument.Bookmarks.Add "StartOfTable"

' Print table heading
Selection.InsertAfter "KeyString" & vbTab & _
   "KeyCategory" & vbTab & "Command" & vbTab _
   & "KeyCode" & vbTab & "KeyCode2" _
  & vbTab & "CommandParameter" & vbCr

'Start the For loop, printing key binding data
Selection.Collapse wdCollapseEnd
For Each kb In KeyBindings
   s = kb.KeyString & vbTab & kb.KeyCategory _
      & vbTab & kb.Command & vbTab & kb.KeyCode _
      & vbTab & kb.KeyCode2 & vbTab _
      & kb.CommandParameter & vbCr
   Selection.InsertAfter s
   Selection.Collapse wdCollapseEnd
Next kb

' Collapse selection to end of document
Selection.Collapse wdCollapseEnd

' Insert end of table bookmark
ActiveDocument.Bookmarks.Add "EndOfTable"

' Select text between bookmarks
ActiveDocument.Bookmarks("StartofTable").Select
With Selection
    .ExtendMode = True
.GoTo wdGoToBookmark, , , "EndOfTable"
    .ExtendMode = False
End With

Set tbl = _
   Selection.ConvertToTable(Separator:=wdSeparateByTabs)
tbl.Columns.AutoFit
Selection.Collapse wdCollapseEnd

End Sub