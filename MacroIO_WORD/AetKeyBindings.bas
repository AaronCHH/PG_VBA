Attribute VB_Name = "AetKeyBindings"

Sub SetKeyBindings_TEST()
  MsgBox "hello"
End Sub

Sub SetKeyBindings()
  CustomizationContext = NormalTemplate
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="InsHLine", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyA)
End Sub

Sub AddKeyBindings()
  CustomizationContext = NormalTemplate
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="SetKeyBindings", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyW)
End Sub

