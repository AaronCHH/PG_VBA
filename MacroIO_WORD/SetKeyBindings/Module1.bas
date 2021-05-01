Attribute VB_Name = "Module1"
Sub SetKeyBindings()
  CustomizationContext = NormalTemplate
  
  ' ==== My Macro ====
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="InsHLine", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeySlash)
    
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="cutParEng", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeySemiColon)
  
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="cutParCht", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeySingleQuote)
  
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="JoinParEng", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeySemiColon)
  
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="JoinParCht", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeySingleQuote)
  
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="JoinCutParEng", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeySemiColon)
  
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="JoinCutParCht", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeySingleQuote)
    
  ' ==== Win Macro ====
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="ClearAllFormatting", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyCloseSquareBrace)
    
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="WindowNewWindow", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeySlash)
    
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="PasteTextOnly", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyCloseSquareBrace)
    
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="EditPaste", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyOpenSquareBrace)
    
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="Highlight", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyPeriod)
    
End Sub

Sub AddKeyBindings()
  CustomizationContext = NormalTemplate
  KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="SetKeyBindings", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyW)
End Sub


