'=========================================================================
' KeyBindings.vbs
'
' Sample script to use the Key Binding Server
' 
' This script is intended for use as a startup script. See the section
' "Using MGCPCB Automation"->"Using startup scripts" in the MGCPCB 
' Automatio Help for details.
'
'=========================================================================


' Modifiers exist for key combinations such as:
'   <key>
'   Alt+<key>
'   Ctrl+<key>
'   Alt+Shift+<key>
'   Ctrl+Shift+<key>
'   Ctrl+Alt+Shift+<key>
' where <key> is the keyname such as: Enter, Space, A-Z, 0-9, F1-F12,
'
' One exception: Do not override any of the existing function keys (F1-F12) without
' using a modifier or that functionality will be disabled in the product.
'
BindMenu = 0
BindAccelerator = 1
BindCommand = 1

Set keyBindTables = Gui.Bindings("Document")

' Bind the execution of a user script to a key
keyBindTables.AddKeyBinding "L", "run C:\Users\xinz\Documents\MyScripts\toggle_all_layers.vbs", BindCommand, BindAccelerator
keyBindTables.AddKeyBinding "B", "run C:\Users\xinz\Documents\MyScripts\toggle_plane.vbs", BindCommand, BindAccelerator
keyBindTables.AddKeyBinding "I", "run C:\Users\xinz\Documents\MyScripts\toggle_top_bottom.vbs", BindCommand, BindAccelerator
keyBindTables.AddKeyBinding "N", "run C:\Users\xinz\Documents\MyScripts\toggle_netline.vbs", BindCommand, BindAccelerator
keyBindTables.AddKeyBinding "S", "run C:\Users\xinz\Documents\MyScripts\toggle_shadow.vbs", BindCommand, BindAccelerator
