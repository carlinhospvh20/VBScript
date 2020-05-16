'SendKeys 
Dim Ws: Set Ws =CreateObject("WScript.Shell") 
Dim sk: sk = InputBox("Keys for simulating insertion:") 
Ws.SendKeys sk 
'SendKeys

'SetMousePosition
Dim Excel: Set Excel = WScript.CreateObject("Excel.Application") 
Dim x: x = InputBox("Coordinates x (Width in pixels):") 
Dim y: y = InputBox("Coordinates y (Height in pixels):")
Excel.ExecuteExcel4Macro "CALL(""user32"",""SetCursorPos"",""JJJ""," & x & "," & y & ")" 
'SetMousePosition