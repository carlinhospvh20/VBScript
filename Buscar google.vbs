set startmundo =createobject("wscript.shell") 
 startmundo.run ("www.google.com")
   wscript.sleep 300
   wscript.sleep 300
 
'SendKeys 
 ws = InputBox("Qual é a definição:") 

'SendKeys
 startmundo.sendkeys " define: "
 startmundo.sendkeys  ws
   wscript.sleep 300
   wscript.sleep 300
 startmundo.sendkeys "{enter}"
   wscript.sleep 300

startmundo.run ("www.google.com")
   wscript.sleep 300
   wscript.sleep 300
 
'SendKeys 
 ws = InputBox("Buscar pelo Título:") 

'SendKeys
 startmundo.sendkeys " allintitle: "
 startmundo.sendkeys  ws
   wscript.sleep 300
   wscript.sleep 300
 startmundo.sendkeys "{enter}"
   wscript.sleep 300
   
   startmundo.run ("www.google.com")
   wscript.sleep 300
   wscript.sleep 300
 
'SendKeys 
 ws = InputBox("""Nome do material:""") 
 extensao = InputBox("Extensao do arquivo:") 

'SendKeys
 startmundo.sendkeys  ws
 startmundo.sendkeys " filetype: "
 startmundo.sendkeys  extensao
 
   wscript.sleep 300
   wscript.sleep 300
 startmundo.sendkeys "{enter}"
   wscript.sleep 300