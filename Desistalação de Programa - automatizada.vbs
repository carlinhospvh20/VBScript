msgbox "Desistalar um programa!",48,ERROR
set Wshshell = wscript.Createobject("WScript.Shell")

'Inicializar o CMD
Wshshell.Run "wmic"
WScript.sleep 500

Wshshell.sendkeys "product get name, version, installlocation, installdate"
WScript.sleep 500

Wshshell.sendkeys "{ENTER}"
WScript.sleep 100

set wsite=createobject("wscript.shell")
RecebeValor=inputbox("Informe  nome do programa", "Desistalador WMIC", "Cole aqui o nome do Programa", 0, 3000)

Wshshell.sendkeys "product where name=""" & RecebeValor & """ call uninstall /nointeractive"
'WScript.sleep 500

'Wshshell.sendkeys "{ENTER}"
'WScript.sleep 100



