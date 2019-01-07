FIXO

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize

CABEÇALHO SAP

MARCAR O ITEM ZERO E EXPANDIR TODA A ÁRVORE

ZERAR TODA A ÁRVORE

PREENCHER


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem "       2998","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "       2998","BOX"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").changeCheckbox "       2998","BOX",true


session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell")

selectItem
ensureVisibleHorizontalItem
changeCheckbox
"       2998","BOX",true (or false)
"..........."

Premissa:
- Toda vez que um flag é ativado:

	selectItem 				+ endereço + "BOX"
	endureVisibleHorizontalItem		+ endereço + "BOX"
	changeCheckbox				+ endereço + "BOX",condição
