myName = "DiscogsAutoTagWeb"
iniSec = "DiscogsAutoTagWeb"

MsgDeleteSettings = "Do you want to remove " & myName & " settings as well?" & vbNewLine & _
					"If you click No, script settings will be left in MediaMonkey.ini"

If (Not (SDB.IniFile Is Nothing)) and (MsgBox(MsgDeleteSettings, vbYesNo) = vbYes) Then
	SDB.IniFile.DeleteSection(iniSec)
End If

Dim inip
inip = SDB.ApplicationPath&"Scripts\Scripts.ini"
Dim inif
Set inif = SDB.Tools.IniFileByPath(inip)
If Not (inif Is Nothing) Then
	inif.DeleteSection(iniSec)
	SDB.RefreshScriptItems
End If

