Dim inip, inif, scriptName

Set inif = SDB.IniFile
scriptname = inif.StringValue( "AlbumBrowser","RunningScriptName")
If scriptname = "DiscogsAutoTagWeb.vbs" Then
	inif.StringValue("AlbumBrowser","RunningScriptName") = "DiscogsWebTag.vbs"
End If
If inif.StringValue("DiscogsAutoTagWeb","AccessToken") = "" Then
	inif.StringValue("AlbumBrowser","RunningScriptName") = "MusicBrainzWebTag.vbs"
End If
inif.Apply


inip = SDB.ApplicationPath&"Scripts\Scripts.ini"
Set inif = SDB.Tools.IniFileByPath(inip)

If Not (inif Is Nothing) Then
	iniSec = "DiscogsAutoTagWeb"
	inif.DeleteSection(iniSec)
	iniSec = "MusicbrainzAutoTagWeb"
	inif.DeleteSection(iniSec)
	SDB.RefreshScriptItems
End If

inip = SDB.ScriptsPath & "Scripts.ini"
Set inif = SDB.Tools.IniFileByPath(inip)

If Not (inif Is Nothing) Then
	iniSec = "DiscogsAutoTagWeb"
	inif.DeleteSection(iniSec)
	iniSec = "MusicbrainzAutoTagWeb"
	inif.DeleteSection(iniSec)
	SDB.RefreshScriptItems
End If

If Not (inif Is Nothing) Then
	scriptName = "DiscogsAutoTagWeb"
	inif.StringValue(scriptName,"Filename") = "DiscogsWebTag.vbs"
	inif.StringValue(scriptName,"Procname") = "DiscogsWebTag"
	inif.StringValue(scriptName,"Order") = "10"
	inif.StringValue(scriptName,"DisplayName") = "Discogs Tagger"
	inif.StringValue(scriptName,"Description") = "Gets track/album information from discogs.com"
	inif.StringValue(scriptName,"Language") = "VBScript"
	inif.StringValue(scriptName,"ScriptType") = "3"
	scriptName = "MusicbrainzAutoTagWeb"
	inif.StringValue(scriptName,"Filename") = "MusicBrainzWebTag.vbs"
	inif.StringValue(scriptName,"Procname") = "MusicBrainzWebTag"
	inif.StringValue(scriptName,"Order") = "10"
	inif.StringValue(scriptName,"DisplayName") = "MusicBrainz Tagger"
	inif.StringValue(scriptName,"Description") = "Gets track/album information from musicbrainz.org"
	inif.StringValue(scriptName,"Language") = "VBScript"
	inif.StringValue(scriptName,"ScriptType") = "3"
	SDB.RefreshScriptItems
End If
