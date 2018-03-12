scriptName = "DiscogsAutoTagWeb"
'Add scripts.ini entries
Dim inip : inip = SDB.ScriptsPath & "Scripts.ini"
Dim inif : Set inif = SDB.Tools.IniFileByPath(inip)
If Not (inif Is Nothing) Then
	inif.StringValue(scriptName,"Filename") = "DiscogsAutoTagWeb.vbs"
	inif.StringValue(scriptName,"Procname") = "DiscogsAutoTagWeb"
	inif.StringValue(scriptName,"Order") = "10"
	inif.StringValue(scriptName,"DisplayName") = "Discogs Tagger"
	inif.StringValue(scriptName,"Description") = "Gets track/album information from discogs.com or musicbrainz.org"
	inif.StringValue(scriptName,"Language") = "VBScript"
	inif.StringValue(scriptName,"ScriptType") = "3"
	SDB.RefreshScriptItems
End If

Dim UI
Set UI = SDB.UI
Set ini = SDB.IniFile
If Not (ini Is Nothing) Then
	If ini.StringValue("DiscogsAutoTagWeb","ReleaseTag") = "" Then
		ini.StringValue("DiscogsAutoTagWeb","ReleaseTag") = "Custom2"
	End If
	If ini.StringValue("DiscogsAutoTagWeb","CatalogTag") = "" Then
		ini.StringValue("DiscogsAutoTagWeb","CatalogTag") = "Custom3"
	End If
	If ini.StringValue("DiscogsAutoTagWeb","CountryTag") = "" Then
		ini.StringValue("DiscogsAutoTagWeb","CountryTag") = "Custom4"
	End If
	If ini.StringValue("DiscogsAutoTagWeb","FormatTag") = "" Then
		ini.StringValue("DiscogsAutoTagWeb","FormatTag") = "Custom5"
	End If
End If

ReleaseTag = ini.StringValue("DiscogsAutoTagWeb","ReleaseTag")
CatalogTag = ini.StringValue("DiscogsAutoTagWeb","CatalogTag")
CountryTag = ini.StringValue("DiscogsAutoTagWeb","CountryTag")
FormatTag = ini.StringValue("DiscogsAutoTagWeb","FormatTag")

CustomField1 = "Custom1 (" & ini.StringValue("CustomFields","Fld1Name") & ")"
CustomField2 = "Custom2 (" & ini.StringValue("CustomFields","Fld2Name") & ")"
CustomField3 = "Custom3 (" & ini.StringValue("CustomFields","Fld3Name") & ")"
CustomField4 = "Custom4 (" & ini.StringValue("CustomFields","Fld4Name") & ")"
CustomField5 = "Custom5 (" & ini.StringValue("CustomFields","Fld5Name") & ")"

Set Form = SDB.UI.NewForm
Form.Common.SetRect 10, 10, 600, 280
Form.Caption = "Please choose the custom tags, where Discogs Tagger save the information"
SDB.Objects("DiscogsOption") = Form

Set Btn = SDB.UI.NewButton(Form)
Btn.Caption = "Close"
Btn.Common.SetRect 10, 10, 100, 20
Btn.UseScript = Script.ScriptPath
REM Btn.OnClickFunc = "OnClose"
Btn.Cancel = true
Btn.Default = true
Btn.ModalResult = 1

Dim DD1, Label1
Set DD1 = UI.NewDropDown(Form)
DD1.Common.SetRect 240, 30, 200, 25
DD1.Style = 2
DD1.AddItem (CustomField1)
DD1.AddItem (CustomField2)
DD1.AddItem (CustomField3)
DD1.AddItem (CustomField4)
DD1.AddItem (CustomField5)
DD1.AddItem ("Don't save")
DD1.AddItem (SDB.Localize("ISRC"))
DD1.AddItem (SDB.Localize("Grouping"))
DD1.AddItem (SDB.Localize("Encoder"))
DD1.AddItem (SDB.Localize("Copyright"))

DD1.Common.ControlName = "ReleaseTag"
If ReleaseTag = "Custom1"Then
	DD1.ItemIndex = 0
End If
If ReleaseTag = "Custom2" Then
	DD1.ItemIndex = 1
End If
If ReleaseTag = "Custom3" Then
	DD1.ItemIndex = 2
End If
If ReleaseTag = "Custom4" Then
	DD1.ItemIndex = 3
End If
If ReleaseTag = "Custom5" Then
	DD1.ItemIndex = 4
End If
If ReleaseTag = "Don't save" Then
	DD1.ItemIndex = 5
End If
If ReleaseTag = "ISRC" Then
	DD1.ItemIndex = 6
End If
If ReleaseTag = "Grouping" Then
	DD1.ItemIndex = 7
End If
If ReleaseTag = "Encoding" Then
	DD1.ItemIndex = 8
End If
If ReleaseTag = "Copyright" Then
	DD1.ItemIndex = 9
End If
Set Label1 = UI.NewLabel(Form)
Label1.Common.SetRect 20, 35, 150, 25
Label1.Caption = "Choose Tag for saving release-number"

Set DD1 = UI.NewDropDown(Form)
DD1.Common.SetRect 240, 60, 200, 25
DD1.Style = 2
DD1.AddItem (CustomField1)
DD1.AddItem (CustomField2)
DD1.AddItem (CustomField3)
DD1.AddItem (CustomField4)
DD1.AddItem (CustomField5)
DD1.AddItem ("Don't save")
DD1.AddItem (SDB.Localize("ISRC"))


DD1.Common.ControlName = "CatalogTag"
If CatalogTag = "Custom1" Then
	DD1.ItemIndex = 0
End If
If CatalogTag = "Custom2" Then
	DD1.ItemIndex = 1
End If
If CatalogTag = "Custom3" Then
	DD1.ItemIndex = 2
End If
If CatalogTag = "Custom4" Then
	DD1.ItemIndex = 3
End If
If CatalogTag = "Custom5" Then
	DD1.ItemIndex = 4
End If
If CatalogTag = "Don't save" Then
	DD1.ItemIndex = 5
End If
If CatalogTag = "ISRC" Then
	DD1.ItemIndex = 6
End If

Set Label1 = UI.NewLabel(Form)
Label1.Common.SetRect 20, 65, 150, 25
Label1.Caption = "Choose Tag for saving catalog number"

Set DD1 = UI.NewDropDown(Form)
DD1.Common.SetRect 240, 90, 200, 25
DD1.Style = 2
DD1.AddItem (CustomField1)
DD1.AddItem (CustomField2)
DD1.AddItem (CustomField3)
DD1.AddItem (CustomField4)
DD1.AddItem (CustomField5)
DD1.AddItem ("Don't save")

DD1.Common.ControlName = "CountryTag"
If CountryTag = "Custom1" Then
	DD1.ItemIndex = 0
End If
If CountryTag = "Custom2" Then
	DD1.ItemIndex = 1
End If
If CountryTag = "Custom3" Then
	DD1.ItemIndex = 2
End If
If CountryTag = "Custom4" Then
	DD1.ItemIndex = 3
End If
If CountryTag = "Custom5" Then
	DD1.ItemIndex = 4
End If
If CountryTag = "Don't save" Then
	DD1.ItemIndex = 5
End If

Set Label1 = UI.NewLabel(Form)
Label1.Common.SetRect 20, 95, 150, 25
Label1.Caption = "Choose Tag for saving release country"

Set DD1 = UI.NewDropDown(Form)
DD1.Common.SetRect 240, 120, 200, 25
DD1.Style = 2
DD1.AddItem (CustomField1)
DD1.AddItem (CustomField2)
DD1.AddItem (CustomField3)
DD1.AddItem (CustomField4)
DD1.AddItem (CustomField5)
DD1.AddItem ("Don't save")

DD1.Common.ControlName = "FormatTag"
If FormatTag = "Custom1" Then
	DD1.ItemIndex = 0
End If
If FormatTag = "Custom2" Then
	DD1.ItemIndex = 1
End If
If FormatTag = "Custom3" Then
	DD1.ItemIndex = 2
End If
If FormatTag = "Custom4" Then
	DD1.ItemIndex = 3
End If
If FormatTag = "Custom5" Then
	DD1.ItemIndex = 4
End If
If FormatTag = "Don't save" Then
	DD1.ItemIndex = 5
End If

Set Label1 = UI.NewLabel(Form)
Label1.Common.SetRect 20, 125, 150, 25
Label1.Caption = "Choose Tag for saving media format"


If Form.ShowModal = 1 Then
	Dim edt
	Set edt = Form.Common.ChildControl("ReleaseTag")
	ini.StringValue("DiscogsAutoTagWeb", "ReleaseTag") = GetCustom(edt.ItemIndex)
	Set edt = Form.Common.ChildControl("CatalogTag")
	ini.StringValue("DiscogsAutoTagWeb", "CatalogTag") = GetCustom(edt.ItemIndex)
	Set edt = Form.Common.ChildControl("CountryTag")
	ini.StringValue("DiscogsAutoTagWeb", "CountryTag") = GetCustom(edt.ItemIndex)
	Set edt = Form.Common.ChildControl("FormatTag")
	ini.StringValue("DiscogsAutoTagWeb", "FormatTag") = GetCustom(edt.ItemIndex)
	SDB.Objects("DiscogsOption") = Nothing
End If

REM Sub OnClose(Btn)
	REM Set Form = SDB.Objects("DiscogsOption")
	REM SaveOptions(Form)
	REM SDB.Objects("DiscogsOption") = Nothing
REM End Sub

Function GetCustom(Index)

	If Index = 0 Then GetCustom = "Custom1"
	If Index = 1 Then GetCustom = "Custom2"
	If Index = 2 Then GetCustom = "Custom3"
	If Index = 3 Then GetCustom = "Custom4"
	If Index = 4 Then GetCustom = "Custom5"
	If Index = 5 Then GetCustom = "Don't save"
	If Index = 6 Then GetCustom = "ISRC"
	If Index = 7 Then GetCustom = "Grouping"
	If Index = 8 Then GetCustom = "Encoder"
	If Index = 9 Then GetCustom = "Copyright"

End Function
