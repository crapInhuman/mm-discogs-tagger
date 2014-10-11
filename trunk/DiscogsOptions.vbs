'
' MediaMonkey Script
'
' NAME: Discogs Tagger Options 1.4
'
' AUTHOR: crap_inhuman
' DATE : 14/11/2013
'
'
' INSTALL: Automatic installation during Discogs Tagger install
'
'Changes from 1.3 to 1.4
'Changed the separator from '|' to ','
'
'Changes from 1.2 to 1.3
'Added 'Don't save' and 4 more fields for saving release-number
'
'Changes from 1.1 to 1.2
'Added 3 new options
'
'Changes from 1.0 to 1.1
'Added option for changing keywords
'




Sub OnStartup

	DiscogsOptions = SDB.UI.AddOptionSheet( "Discogs Tagger", Script.ScriptPath, "InitSheet", "SaveSheet", -3)

End Sub

Sub InitSheet(Sheet)

	Dim UI : Set UI = SDB.UI
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

		If ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords") = "Lyrics By,Words By"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords") = "Conductor"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords") = "Producer,Arranged By,Recorded By"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords") = "Composed By,Score,Written-By,Written By,Music By,Programmed By,Songwriter"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckNotAlwaysSaveimage") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckNotAlwaysSaveimage") = false
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = true
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckStyleField") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","CheckStyleField") = "Default (stored with Genre)"
		End If


		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords"), "|", ",")
		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords"), "|", ",")
		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords"), "|", ",")
		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords"), "|", ",")
	End If


	ReleaseTag = ini.StringValue("DiscogsAutoTagWeb","ReleaseTag")
	CatalogTag = ini.StringValue("DiscogsAutoTagWeb","CatalogTag")
	CountryTag = ini.StringValue("DiscogsAutoTagWeb","CountryTag")
	FormatTag = ini.StringValue("DiscogsAutoTagWeb","FormatTag")
	LyricistKeywords = ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords")
	ConductorKeywords = ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords")
	ProducerKeywords = ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords")
	ComposerKeywords = ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords")
	CheckNotAlwaysSaveimage = ini.BoolValue("DiscogsAutoTagWeb","CheckNotAlwaysSaveimage")
	CheckOriginalDiscogsTrack = ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack")
	CheckStyleField = ini.StringValue("DiscogsAutoTagWeb","CheckStyleField")

	CustomField1 = "Custom1 (" & ini.StringValue("CustomFields","Fld1Name") & ")"
	CustomField2 = "Custom2 (" & ini.StringValue("CustomFields","Fld2Name") & ")"
	CustomField3 = "Custom3 (" & ini.StringValue("CustomFields","Fld3Name") & ")"
	CustomField4 = "Custom4 (" & ini.StringValue("CustomFields","Fld4Name") & ")"
	CustomField5 = "Custom5 (" & ini.StringValue("CustomFields","Fld5Name") & ")"
	
	
	Dim GroupBox0
	Set GroupBox0 = UI.NewGroupBox(Sheet)
	GroupBox0.Caption = "Please choose the custom fields, where Discogs Tagger save the information"
	GroupBox0.Common.SetRect 10, 10, 500, 190

	Dim Label1
	Set Label1 = UI.NewLabel(GroupBox0)
	Label1.Common.SetRect 65, 15, 150, 25
	Label1.Caption = "Don't choose a Custom Field more than once !!"
	Dim DD1
	Set DD1 = UI.NewDropDown(GroupBox0)
	DD1.Common.SetRect 240, 40, 200, 25
	DD1.Style = 2
	DD1.AddItem (CustomField1)
	DD1.AddItem (CustomField2)
	DD1.AddItem (CustomField3)
	DD1.AddItem (CustomField4)
	DD1.AddItem (CustomField5)
	DD1.AddItem ("Don't save")
	DD1.AddItem (SDB.Localize("Grouping"))
	DD1.AddItem (SDB.Localize("ISRC"))
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
	If ReleaseTag = "Grouping" Then
		DD1.ItemIndex = 6
	End If
	If ReleaseTag = "ISRC" Then
		DD1.ItemIndex = 7
	End If
	If ReleaseTag = "Encoding" Then
		DD1.ItemIndex = 8
	End If
	If ReleaseTag = "Copyright" Then
		DD1.ItemIndex = 9
	End If
	Set Label1 = UI.NewLabel(GroupBox0)
	Label1.Common.SetRect 20, 45, 150, 25
	Label1.Caption = "Choose field for saving release-number"
	Label1.Common.Hint = "In brackets you see the name you chose for the custom tag"

	Dim DD2
	Set DD2 = UI.NewDropDown(GroupBox0)
	DD2.Common.SetRect 240, 70, 200, 25
	DD2.Style = 2
	DD2.AddItem (CustomField1)
	DD2.AddItem (CustomField2)
	DD2.AddItem (CustomField3)
	DD2.AddItem (CustomField4)
	DD2.AddItem (CustomField5)
	DD2.AddItem ("Don't save")

	DD2.Common.ControlName = "CatalogTag"
	If CatalogTag = "Custom1" Then
		DD2.ItemIndex = 0
	End If
	If CatalogTag = "Custom2" Then
		DD2.ItemIndex = 1
	End If
	If CatalogTag = "Custom3" Then
		DD2.ItemIndex = 2
	End If
	If CatalogTag = "Custom4" Then
		DD2.ItemIndex = 3
	End If
	If CatalogTag = "Custom5" Then
		DD2.ItemIndex = 4
	End If
	If CatalogTag = "Don't save" Then
		DD2.ItemIndex = 5
	End If
	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 75, 150, 25
	Label2.Caption = "Choose field for saving catalog number"
	Label2.Common.Hint = "In brackets you see the name you chose for the custom tag"

	Set DD2 = UI.NewDropDown(GroupBox0)
	DD2.Common.SetRect 240, 100, 200, 25
	DD2.Style = 2
	DD2.AddItem (CustomField1)
	DD2.AddItem (CustomField2)
	DD2.AddItem (CustomField3)
	DD2.AddItem (CustomField4)
	DD2.AddItem (CustomField5)
	DD2.AddItem ("Don't save")

	DD2.Common.ControlName = "CountryTag"
	If CountryTag = "Custom1" Then
		DD2.ItemIndex = 0
	End If
	If CountryTag = "Custom2" Then
		DD2.ItemIndex = 1
	End If
	If CountryTag = "Custom3" Then
		DD2.ItemIndex = 2
	End If
	If CountryTag = "Custom4" Then
		DD2.ItemIndex = 3
	End If
	If CountryTag = "Custom5" Then
		DD2.ItemIndex = 4
	End If
	If CountryTag = "Don't save" Then
		DD2.ItemIndex = 5
	End If

	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 105, 150, 25
	Label2.Caption = "Choose field for saving release country"
	Label2.Common.Hint = "In brackets you see the name you chose for the custom tag"

	Set DD2 = UI.NewDropDown(GroupBox0)
	DD2.Common.SetRect 240, 130, 200, 25
	DD2.Style = 2
	DD2.AddItem (CustomField1)
	DD2.AddItem (CustomField2)
	DD2.AddItem (CustomField3)
	DD2.AddItem (CustomField4)
	DD2.AddItem (CustomField5)
	DD2.AddItem ("Don't save")

	DD2.Common.ControlName = "FormatTag"
	If FormatTag = "Custom1" Then
		DD2.ItemIndex = 0
	End If
	If FormatTag = "Custom2" Then
		DD2.ItemIndex = 1
	End If
	If FormatTag = "Custom3" Then
		DD2.ItemIndex = 2
	End If
	If FormatTag = "Custom4" Then
		DD2.ItemIndex = 3
	End If
	If FormatTag = "Custom5" Then
		DD2.ItemIndex = 4
	End If
	If FormatTag = "Don't save" Then
		DD2.ItemIndex = 5
	End If

	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 135, 150, 25
	Label2.Caption = "Choose field for saving media format"
	Label2.Common.Hint = "In brackets you see the name you chose for the custom tag"

	Dim Combo
	Set Combo = UI.NewDropDown(GroupBox0)
	Combo.Common.SetRect 240, 160, 200, 25
	Combo.Style = 2     ' List
	Combo.Common.ControlName = "CheckStyleField"

	Combo.AddItem ("Default (stored with Genre)")
	Combo.AddItem (CustomField1)
	Combo.AddItem (CustomField2)
	Combo.AddItem (CustomField3)
	Combo.AddItem (CustomField4)
	Combo.AddItem (CustomField5)

	If CheckStyleField = "Default (stored with Genre)" Then
		Combo.ItemIndex = 0
	End If
	If CheckStyleField = "Custom1" Then
		Combo.ItemIndex = 1
	End If
	If CheckStyleField = "Custom2" Then
		Combo.ItemIndex = 2
	End If
	If CheckStyleField = "Custom3" Then
		Combo.ItemIndex = 3
	End If
	If CheckStyleField = "Custom4" Then
		Combo.ItemIndex = 4
	End If
	If CheckStyleField = "Custom5" Then
		Combo.ItemIndex = 5
	End If

	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 165, 50, 25
	Label2.Caption = "Choose field for saving Style"
	Label2.Common.Hint = "In brackets you see the name you chose for the custom tag"

	Dim tmp, editText, x

	Dim GroupBox1
	Set GroupBox1 = UI.NewGroupBox(Sheet)
	GroupBox1.Caption = "Enter the keywords for linking with discogs"
	GroupBox1.Common.Hint = "If you don't know what to enter here, let the keywords as is !!"
	GroupBox1.Common.SetRect 10, 210, 500, 190


	Set Label2 = UI.NewLabel(GroupBox1)
	Label2.Common.SetRect 20, 20, 50, 25
	Label2.Caption = SDB.Localize("Lyricist")
	Set EditLyricist = UI.NewEdit(GroupBox1)
	EditLyricist.Common.SetRect 20, 35, 450, 35
	EditLyricist.Common.ControlName = "LyricistKeywords"
	EditLyricist.Text = LyricistKeywords


	Set Label2 = UI.NewLabel(GroupBox1)
	Label2.Common.SetRect 20, 60, 50, 25
	Label2.Caption = SDB.Localize("Conductor")
	Set EditConductor = UI.NewEdit(GroupBox1)
	EditConductor.Common.SetRect 20, 75, 450, 35
	EditConductor.Common.ControlName = "ConductorKeywords"
	EditConductor.Text = ConductorKeywords


	Set Label2 = UI.NewLabel(GroupBox1)
	Label2.Common.SetRect 20, 100, 50, 25
	Label2.Caption = SDB.Localize("Producer")
	Set EditProducer = UI.NewEdit(GroupBox1)
	EditProducer.Common.SetRect 20, 115, 450, 35
	EditProducer.Common.ControlName = "ProducerKeywords"
	EditProducer.Text = ProducerKeywords


	Set Label2 = UI.NewLabel(GroupBox1)
	Label2.Common.SetRect 20, 140, 50, 25
	Label2.Caption = SDB.Localize("Composer")
	Set EditComposer = UI.NewEdit(GroupBox1)
	EditComposer.Common.SetRect 20, 155, 450, 35
	EditComposer.Common.ControlName = "ComposerKeywords"
	EditComposer.Text = ComposerKeywords

	Set Label2 = UI.NewLabel(Sheet)
	Label2.Common.SetRect 40, 410, 50, 25
	Label2.Caption = "Check 'Save Image' Checkbox only if release have no image"

	Dim Checkbox1
	Set Checkbox1 = UI.NewCheckBox(Sheet)
	Checkbox1.Common.SetRect 20, 410, 15, 15
	Checkbox1.Common.ControlName = "NotAlwaysSaveimage"
	If CheckNotAlwaysSaveimage = true Then Checkbox1.Checked = true

	Set Label2 = UI.NewLabel(Sheet)
	Label2.Common.SetRect 40, 440, 50, 25
	Label2.Caption = "Show the original Discogs track position"

	Dim Checkbox2
	Set Checkbox2 = UI.NewCheckBox(Sheet)
	Checkbox2.Common.SetRect 20, 440, 15, 15
	Checkbox2.Common.ControlName = "CheckOriginalDiscogsTrack"
	If CheckOriginalDiscogsTrack = true Then Checkbox2.Checked = true

End Sub

Sub SaveSheet(Sheet)

	Dim ini
	Set ini = SDB.IniFile
	Dim edt
	Set edt = Sheet.Common.ChildControl("ReleaseTag")
	ini.StringValue("DiscogsAutoTagWeb", "ReleaseTag") = GetCustom(edt.ItemIndex)
	Set edt = Sheet.Common.ChildControl("CatalogTag")
	ini.StringValue("DiscogsAutoTagWeb", "CatalogTag") = GetCustom(edt.ItemIndex)
	Set edt = Sheet.Common.ChildControl("CountryTag")
	ini.StringValue("DiscogsAutoTagWeb", "CountryTag") = GetCustom(edt.ItemIndex)
	Set edt = Sheet.Common.ChildControl("FormatTag")
	ini.StringValue("DiscogsAutoTagWeb", "FormatTag") = GetCustom(edt.ItemIndex)
	Set edt = Sheet.Common.ChildControl("CheckStyleField")
	ini.StringValue("DiscogsAutoTagWeb", "CheckStyleField") = GetCustom(edt.ItemIndex -1)

	Dim tmp, x, editText
	Set edt = Sheet.Common.ChildControl("LyricistKeywords")
	tmp = Split(edt.Text, ",")
	editText = ""
	For each x in tmp
		editText = editText & Trim(x) & ","
	Next
	ini.StringValue("DiscogsAutoTagWeb", "LyricistKeywords") = Left(editText, Len(editText)-1)

	Set edt = Sheet.Common.ChildControl("ConductorKeywords")
	tmp = Split(edt.Text, ",")
	editText = ""
	For each x in tmp
		editText = editText & Trim(x) & ","
	Next
	ini.StringValue("DiscogsAutoTagWeb", "ConductorKeywords") = Left(editText, Len(editText)-1)

	Set edt = Sheet.Common.ChildControl("ProducerKeywords")
	tmp = Split(edt.Text, ",")
	editText = ""
	For each x in tmp
		editText = editText & Trim(x) & ","
	Next
	ini.StringValue("DiscogsAutoTagWeb", "ProducerKeywords") = Left(editText, Len(editText)-1)

	Set edt = Sheet.Common.ChildControl("ComposerKeywords")
	tmp = Split(edt.Text, ",")
	editText = ""
	For each x in tmp
		editText = editText & Trim(x) & ","
	Next
	ini.StringValue("DiscogsAutoTagWeb", "ComposerKeywords") = Left(editText, Len(editText)-1)

	Set checkbox = Sheet.Common.ChildControl("NotAlwaysSaveimage")
	If checkbox.checked Then
		ini.BoolValue("DiscogsAutoTagWeb", "CheckNotAlwaysSaveimage") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "CheckNotAlwaysSaveimage") = false
	End If

	Set checkbox = Sheet.Common.ChildControl("CheckOriginalDiscogsTrack")
	If checkbox.checked Then
		ini.BoolValue("DiscogsAutoTagWeb", "CheckOriginalDiscogsTrack") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "CheckOriginalDiscogsTrack") = false
	End If

End Sub

Function GetCustom(Index)

	Dim ini : Set ini = SDB.IniFile
	If Index = 0 Then GetCustom = "Custom1"
	If Index = 1 Then GetCustom = "Custom2"
	If Index = 2 Then GetCustom = "Custom3"
	If Index = 3 Then GetCustom = "Custom4"
	If Index = 4 Then GetCustom = "Custom5"
	If Index = 5 Then GetCustom = "Don't save"
	If Index = -1 Then GetCustom = "Default (stored with Genre)"
	If Index = 6 Then GetCustom = "Grouping"
	If Index = 7 Then GetCustom = "ISRC"
	If Index = 8 Then GetCustom = "Encoder"
	If Index = 9 Then GetCustom = "Copyright"

End Function
