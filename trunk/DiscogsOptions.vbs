'
' MediaMonkey Script
'
' NAME: Discogs Tagger Options 1.9
'
' AUTHOR: crap_inhuman
' DATE : 25/03/2014
'
'
' INSTALL: Automatic installation during Discogs Tagger install
'
'Changes from 1.8 to 1.9
'Added metal-archives.com for release search instead of discogs(BETA)

'Changes from 1.7 to 1.8
'Split the options in 2 parts
'
'Changes from 1.6 to 1.7
'Added the option for switching the last artist separator ("&" or "chosen separator")
'
'Changes from 1.5 to 1.6
'Added the option for changing the artist separator
'
'Changes from 1.4 to 1.5
'Added the "featuring" keywords
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
	Call SDB.UI.AddOptionSheet("Keywords",Script.ScriptPath,"InitSheet2","SaveSheet2", DiscogsOptions)

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

		If ini.StringValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = true
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckStyleField") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","CheckStyleField") = "Default (stored with Genre)"
		End If

		If ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = ", "
		End If
		If ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = "" Then
			If ini.ValueExists("DiscogsAutoTagWeb","CheckNotAlwaysSaveimage") Then
				If ini.BoolValue("DiscogsAutoTagWeb","CheckNotAlwaysSaveimage") = false Then
					ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = 0
				Else
					ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = 1
				End If
				ini.DeleteKey "DiscogsAutoTagWeb","CheckNotAlwaysSaveimage"
				
			Else
				ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = 1
			End If
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSmallCover") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckSmallCover") = False
		End If
		If ini.ValueExists("DiscogsAutoTagWeb","CheckCover") Then
			ini.DeleteKey "DiscogsAutoTagWeb","CheckCover"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","UseMetalArchives") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","UseMetalArchives") = False
		End If
	End If


	ReleaseTag = ini.StringValue("DiscogsAutoTagWeb","ReleaseTag")
	CatalogTag = ini.StringValue("DiscogsAutoTagWeb","CatalogTag")
	CountryTag = ini.StringValue("DiscogsAutoTagWeb","CountryTag")
	FormatTag = ini.StringValue("DiscogsAutoTagWeb","FormatTag")
	REM CheckNotAlwaysSaveimage = ini.BoolValue("DiscogsAutoTagWeb","CheckNotAlwaysSaveimage")
	CheckOriginalDiscogsTrack = ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack")
	CheckStyleField = ini.StringValue("DiscogsAutoTagWeb","CheckStyleField")
	ArtistSeparator = ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator")
	ArtistLastSeparator = ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator")
	CheckSaveImage = ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage")			'0 = Always save - 1 = Only when no image found - 2 = always don't save
	CheckSmallCover = ini.BoolValue("DiscogsAutoTagWeb","CheckSmallCover")
	UseMetalArchives = ini.BoolValue("DiscogsAutoTagWeb","UseMetalArchives")

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

	Dim GroupBox1
	Set GroupBox1 = UI.NewGroupBox(Sheet)
	GroupBox1.Caption = "Cover-Images"
	GroupBox1.Common.SetRect 10, 210, 500, 80

	Dim Checkbox1
	Set Checkbox1 = UI.NewCheckBox(GroupBox1)
	Checkbox1.Common.SetRect 20, 20, 250, 15
	Checkbox1.Common.ControlName = "ControlSaveImage1"
	Checkbox1.Caption = "Always set option for saving Cover-Images"
	Checkbox1.Common.Hint = "The script always set the option to save the Cover-Image."
	Set SDB.Objects("CoverSaveOn") = Checkbox1
	If CheckSaveImage = 0 or CheckSaveImage = 1 Then
		Checkbox1.checked = True
	Else
		Checkbox1.checked = False
	End If

	Dim Checkbox12
	Set Checkbox12 = UI.NewCheckBox(GroupBox1)
	Checkbox12.Common.SetRect 40, 40, 250, 15
	Checkbox12.Common.ControlName = "ControlSaveImage12"
	Checkbox12.Caption = "Only if no image already exists"
	Checkbox12.Common.Hint = "If option set the script only mark covers for save when no image already exists."
	Set SDB.Objects("CoverSaveIfEmpty") = Checkbox12
	If CheckSaveImage = 0 Then
		Checkbox12.checked = False
		Checkbox12.Common.Enabled = True
	ElseIf CheckSaveImage = 1 Then
		Checkbox12.checked = True
		Checkbox12.Common.Enabled = True
	Else
		Checkbox12.checked = False
		Checkbox12.Common.Enabled = False
	End If

	Dim Checkbox13
	Set Checkbox13 = UI.NewCheckBox(GroupBox1)
	Checkbox13.Common.SetRect 40, 60, 250, 15
	Checkbox13.Common.ControlName = "ControlSaveImage13"
	Checkbox13.Caption = "Small Cover (150x150)"
	Checkbox13.Common.Hint = "If option not set the script get the large cover images."
	Set SDB.Objects("SmallCoverSave") = Checkbox13
	If CheckSmallCover = False Then
		Checkbox13.checked = False
	Else
		Checkbox13.checked = True
	End If
	If CheckSaveImage = 0 or CheckSaveImage = 1 Then
		Checkbox13.Common.Enabled = True
	Else
		Checkbox13.Common.Enabled = False
	End If
	

	Script.RegisterEvent Checkbox1.Common, "OnClick", "ChBClick"

	Dim GroupBox2
	Set GroupBox2 = UI.NewGroupBox(Sheet)
	GroupBox2.Caption = "Misc"
	GroupBox2.Common.SetRect 10, 300, 500, 100


	Set Label2 = UI.NewLabel(GroupBox2)
	Label2.Common.SetRect 40, 20, 50, 25
	Label2.Caption = "Show the original Discogs track position"

	Dim Checkbox2
	Set Checkbox2 = UI.NewCheckBox(GroupBox2)
	Checkbox2.Common.SetRect 20, 20, 15, 15
	Checkbox2.Common.ControlName = "CheckOriginalDiscogsTrack"
	If CheckOriginalDiscogsTrack = true Then Checkbox2.Checked = true

	Set Label2 = UI.NewLabel(GroupBox2)
	Label2.Common.SetRect 20, 50, 50, 25
	Label2.Caption = SDB.Localize("Artist Separator")
	Label2.Common.Hint = "Standard is ', ' without apostrophe"

	Set EditArtistSep = UI.NewEdit(GroupBox2)
	EditArtistSep.Common.SetRect 20, 65, 50, 35
	EditArtistSep.Common.ControlName = "ArtistSeparator"
	EditArtistSep.Text = ArtistSeparator
	EditArtistSep.Common.Hint = "Standard is ', ' without apostrophe"

	Set Label2 = UI.NewLabel(GroupBox2)
	Label2.Common.SetRect 165, 67, 125, 25
	Label2.Caption = "Artist Last Separator = &&"
	Label2.Common.Hint = "If checked artist list will be Artist1" & ArtistSeparator & "Artist2 & Artist3" & vbCrLf & "If not checked it will be Artist1" & ArtistSeparator & "Artist2" & ArtistSeparator & "Artist3"

	Dim Checkbox3
	Set Checkbox3 = UI.NewCheckBox(GroupBox2)
	Checkbox3.Common.SetRect 145, 67, 15, 15
	Checkbox3.Common.ControlName = "EditArtistLastSep"
	Checkbox3.Common.Hint = "If checked artist list will be Artist1" & ArtistSeparator & "Artist2 & Artist3" & vbCrLf & "If not checked it will be Artist1" & ArtistSeparator & "Artist2" & ArtistSeparator & "Artist3"
	If ArtistLastSeparator = true Then Checkbox3.Checked = true

	Dim GroupBox3
	Set GroupBox3 = UI.NewGroupBox(Sheet)
	GroupBox3.Caption = "BETA"
	GroupBox3.Common.SetRect 10, 410, 500, 45

	Dim Checkbox4
	Set Checkbox4 = UI.NewCheckBox(GroupBox3)
	Checkbox4.Common.SetRect 20, 20, 240, 15
	Checkbox4.Common.ControlName = "UseMetalArchives"
	CheckBox4.Caption = "Using Metal-Archives.com instead of Discogs"
	Checkbox4.Common.Hint = "At the moment only the Title will be compared"
	If UseMetalArchives = true Then Checkbox4.Checked = true

End Sub

Sub ChBClick(CheckBox1)

	Set CB1 = SDB.Objects("CoverSaveOn")
	Set CB12 = SDB.Objects("CoverSaveIfEmpty")
	Set CB13 = SDB.Objects("SmallCoverSave")
	CB12.Common.Enabled = CB1.checked
	CB13.Common.Enabled = CB1.checked

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
	ini.StringValue("DiscogsAutoTagWeb", "CheckStyleField") = GetCustom(edt.ItemIndex - 1)

	If Sheet.Common.ChildControl("ControlSaveImage1").Checked = False Then
		ini.StringValue("DiscogsAutoTagWeb", "CheckSaveImage") = 2
	Else
		If Sheet.Common.ChildControl("ControlSaveImage12").Checked = False Then
			ini.StringValue("DiscogsAutoTagWeb", "CheckSaveImage") = 0
		Else
			ini.StringValue("DiscogsAutoTagWeb", "CheckSaveImage") = 1
		End If
	End If

	If Sheet.Common.ChildControl("ControlSaveImage13").Checked = True Then
		ini.BoolValue("DiscogsAutoTagWeb", "CheckSmallCover") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "CheckSmallCover") = false
	End If

	Set checkbox = Sheet.Common.ChildControl("CheckOriginalDiscogsTrack")
	If checkbox.checked Then
		ini.BoolValue("DiscogsAutoTagWeb", "CheckOriginalDiscogsTrack") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "CheckOriginalDiscogsTrack") = false
	End If

	Set edt = Sheet.Common.ChildControl("ArtistSeparator")
	ini.StringValue("DiscogsAutoTagWeb", "ArtistSeparator") = edt.Text

	Set checkbox = Sheet.Common.ChildControl("EditArtistLastSep")
	If checkbox.checked Then
		ini.BoolValue("DiscogsAutoTagWeb", "ArtistLastSeparator") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "ArtistLastSeparator") = false
	End If

	Set checkbox = Sheet.Common.ChildControl("UseMetalArchives")
	If checkbox.checked Then
		ini.BoolValue("DiscogsAutoTagWeb", "UseMetalArchives") = true
	Else
		ini.BoolValue("DiscogsAutoTagWeb", "UseMetalArchives") = false
	End If
	

	Script.UnregisterAllEvents

End Sub

Sub InitSheet2(Sheet)

	Dim UI : Set UI = SDB.UI
	Set ini = SDB.IniFile
	If Not (ini Is Nothing) Then
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
		If ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords") = "featuring,feat.,ft.,ft ,feat ,Rap,Rap [Featuring],Vocals [Featuring]"
		End If

		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords"), "|", ",")
		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords"), "|", ",")
		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords"), "|", ",")
		If Not InStr(ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords"), "|") = 0 Then ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords") = Replace(ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords"), "|", ",")
	End If

	LyricistKeywords = ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords")
	ConductorKeywords = ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords")
	ProducerKeywords = ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords")
	ComposerKeywords = ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords")
	FeaturingKeywords = ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords")

	Dim GroupBox0
	Set GroupBox0 = UI.NewGroupBox(Sheet)
	GroupBox0.Caption = "Enter the keywords for linking with discogs"
	GroupBox0.Common.Hint = "If you don't know what to enter here, let the keywords as is !!"
	GroupBox0.Common.SetRect 10, 10, 500, 235

	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 20, 50, 25
	Label2.Caption = SDB.Localize("Lyricist")
	Set EditLyricist = UI.NewEdit(GroupBox0)
	EditLyricist.Common.SetRect 20, 35, 450, 35
	EditLyricist.Common.ControlName = "LyricistKeywords"
	EditLyricist.Text = LyricistKeywords


	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 60, 50, 25
	Label2.Caption = SDB.Localize("Conductor")
	Set EditConductor = UI.NewEdit(GroupBox0)
	EditConductor.Common.SetRect 20, 75, 450, 35
	EditConductor.Common.ControlName = "ConductorKeywords"
	EditConductor.Text = ConductorKeywords


	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 100, 50, 25
	Label2.Caption = SDB.Localize("Producer")
	Set EditProducer = UI.NewEdit(GroupBox0)
	EditProducer.Common.SetRect 20, 115, 450, 35
	EditProducer.Common.ControlName = "ProducerKeywords"
	EditProducer.Text = ProducerKeywords


	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 140, 50, 25
	Label2.Caption = SDB.Localize("Composer")
	Set EditComposer = UI.NewEdit(GroupBox0)
	EditComposer.Common.SetRect 20, 155, 450, 35
	EditComposer.Common.ControlName = "ComposerKeywords"
	EditComposer.Text = ComposerKeywords

	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 180, 50, 25
	Label2.Caption = SDB.Localize("Featuring")
	Set EditFeaturing = UI.NewEdit(GroupBox0)
	EditFeaturing.Common.SetRect 20, 195, 450, 35
	EditFeaturing.Common.ControlName = "FeaturingKeywords"
	EditFeaturing.Text = FeaturingKeywords

End Sub

Sub SaveSheet2(Sheet)

	Dim ini
	Set ini = SDB.IniFile
	Dim edt
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

	Set edt = Sheet.Common.ChildControl("FeaturingKeywords")
	tmp = Split(edt.Text, ",")
	editText = ""
	For each x in tmp
		editText = editText & Trim(x) & ","
	Next
	ini.StringValue("DiscogsAutoTagWeb", "FeaturingKeywords") = Left(editText, Len(editText)-1)

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
