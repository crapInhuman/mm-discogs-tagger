'
' MediaMonkey Script
'
' NAME: Discogs Tagger Options 1.0
'
' AUTHOR: crap_inhuman
' DATE : 14/07/2013
'
'
' INSTALL: Automatic installation during Discogs Tagger install
'
'
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
	
	
	Dim GroupBox0
	Set GroupBox0 = UI.NewGroupBox(Sheet)
	GroupBox0.Caption = "Please choose the custom tags, where Discogs Tagger save the information"
	GroupBox0.Common.SetRect 10, 10, 500, 160

	Dim DD1
	Set DD1 = UI.NewDropDown(GroupBox0)
	DD1.Common.SetRect 240, 30, 200, 25
	DD1.Style = 2
	DD1.AddItem (CustomField1)
	DD1.AddItem (CustomField2)
	DD1.AddItem (CustomField3)
	DD1.AddItem (CustomField4)
	DD1.AddItem (CustomField5)

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
	Set Label1 = UI.NewLabel(GroupBox0)
	Label1.Common.SetRect 20, 35, 150, 25
	Label1.Caption = "Choose Tag for saving release-number"

	Dim DD2
	Set DD2 = UI.NewDropDown(GroupBox0)
	DD2.Common.SetRect 240, 60, 200, 25
	DD2.Style = 2
	DD2.AddItem (CustomField1)
	DD2.AddItem (CustomField2)
	DD2.AddItem (CustomField3)
	DD2.AddItem (CustomField4)
	DD2.AddItem (CustomField5)

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
	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 65, 150, 25
	Label2.Caption = "Choose Tag for saving catalog number"

	Set DD2 = UI.NewDropDown(GroupBox0)
	DD2.Common.SetRect 240, 90, 200, 25
	DD2.Style = 2
	DD2.AddItem (CustomField1)
	DD2.AddItem (CustomField2)
	DD2.AddItem (CustomField3)
	DD2.AddItem (CustomField4)
	DD2.AddItem (CustomField5)

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
	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 95, 150, 25
	Label2.Caption = "Choose Tag for saving release country"

	Set DD2 = UI.NewDropDown(GroupBox0)
	DD2.Common.SetRect 240, 120, 200, 25
	DD2.Style = 2
	DD2.AddItem (CustomField1)
	DD2.AddItem (CustomField2)
	DD2.AddItem (CustomField3)
	DD2.AddItem (CustomField4)
	DD2.AddItem (CustomField5)

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
	Set Label2 = UI.NewLabel(GroupBox0)
	Label2.Common.SetRect 20, 125, 150, 25
	Label2.Caption = "Choose Tag for saving media format"

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

End Sub

Function GetCustom(Index)

	Dim ini : Set ini = SDB.IniFile
	If Index = 0 Then GetCustom = "Custom1"
	If Index = 1 Then GetCustom = "Custom2"
	If Index = 2 Then GetCustom = "Custom3"
	If Index = 3 Then GetCustom = "Custom4"
	If Index = 4 Then GetCustom = "Custom5"

End Function
