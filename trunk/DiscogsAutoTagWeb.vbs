

Option Explicit
'
' Discogs Tagger Script for MediaMonkey ( Let & eepman & crap_inhuman )
'
Const VersionStr = "v4.50"

'Changes from 4.48 to 4.50 by crap_inhuman in 07.2014
'	Bug removed with utf-8 characters in searchstring (with big help from tillmanj !!)


'Changes from 4.47 to 4.48 by crap_inhuman in 07.2014
'	In the options menu you can now enter the access token manually
'	Bug removed in Keywords routine


'Changes from 4.46 to 4.47 by crap_inhuman in 07.2014
'	Changed the Delay - Function, WScript.sleep didn't work on all windows plattforms.


'Changes from 4.45 to 4.46 by crap_inhuman in 07.2014
'	The default settings for saving the Cover Images can now be changed in the options menu
'	Bug removed: Empty format-tag produced an error
'	Bug removed: Parsing wrong Artist Roles if a comma is between box brackets
'	Added OAuth authentication
'	Added option: Using Metal-Archives for release search (BETA)
'	Now it's possible to use * as wildcard in the Keywords
'	Added option: Print every involved people in a single line


'Changes from 4.44 to 4.45 by crap_inhuman in 05.2014
'	Bug removed: Didn't display the additional Image
'	Adjust the script for fetching the small album art
'	Adjust the script for removing leading and trailing spaces in Extra Artists
'	Add option to turn off subtrack detection


'Changes from 4.43 to 4.44 by crap_inhuman in 04.2014
'	Added simple routine to check and remove point in track positions (1. , 2. , 3. )
'	Bug removed: track position part
'	Max count for releases is set to 250


'Changes from 4.42 to 4.43 by crap_inhuman in 04.2014
'	Bug removed: Filter now work correctly
'	There's no max count for release results
'	Bug removed: Artist releases and Label releases work again


'Changes from 4.41 to 4.42 by crap_inhuman in 03.2014
'	Bug removed: Sub-Track do not select(set) the song
'	Added the option for switching the last artist separator ("&" or "chosen separator")


'Changes from 4.40 to 4.41 by crap_inhuman in 03.2014
'	Removed bug with more than one artist for a title
'	Added Artist separator to options menu
'	Added simple routine to check for false position separators


'Changes from 4.39 to 4.40 by crap_inhuman in 03.2014
'	featuring Keywords are now not case sensitive


'Changes from 4.38 to 4.39 by crap_inhuman in 03.2014
'	Keywords are now not case sensitive
'	Added Set Locale for supporting more countries


'Changes from 4.37 to 4.38 by crap_inhuman in 02.2014
'	Added the Featuring Keywords
'	Fixed a bug with the new submission form of discogs


'Changes from 4.36 to 4.37 by crap_inhuman in 02.2014
'	Changed the image access method


'Changes from 4.35 to 4.36 by crap_inhuman in 11.2013
'	The script now shows the filtered total and the matched total


'Changes from 4.34 to 4.35 by crap_inhuman in 11.2013
'	Raise the max count of release results to 100
'	Display the number of matched releases and which one you are viewing in the search bar


'Changes from 4.33 to 4.34 by crap_inhuman in 11.2013
'	Now it's possible to change the search string in the top bar


'Changes from 4.32 to 4.33 by crap_inhuman in 10.2013
'	Fixed a bug with the Separator


'Changes from 4.31 to 4.32 by crap_inhuman in 10.2013
'	Removed bug in extra artist assignment
'	Added 'Don't save' and 4 more fields for saving release-number


'Changes from 4.30 to 4.31 by crap_inhuman in 09.2013
'	Removed bug: Sub track name will not recognized if it is the last track
'	Removed bug: Script-Error occurred after closing the script-window, when no release found
'	Background of filter dropdown menu change to red if filter is selected (For better recognition)


'Changes from 4.00 to 4.30 by crap_inhuman in 07-09.2013
'	Added Sub tracks option.
'	Added option 'Unselect tracks without track-number'
'		Some albums at discogs have 'Index-Tracks'.
'		These tracks aren't song-tracks (e.g. Track-Name: 'Bonus track' or 'Live side')
'		This option unselect these tracks automatically
'		-------------------------------------------------------
'	Show a warning if the number of songs are different
'	For the catalog-number, release-country and media-format you can choose "Don't save" in the option menu, if you don't need it.
'	You can edit the keywords for linking the composer, producer, conductor,... tags with discogs
'	included DiscogsImages: you can choose more than one image for an album
'	New Option: Check 'Save Image' Checkbox only if release have no image
'	New Option: Choose another field for saving Style


'Changes from 3.65 to 4.00 by crap_inhuman in 07.2013
'	Bug removed with releases having leading zero in track-position
'	Added option for "Force NO Disc Usage". Helpful if a release have tracks with varying track-numbers (e.g. http://www.discogs.com/release/2942314 )
'		Without the option the script translate the varying track position to disc sides
'	Added option to show the original discogs track position
'	Moved the options to the left side for more place for the tracklisting
'	Moving the mouse-pointer over a checkbox now show more information about the usage
'	Now the chosen filters will be saved with the options
'		Choose one MediaType, MediaFormat, Country or Year from the drop-down list and save the options
'		or press one of the "Set ... Filter" button to select more than one Mediatype, MediaFormat, Country or Year
'		Choosing "Use ... Filter" in the drop-down list uses the custom filter-settings
'		Choosing "No ... Filter" from the drop-down list stop filtering the result
'		The Filter settings will only be saved if you press the "Save Options" button
'	The Custom Tags for saving the release, catalog, country and format will now be chosen in the options -> Discogs Tagger or during script installation
'	Showing the Data Quality of the Discogs release

'Changes from 3.64 to 3.65 by crap_inhuman in 07.2013
'	Bug removed: bug with additional artists removed, which only occur in rare cases

'Changes from 3.63 to 3.64 by crap_inhuman in 06.2013
'	Bug removed: selecting "Sides To Disc" and "Add Leading Zero", zero is dropped from track number and is displayed as a single digit

'Changes from 3.62 to 3.63 by crap_inhuman in 02.2013 (not released)
'	Bug removed: no search result -> no output

'Changes from 3.61 to 3.62 by crap_inhuman in 02.2013
'	Insert code for supporting french language machines
'	Comments will now be saved
'	Delete some unused but declared variables
'	Name for "feat." can be edit
'	Some small bugfixes

'Changes from 3.6 to 3.61 by crap_inhuman in 02.2013
'	Removed a bug in the option 'Featuring Artist behind title'
'	Better implementation of the option 'Featuring Artist behind title'
'	Inserting Master and Release URLs now work in the Search-Panel

'Changes from 3.5 to 3.6 by crap_inhuman in 02.2013
'	Implementation of eepman's JSON-Parser
'	Now read the user-specific Separator characters and use it for separating
'	Label / Artist / Master Search now using the JSON Parser too
'	Some bugfixes

'Changes from 3.3 to 3.5 by crap_inhuman in 01.2013
'	Now you can choose which Custom Tag will be used for the Tags: ReleaseID, Catalog, Country and Format
'	The "Credits for ExtraArtists in tracks" will now saved in MediaMonkey !
'	Added option for "Add Leading zero to Tracknumbers"
'	Added option for "Include Producer"
'	Added JSON Parser for the new Discogs-API

'	Some bugfixes (Filter now working correct)
'	Added the option to choose the place for Featuring Artist (Artist or Title)
'	e.g. Aaliyah - We Need a Resolution (ft. Timbaland) -or- Aaliyah (ft. Timbaland) - We Need a Resolution
'	Changeable Name for "Various" Artists (Various Artists)
'	Added option for "Adding comment"
'	Get OriginalDate from Master-Release if available
'	The Script now reads the saved Discogs Release-ID from the chosen Release-Tag


' ToDo: Add more tooltips to the html
'		Erster und letzter Buchstabe in SearchArtist fehlt (wenn nächster Buchstabe blank ist e.g. "3 doors down", "Miss may i")
'		Add option to limit release result for faster results
'		Wrong Publisher, Producer, etc. in Subtracks. The script only take the info from the first subtrack..

' WebBrowser is visible browser object with display of discogs album info
Dim WebBrowser

' decoded json object representing currently selected release
Dim CurrentRelease

Dim UI

Dim Results, ResultsReleaseID ' result list
Dim CurrentResultID
Dim ini

Dim CheckAlbum, CheckArtist, CheckAlbumArtist, CheckAlbumArtistFirst, CheckLabel, CheckDate, CheckOrigDate, CheckGenre
Dim CheckCountry, CheckCover, CheckSmallCover, SmallCover, CheckStyle, CheckCatalog, CheckRelease, CheckInvolved, CheckLyricist
Dim CheckComposer, CheckConductor, CheckProducer, CheckDiscNum, CheckTrackNum, CheckFormat, CheckUseAnv, CheckYearOnlyDate
Dim CheckForceNumeric, CheckSidesToDisc, CheckForceDisc, CheckNoDisc, CheckLeadingZero, CheckVarious, TxtVarious
Dim CheckTitleFeaturing, CheckComment, CheckFeaturingName, TxtFeaturingName, CheckOriginalDiscogsTrack, CheckSaveImage
Dim CheckStyleField, CheckTurnOffSubTrack, CheckInvolvedPeopleSingleLine
Dim SubTrackNameSelection
Dim CountryFilterList, MediaTypeFilterList, MediaFormatFilterList, YearFilterList
Dim LyricistKeywords, ConductorKeywords, ProducerKeywords, ComposerKeywords, FeaturingKeywords

Dim SavedReleaseId
Dim SavedSearchTerm, SavedSearchArtist, SavedSearchAlbum
Dim SavedMasterId, SavedArtistId, SavedLabelId

Dim FilterMediaType, FilterCountry, FilterYear, FilterMediaFormat, CurrentLoadType
Dim MediaTypeList, MediaFormatList, CountryList, YearList, AlternativeList, LoadList
Dim ArtistSeparator, ArtistLastSeparator

Dim FirstTrack
Dim AlbumArtURL, AlbumArtThumbNail
Dim iMaxTracks
Dim iAutoTrackNumber, iAutoDiscNumber
Dim LastDisc
Dim SelectAll, UnselectedTracks(1000)

Dim ReleaseTag, CountryTag, CatalogTag, FormatTag
Dim ReleaseTagList, CountryTagList, CatalogTagList, FormatTagList
Dim OriginalDate, Separator
Dim OptionsChanged
Dim AccessToken, AccessTokenSecret

Dim fso, loc, logf

Dim UseMetalArchives
REM UseMetalArchives = False

'----------------------------------DiscogsImages----------------------------------------
Dim SaveImageType, SaveImage, CoverStorage, FileNameList
Dim ImageTypeList, ImageList
Dim list
Dim ImagesCount
Dim SaveMoreImages
Dim WebBrowser3
Dim SelectedSongsGlobal
'----------------------------------DiscogsImages----------------------------------------

' Easier access of SDB.UI
Set UI = SDB.UI

' MediaMonkey calls this method whenever a search is started using this script
Sub StartSearch(Panel, SearchTerm, SearchArtist, SearchAlbum)

	Dim tmpCountry, tmpCountry2, tmpMediaType, tmpMediaType2, tmpMediaFormat, tmpMediaFormat2, tmpYear, tmpYear2
	Dim i, a, tmp
	Set CountryFilterList = SDB.NewStringList
	Set MediaTypeFilterList = SDB.NewStringList
	Set MediaFormatFilterList = SDB.NewStringList
	Set YearFilterList = SDB.NewStringList
	SearchTerm = LTrim(SearchTerm)

	OptionsChanged = False

	'*FilterList.Item(0) = "0" -> No Filter
	'*FilterList.Item(0) = "1" -> Custom Filter
	'*FilterList.Item(0) = "2" -> Selected Country/MediaType/MediaFormat/Year

	Set ini = SDB.IniFile
	If Not (ini Is Nothing) Then
		'We init default settings only if they do not exist in ini file yet
		If ini.StringValue("DiscogsAutoTagWeb","CheckAlbum") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckAlbum") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckArtist") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckArtist") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckAlbumArtist") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtist") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckLabel") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckLabel") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckDate") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckDate") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckOrigDate") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckOrigDate") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckGenre") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckGenre") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckStyle") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckStyle") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckCountry") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckCountry") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage") = 1
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSmallCover") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckSmallCover") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckCatalog") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckCatalog") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckRelease") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckRelease") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckInvolved") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckInvolved") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckLyricist") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckLyricist") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckComposer") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckComposer") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckConductor") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckConductor") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckProducer") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckProducer") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckDiscNum") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckDiscNum") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckTrackNum") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckTrackNum") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckFormat") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckFormat") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckUseAnv") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckUseAnv") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckYearOnlyDate") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckYearOnlyDate") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckForceNumeric") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckForceNumeric") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSidesToDisc") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckSidesToDisc") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckForceDisc") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckForceDisc") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckNoDisc") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckNoDisc") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = False
		End If
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
		If ini.StringValue("DiscogsAutoTagWeb","CheckLeadingZero") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZero") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckVarious") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckVarious") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","TxtVarious") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","TxtVarious") = "Various Artists"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckTitleFeaturing") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckTitleFeaturing") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckFeaturingName") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckFeaturingName") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName") = "feat."
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckComment") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckComment") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","SubTrackNameSelection") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","SubTrackNameSelection") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter") = "" Then
			tmp = "0"
			For a = 1 to 282
				tmp = tmp & ",0"
			Next
			ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter") = tmp
		End If

		If ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter") = "" Then
			tmp = "0"
			For a = 1 to 38
				tmp = tmp & ",0"
			Next
			ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter") = tmp
		End If

		If ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter") = "" Then
			tmp = "0"
			For a = 1 to 48
				tmp = tmp & ",0"
			Next
			ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter") = tmp
		End If

		If ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = "" Then
			tmp = "0"
			For a = Year(Date) To 1900 Step -1
				tmp = tmp & ",0"
			Next
			ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = tmp
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
		If ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords") = "featuring,feat.,ft.,ft ,feat ,Rap,Rap [Featuring],Vocals [Featuring]"
		End If

		If ini.StringValue("DiscogsAutoTagWeb","CheckStyleField") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","CheckStyleField") = "Default (stored with Genre)"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = ", "
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ArtistLastSeparator") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckTurnOffSubTrack") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckTurnOffSubTrack") = False
		End If

		If ini.ValueExists("DiscogsAutoTagWeb","CheckNotAlwaysSaveimage") Then
			ini.DeleteKey "DiscogsAutoTagWeb","CheckNotAlwaysSaveimage"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","AccessToken") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","AccessToken") = ""
		End If
		If ini.StringValue("DiscogsAutoTagWeb","AccessTokenSecret") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","AccessTokenSecret") = ""
		End If
		If ini.StringValue("DiscogsAutoTagWeb","UseMetalArchives") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","UseMetalArchives") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckInvolvedPeopleSingleLine") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckInvolvedPeopleSingleLine") = False
		End If

		'----------------------------------DiscogsImages----------------------------------------
		CoverStorage = ini.StringValue("PreviewSettings","DefaultCoverStorage")
		'Coverstorage = 0 -> Save image to tag (if possible) otherwise save to file folder
		'Coverstorage = 1 -> Save image to file folder
		'Coverstorage = 2 -> Save image to cover folder (is deprecated and will not be supported !!)
		'Coverstorage = 3 -> Save image to tag (if possible) and to file folder
		If CoverStorage = 2 Then
			Call SDB.MessageBox("Discogs Images: Your Cover Storage is not supported by DiscogsImages !",mtError,Array(mbOk))
			Exit Sub
		End If
		'----------------------------------DiscogsImages----------------------------------------

	End If

	CheckAlbum = ini.BoolValue("DiscogsAutoTagWeb","CheckAlbum")
	CheckArtist = ini.BoolValue("DiscogsAutoTagWeb","CheckArtist")
	CheckAlbumArtist = ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtist")
	CheckAlbumArtistFirst = ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst")
	CheckLabel = ini.BoolValue("DiscogsAutoTagWeb","CheckLabel")
	CheckDate = ini.BoolValue("DiscogsAutoTagWeb","CheckDate")
	CheckOrigDate = ini.BoolValue("DiscogsAutoTagWeb","CheckOrigDate")
	CheckGenre = ini.BoolValue("DiscogsAutoTagWeb","CheckGenre")
	CheckStyle = ini.BoolValue("DiscogsAutoTagWeb","CheckStyle")
	CheckCountry = ini.BoolValue("DiscogsAutoTagWeb","CheckCountry")
	CheckSaveImage = ini.StringValue("DiscogsAutoTagWeb","CheckSaveImage")
	CheckSmallCover = ini.BoolValue("DiscogsAutoTagWeb","CheckSmallCover")
	CheckCatalog = ini.BoolValue("DiscogsAutoTagWeb","CheckCatalog")
	CheckRelease = ini.BoolValue("DiscogsAutoTagWeb","CheckRelease")
	CheckInvolved = ini.BoolValue("DiscogsAutoTagWeb","CheckInvolved")
	CheckLyricist = ini.BoolValue("DiscogsAutoTagWeb","CheckLyricist")
	CheckComposer = ini.BoolValue("DiscogsAutoTagWeb","CheckComposer")
	CheckConductor = ini.BoolValue("DiscogsAutoTagWeb","CheckConductor")
	CheckProducer = ini.BoolValue("DiscogsAutoTagWeb","CheckProducer")
	CheckDiscNum = ini.BoolValue("DiscogsAutoTagWeb","CheckDiscNum")
	CheckTrackNum = ini.BoolValue("DiscogsAutoTagWeb","CheckTrackNum")
	CheckFormat = ini.BoolValue("DiscogsAutoTagWeb","CheckFormat")
	CheckUseAnv = ini.BoolValue("DiscogsAutoTagWeb","CheckUseAnv")
	CheckYearOnlyDate = ini.BoolValue("DiscogsAutoTagWeb","CheckYearOnlyDate")
	CheckForceNumeric = ini.BoolValue("DiscogsAutoTagWeb","CheckForceNumeric")
	CheckSidesToDisc = ini.BoolValue("DiscogsAutoTagWeb","CheckSidesToDisc")
	CheckForceDisc = ini.BoolValue("DiscogsAutoTagWeb","CheckForceDisc")
	CheckOriginalDiscogsTrack = ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack")
	CheckNoDisc = ini.BoolValue("DiscogsAutoTagWeb","CheckNoDisc")
	ReleaseTag = ini.StringValue("DiscogsAutoTagWeb","ReleaseTag")
	CatalogTag = ini.StringValue("DiscogsAutoTagWeb","CatalogTag")
	CountryTag = ini.StringValue("DiscogsAutoTagWeb","CountryTag")
	FormatTag = ini.StringValue("DiscogsAutoTagWeb","FormatTag")
	CheckLeadingZero = ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZero")
	CheckVarious = ini.BoolValue("DiscogsAutoTagWeb","CheckVarious")
	TxtVarious = ini.StringValue("DiscogsAutoTagWeb","TxtVarious")
	CheckTitleFeaturing = ini.BoolValue("DiscogsAutoTagWeb","CheckTitleFeaturing")
	CheckFeaturingName = ini.boolValue("DiscogsAutoTagWeb","CheckFeaturingName")
	TxtFeaturingName = ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName")
	CheckComment = ini.BoolValue("DiscogsAutoTagWeb","CheckComment")
	SubTrackNameSelection = ini.BoolValue("DiscogsAutoTagWeb","SubTrackNameSelection")
	Separator = ini.StringValue("Appearance","MultiStringSeparator")
	tmpCountry = ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter")
	tmpCountry2 = Split(tmpCountry, ",")
	tmpMediaType = ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter")
	tmpMediaType2 = Split(tmpMediaType, ",")
	tmpMediaFormat = ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter")
	tmpMediaFormat2 = Split(tmpMediaFormat, ",")
	tmpYear = ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter")
	tmpYear2 = Split(tmpYear, ",")
	LyricistKeywords = ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords")
	ConductorKeywords = ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords")
	ProducerKeywords = ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords")
	ComposerKeywords = ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords")
	FeaturingKeywords = ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords")
	REM CheckNotAlwaysSaveImage = ini.BoolValue("DiscogsAutoTagWeb","CheckNotAlwaysSaveImage")
	CheckStyleField = ini.StringValue("DiscogsAutoTagWeb","CheckStyleField")
	ArtistSeparator = ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator")
	ArtistLastSeparator = ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator")
	CheckTurnOffSubTrack = ini.BoolValue("DiscogsAutoTagWeb","CheckTurnOffSubTrack")
	AccessToken = ini.StringValue("DiscogsAutoTagWeb","AccessToken")
	AccessTokenSecret = ini.StringValue("DiscogsAutoTagWeb","AccessTokenSecret")
	UseMetalArchives = ini.BoolValue("DiscogsAutoTagWeb","UseMetalArchives")
	CheckInvolvedPeopleSingleLine = ini.BoolValue("DiscogsAutoTagWeb","CheckInvolvedPeopleSingleLine")

	Separator = Left(Separator, Len(Separator)-1)
	Separator = Right(Separator, Len(Separator)-1)

	SelectAll = True


	WriteLogInit  'Only use for debugging

	WriteLog " "
	WriteLog("SearchTerm=" & SearchTerm)
	WriteLog("SearchArtist=" & SearchArtist)
	WriteLog("SearchAlbum=" & SearchAlbum)
	WriteLog " "


	Set MediaTypeList = SDB.NewStringList
	Set MediaFormatList = SDB.NewStringList
	Set CountryList = SDB.NewStringList
	Set YearList = SDB.NewStringList
	Set AlternativeList = SDB.NewStringList
	Set LoadList = SDB.NewStringList

	LoadList.Add "Search Results"
	LoadList.Add "Master Release"
	LoadList.Add "Releases of Artist"
	LoadList.Add "Releases of Label"

	MediaTypeList.Add "None"
	MediaTypeList.Add "Vinyl"
	MediaTypeList.Add "CD"
	MediaTypeList.Add "DVD"
	MediaTypeList.Add "Blu-Ray"
	MediaTypeList.Add "Cassette"
	MediaTypeList.Add "DAT"
	MediaTypeList.Add "Minidisc"
	MediaTypeList.Add "File"
	MediaTypeList.Add "Acetate"
	MediaTypeList.Add "Flexi-disc"
	MediaTypeList.Add "Lathe Cut"
	MediaTypeList.Add "Shellac"
	MediaTypeList.Add "Pathé Disc"
	MediaTypeList.Add "Edison Disc"
	MediaTypeList.Add "Cylinder"
	MediaTypeList.Add "CDr"
	MediaTypeList.Add "CDV"
	MediaTypeList.Add "DVDr"
	MediaTypeList.Add "HD DVD"
	MediaTypeList.Add "HD DVD-R"
	MediaTypeList.Add "Blue-ray-R"
	MediaTypeList.Add "4-Track Cartridge"
	MediaTypeList.Add "8-Track Cartridge"
	MediaTypeList.Add "DCC"
	MediaTypeList.Add "Microcassette"
	MediaTypeList.Add "Reel-To-Reel"
	MediaTypeList.Add "Betamax"
	MediaTypeList.Add "VHS"
	MediaTypeList.Add "Video 2000"
	MediaTypeList.Add "Laserdisc"
	MediaTypeList.Add "SelectaVision"
	MediaTypeList.Add "VHD"
	MediaTypeList.Add "MVD"
	MediaTypeList.Add "UMD"
	MediaTypeList.Add "Floppy Disk"
	MediaTypeList.Add "Memory Stick"
	MediaTypeList.Add "Hybrid"
	MediaTypeList.Add "Box Set"

	MediaFormatList.Add "None"
	MediaFormatList.Add "Album"
	MediaFormatList.Add "Mini-Album"
	MediaFormatList.Add "Compilation"
	MediaFormatList.Add "Single"
	MediaFormatList.Add "Maxi-Single"
	MediaFormatList.Add "7"""
	MediaFormatList.Add "12"""
	MediaFormatList.Add "LP"
	MediaFormatList.Add "EP"
	MediaFormatList.Add "Single Sided"
	MediaFormatList.Add "Enhanced"
	MediaFormatList.Add "Limited Edition"
	MediaFormatList.Add "Reissue"
	MediaFormatList.Add "Remastered"
	MediaFormatList.Add "Repress"
	MediaFormatList.Add "Test Pressing"
	MediaFormatList.Add "Unofficial"
	MediaFormatList.Add "Promo"
	MediaFormatList.Add "White Label"
	MediaFormatList.Add "Mixed"
	MediaFormatList.Add "Sampler"
	MediaFormatList.Add "MP3"
	MediaFormatList.Add "FLAC"
	MediaFormatList.Add "16"""
	MediaFormatList.Add "11"""
	MediaFormatList.Add "10"""
	MediaFormatList.Add "9"""
	MediaFormatList.Add "8"""
	MediaFormatList.Add "6"""
	MediaFormatList.Add "5"""
	MediaFormatList.Add "4"""
	MediaFormatList.Add "3"""
	MediaFormatList.Add "45 RPM"
	MediaFormatList.Add "78 RPM"
	MediaFormatList.Add "Shape"
	MediaFormatList.Add "Card Backed"
	MediaFormatList.Add "Etched"
	MediaFormatList.Add "Picture Disc"
	MediaFormatList.Add "Stereo"
	MediaFormatList.Add "Mono"
	MediaFormatList.Add "Quadraphonic"
	MediaFormatList.Add "Ambisonic"
	MediaFormatList.Add "Mispress"
	MediaFormatList.Add "Misprint"
	MediaFormatList.Add "Partially Mixed"
	MediaFormatList.Add "Unofficial Release"
	MediaFormatList.Add "Partially Unofficial"
	MediaFormatList.Add "Copy Protected"

	CountryList.Add "None"
	CountryList.Add "Australia"
	CountryList.Add "Belgium"
	CountryList.Add "Brazil"
	CountryList.Add "Canada"
	CountryList.Add "China"
	CountryList.Add "Cuba"
	CountryList.Add "France"
	CountryList.Add "Germany"
	CountryList.Add "Italy"
	CountryList.Add "Jamaica"
	CountryList.Add "Japan"
	CountryList.Add "Ireland"
	CountryList.Add "India"
	CountryList.Add "Mexico"
	CountryList.Add "Netherlands"
	CountryList.Add "New Zealand"
	CountryList.Add "Spain"
	CountryList.Add "Sweden"
	CountryList.Add "Switzerland"
	CountryList.Add "UK"
	CountryList.Add "US"
	CountryList.Add "=========="
	CountryList.Add "Africa"
	CountryList.Add "Asia"
	CountryList.Add "Australasia"
	CountryList.Add "Benelux"
	CountryList.Add "Central America"
	CountryList.Add "Europe"
	CountryList.Add "Gulf Cooperation Council"
	CountryList.Add "North America"
	CountryList.Add "Scandinavia"
	CountryList.Add "South America"
	CountryList.Add "==========="
	CountryList.Add "Afghanistan"
	CountryList.Add "Akrotiri"
	CountryList.Add "Albania"
	CountryList.Add "Algeria"
	CountryList.Add "American Samoa"
	CountryList.Add "Andorra"
	CountryList.Add "Angola"
	CountryList.Add "Anguilla"
	CountryList.Add "Antarctica"
	CountryList.Add "Antigua & Barbuda"
	CountryList.Add "Argentina"
	CountryList.Add "Armenia"
	CountryList.Add "Aruba"
	CountryList.Add "Ashmore & Cartier Islands"
	CountryList.Add "Austria"
	CountryList.Add "Azerbaijan"
	CountryList.Add "Bahamas"
	CountryList.Add "Bahrain"
	CountryList.Add "Baker Island"
	CountryList.Add "Bangladesh"
	CountryList.Add "Barbados"
	CountryList.Add "Bassas da India"
	CountryList.Add "Belarus"
	CountryList.Add "Belize"
	CountryList.Add "Benin"
	CountryList.Add "Bermuda"
	CountryList.Add "Bhutan"
	CountryList.Add "Bolivia"
	CountryList.Add "Bosnia & Herzegovina"
	CountryList.Add "Botswana"
	CountryList.Add "Bouvet Island"
	CountryList.Add "British Indian Ocean"
	CountryList.Add "British Virgin Islands"
	CountryList.Add "Brunei"
	CountryList.Add "Bulgaria"
	CountryList.Add "Burkina Faso"
	CountryList.Add "Burma"
	CountryList.Add "Burundi"
	CountryList.Add "Cambodia"
	CountryList.Add "Cameroon"
	CountryList.Add "Cape Verde"
	CountryList.Add "Cayman Islands"
	CountryList.Add "Central African Republic"
	CountryList.Add "Chad"
	CountryList.Add "Chile"
	CountryList.Add "Christmas Island"
	CountryList.Add "Clipperton Island"
	CountryList.Add "Cocos Islands"
	CountryList.Add "Colombia"
	CountryList.Add "Comoros"
	CountryList.Add "Congo"
	CountryList.Add "Cook Islands"
	CountryList.Add "Coral Sea Islands"
	CountryList.Add "Costa Rica"
	CountryList.Add "Croatia"
	CountryList.Add "Cyprus"
	CountryList.Add "Czech Republic"
	CountryList.Add "Czechoslovakia"
	CountryList.Add "Denmark"
	CountryList.Add "Dhekelia"
	CountryList.Add "Djibouti"
	CountryList.Add "Dominica"
	CountryList.Add "Dominican Republic"
	CountryList.Add "East Timor"
	CountryList.Add "Ecuador"
	CountryList.Add "Egypt"
	CountryList.Add "El Salvador"
	CountryList.Add "Equatorial Guinea"
	CountryList.Add "Eritrea"
	CountryList.Add "Estonia"
	CountryList.Add "Ethiopia"
	CountryList.Add "Europa Island"
	CountryList.Add "Falkland Islands"
	CountryList.Add "Faroe Islands"
	CountryList.Add "Fiji"
	CountryList.Add "Finland"
	CountryList.Add "French Guiana"
	CountryList.Add "French Polynesia"
	CountryList.Add "French Southern"
	CountryList.Add "Gabon"
	CountryList.Add "Gambia"
	CountryList.Add "Gaza Strip"
	CountryList.Add "Georgia"
	CountryList.Add "German Democratic Republic"
	CountryList.Add "Ghana"
	CountryList.Add "Gibraltar"
	CountryList.Add "Glorioso Islands"
	CountryList.Add "Greece"
	CountryList.Add "Greenland"
	CountryList.Add "Grenada"
	CountryList.Add "Guadeloupe"
	CountryList.Add "Guam"
	CountryList.Add "Guatemala"
	CountryList.Add "Guernsey"
	CountryList.Add "Guinea"
	CountryList.Add "Guinea-Bissau"
	CountryList.Add "Guyana"
	CountryList.Add "Haiti"
	CountryList.Add "Heard Island"
	CountryList.Add "McDonald Islands"
	CountryList.Add "Holy See"
	CountryList.Add "Honduras"
	CountryList.Add "Hong Kong"
	CountryList.Add "Howland Island"
	CountryList.Add "Hungary"
	CountryList.Add "Iceland"
	CountryList.Add "Indonesia"
	CountryList.Add "Iran"
	CountryList.Add "Iraq"
	CountryList.Add "Israel"
	CountryList.Add "Ivory Coast"
	CountryList.Add "Jan Mayen"
	CountryList.Add "Jarvis Island"
	CountryList.Add "Jersey"
	CountryList.Add "Johnston Atoll"
	CountryList.Add "Jordan"
	CountryList.Add "Juan de Nova Island"
	CountryList.Add "Kazakhstan"
	CountryList.Add "Kenya"
	CountryList.Add "Kingman Reef"
	CountryList.Add "Kiribati"
	CountryList.Add "Kuwait"
	CountryList.Add "Kyrgyzstan"
	CountryList.Add "Laos"
	CountryList.Add "Latvia"
	CountryList.Add "Lebanon"
	CountryList.Add "Lesotho"
	CountryList.Add "Liberia"
	CountryList.Add "Libya"
	CountryList.Add "Liechtenstein"
	CountryList.Add "Lithuania"
	CountryList.Add "Luxembourg"
	CountryList.Add "Macau"
	CountryList.Add "Macedonia"
	CountryList.Add "Madagascar"
	CountryList.Add "Malawi"
	CountryList.Add "Malaysia"
	CountryList.Add "Maldives"
	CountryList.Add "Mali"
	CountryList.Add "Malta"
	CountryList.Add "Man, Isle of"
	CountryList.Add "Marshall Islands"
	CountryList.Add "Martinique"
	CountryList.Add "Mauritania"
	CountryList.Add "Mauritius"
	CountryList.Add "Mayotte"
	CountryList.Add "Micronesia"
	CountryList.Add "Midway Islands"
	CountryList.Add "Moldova"
	CountryList.Add "Monaco"
	CountryList.Add "Mongolia"
	CountryList.Add "Montenegro"
	CountryList.Add "Montserrat"
	CountryList.Add "Morocco"
	CountryList.Add "Mozambique"
	CountryList.Add "Namibia"
	CountryList.Add "Nauru"
	CountryList.Add "Navassa Island"
	CountryList.Add "Nepal"
	CountryList.Add "Netherlands Antilles"
	CountryList.Add "New Caledonia"
	CountryList.Add "Nicaragua"
	CountryList.Add "Niger"
	CountryList.Add "Nigeria"
	CountryList.Add "Niue"
	CountryList.Add "Norfolk Island"
	CountryList.Add "Northern Mariana Islands"
	CountryList.Add "North Korea"
	CountryList.Add "Norway"
	CountryList.Add "Oman"
	CountryList.Add "Pakistan"
	CountryList.Add "Palau"
	CountryList.Add "Palmyra Atoll"
	CountryList.Add "Panama"
	CountryList.Add "Papua New Guinea"
	CountryList.Add "Paracel Islands"
	CountryList.Add "Paraguay"
	CountryList.Add "Peru"
	CountryList.Add "Philippines"
	CountryList.Add "Pitcairn Islands"
	CountryList.Add "Poland"
	CountryList.Add "Portugal"
	CountryList.Add "Puerto Rico"
	CountryList.Add "Qatar"
	CountryList.Add "Reunion"
	CountryList.Add "Romania"
	CountryList.Add "Russia"
	CountryList.Add "Rwanda"
	CountryList.Add "Saint Helena"
	CountryList.Add "Saint Kitts and Nevis"
	CountryList.Add "Saint Lucia"
	CountryList.Add "Saint Pierre"
	CountryList.Add "Saint Vincent"
	CountryList.Add "Samoa"
	CountryList.Add "San Marino"
	CountryList.Add "Sao Tome and Principe"
	CountryList.Add "Saudi Arabia"
	CountryList.Add "Senegal"
	CountryList.Add "Serbia"
	CountryList.Add "Serbia and Montenegro"
	CountryList.Add "Seychelles"
	CountryList.Add "Sierra Leone"
	CountryList.Add "Singapore"
	CountryList.Add "Slovakia"
	CountryList.Add "Slovenia"
	CountryList.Add "Solomon Islands"
	CountryList.Add "Somalia"
	CountryList.Add "South Africa"
	CountryList.Add "South Korea"
	CountryList.Add "Spratly Islands"
	CountryList.Add "Sri Lanka"
	CountryList.Add "Sudan"
	CountryList.Add "Suriname"
	CountryList.Add "Svalbard"
	CountryList.Add "Swaziland"
	CountryList.Add "Syria"
	CountryList.Add "Tajikistan"
	CountryList.Add "Tanzania"
	CountryList.Add "Thailand"
	CountryList.Add "Taiwan"
	CountryList.Add "Togo"
	CountryList.Add "Tokelau"
	CountryList.Add "Tonga"
	CountryList.Add "Trinidad & Tobago"
	CountryList.Add "Tromelin Island"
	CountryList.Add "Tunisia"
	CountryList.Add "Turkey"
	CountryList.Add "Turkmenistan"
	CountryList.Add "Turks and Caicos Islands"
	CountryList.Add "Tuvalu"
	CountryList.Add "Uganda"
	CountryList.Add "Ukraine"
	CountryList.Add "United Arab Emirates"
	CountryList.Add "Uruguay"
	CountryList.Add "USSR"
	CountryList.Add "Uzbekistan"
	CountryList.Add "Vatican City"
	CountryList.Add "Vanuatu"
	CountryList.Add "Venezuela"
	CountryList.Add "Vietnam"
	CountryList.Add "Virgin Islands"
	CountryList.Add "Wake Island"
	CountryList.Add "Wallis and Futuna"
	CountryList.Add "West Bank"
	CountryList.Add "Western Sahara"
	CountryList.Add "Yemen"
	CountryList.Add "Yugoslavia"
	CountryList.Add "Zambia"
	CountryList.Add "Zimbabwe"

	YearList.Add "None"
	For i=Year(Date) To 1900 Step -1
		YearList.Add i
	Next

	If UBound(tmpYear2) <> YearList.Count -1 Then
		'MsgBox UBound(tmpYear2) & " -- " & YearList.Count -1
		tmpYear = tmpYear & ",1"
		ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = tmpYear
		tmpYear2 = Split(tmpYear, ",")
	End If

	For a = 0 To CountryList.Count - 1
		CountryFilterList.Add tmpCountry2(a)
	Next

	For a = 0 To MediaTypeList.Count - 1
		MediaTypeFilterList.Add tmpMediaType2(a)
	Next

	For a = 0 To MediaFormatList.Count - 1
		MediaFormatFilterList.Add tmpMediaFormat2(a)
	Next

	For a = 0 To YearList.Count - 1
		YearFilterList.Add tmpYear2(a)
	Next

	AddAlternative SearchTerm
	AddAlternative SearchArtist
	AddAlternative SearchAlbum

	For i = 0 To SDB.Tools.WebSearch.NewTracks.Count - 1
		AddAlternatives SDB.Tools.WebSearch.NewTracks.item(i)
	Next

	If MediaTypeFilterList.Item(0) = "0" Then
		FilterMediaType = "None"
	ElseIf MediaTypeFilterList.Item(0) = "1" Then
		FilterMediaType = "Use MediaType Filter"
	Else
		FilterMediaType = MediaTypeFilterList.Item(0)
	End If

	If MediaFormatFilterList.Item(0) = "0" Then
		FilterMediaFormat = "None"
	ElseIf MediaFormatFilterList.Item(0) = "1" Then
		FilterMediaFormat = "Use MediaFormat Filter"
	Else
		FilterMediaFormat = MediaFormatFilterList.Item(0)
	End If

	If CountryFilterList.Item(0) = "0" Then
		FilterCountry = "None"
	ElseIf CountryFilterList.Item(0) = "1" Then
		FilterCountry = "Use Country Filter"
	Else
		FilterCountry = CountryFilterList.Item(0)
	End If

	If YearFilterList.Item(0) = "0" Then
		FilterYear = "None"
	ElseIf YearFilterList.Item(0) = "1" Then
		FilterYear = "Use Year Filter"
	Else
		FilterYear = YearFilterList.Item(0)
	End If

	CurrentLoadType = "Search Results"

	' This is a web browser that we use to present results to the user
	Set WebBrowser = UI.NewActiveX(Panel, "Shell.Explorer")
	WebBrowser.Common.Align = 5
	WebBrowser.Common.ControlName = "WebBrowser"
	WebBrowser.Common.Top = 0
	WebBrowser.Common.Left = 0
	SDB.Objects("WebBrowser") = WebBrowser
	WebBrowser.Interf.Visible = true
	WebBrowser.Common.BringToFront

	If SDB.Tools.WebSearch.NewTracks.Count > 0 Then
		Set FirstTrack = SDB.Tools.WebSearch.NewTracks.item(0)
		SavedReleaseId = get_release_ID(FirstTrack) 'get saved Release_ID from User-Defined Custom-Tag
		SavedSearchTerm = SearchTerm
		SavedSearchArtist = SearchArtist
		SavedSearchAlbum = SearchAlbum
	End If





	If UseMetalArchives = False Then
		WriteLog "Start Discogs Request"

		Dim IEobj, objShell, objShellWindows
		Dim dteWait, objIE, strURL, retIE

		If AccessToken = "" Or AccessTokenSecret = "" Then
			MsgBox("Starting August 15th, access to discogs database will require authentication." & vbNewLine & "This is part of an ongoing effort to improve API uptime and response times" & vbNewLine & vbNewLine & "You need an account at discogs in order to use Discogs Tagger.")
			set IEobj = CreateObject("InternetExplorer.Application")
			Set objShell = CreateObject("Shell.Application")
			Set objShellWindows = objShell.Windows
			IEobj.visible = true

			IEobj.navigate ("http://www.germanc64.de/mm/oauth/login_with_discogs.php")

			WriteLog "IE started"

			dteWait = DateAdd("s", 20, Now())
	
			Do Until (Now() > dteWait)
				SDB.ProcessMessages
			Loop
			
			Do
				If objShellWindows.Count = 0 Then
					Exit Do
				End If
				For i = 0 to objShellWindows.Count - 1
					Set objIE = objShellWindows.Item(i)
					strURL = objIE.LocationURL
					If InStr(strURL, "oauth_verifier") <> 0 Then
						writeLog strURL
						retIE = objIE.document.body.innerText
						WriteLog retIE
						Exit Do
					End If
				Next
			Loop While 1 = 1

			WriteLog "IE finished"
			
			Dim start
			If InStr(retIE, "AccessToken=") <> 0 Then
				start = InStr(retIE, "AccessToken=")
				retIE = Mid(retIE, start + 12)
				start = InStr(retIE, " ")
				AccessToken = Left(retIE, start -1)
				ini.StringValue("DiscogsAutoTagWeb","AccessToken") = AccessToken
				WriteLog "AccessToken=" & AccessToken
				start = InStr(retIE, "AccessTokenSecret=")
				retIE = Mid(retIE, start + 18)
				start = InStr(retIE, " ")
				AccessTokenSecret = Left(retIE, start -1)
				ini.StringValue("DiscogsAutoTagWeb","AccessTokenSecret") = AccessTokenSecret
				WriteLog "AccessTokenSecret=" & AccessTokenSecret
			End If
				
			objIE.visible = False 
			'IEobj = Nothing
			objIE.Quit
		Else
			WriteLog "AccessToken found in ini = " & AccessToken
			WriteLog "AccessTokenSecret found in ini = " & AccessTokenSecret
		End If
		WriteLog "End Discogs Request"
	End If


	Dim AlbumArt
	CheckCover = False
	SmallCover = False
	Set AlbumArt = FirstTrack.AlbumArt
	If CheckSaveImage = 0 Or CheckSaveImage = 1 Then
		If AlbumArt.Count = 0 Then
			If CheckSmallCover = True Then
				SmallCover = True
			Else
				CheckCover = True
			End If
		End If
	Else
		If CheckSmallCover = True Then
			SmallCover = True
		Else
			CheckCover = True
		End If
	End If

	Dim itmAlbum
	For a = 0 To SDB.Tools.WebSearch.NewTracks.Count - 1
		Set tmp = SDB.Tools.WebSearch.NewTracks.item(a)
		Set itmAlbum = tmp.Album
		WriteLog "Disc=" & tmp.DiscNumberStr & " / Track=" & tmp.TrackOrderStr & " / AlbumID=" & itmAlbum.ID & " / Artist=" & tmp.ArtistName & " / Album=" & tmp.AlbumName & " / Title=" & tmp.Title
	Next
	WriteLog " "

	REM If SearchArtist = "" Or SearchAlbum = "" Then Exit Sub

	If UseMetalArchives = True Then
		FindResults2(SavedSearchTerm)
	Else
		If AccessToken = "" Or AccessTokenSecret = "" Then
			MsgBox("You don't have a valid AccessToken !")
			WebBrowser.Common.DestroyControl      ' Destroy the external control
			Set WebBrowser = Nothing              ' Release global variable
			SDB.Objects("WebBrowser") = Nothing

			Set ini = Nothing
			Set ResultsReleaseID = Nothing
			Script.UnregisterAllEvents
		Else
			FindResults(SavedSearchTerm)
		End If
	End If

End Sub


Sub FindResults2(SearchTerm)

	Dim TXTBegin, TXTEnd, ResponseHTML, ReleaseDesc, searchURL, ErrorMessage, i, tmp
	
	Dim oXMLHTTP, MAReleases
	
	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
	SearchTerm = LTrim(SearchTerm)
	WriteLog("Start FindResults2")
	WriteLog("SavedSearchArtist=" & SavedSearchArtist)
	WriteLog("SavedSearchAlbum=" & SavedSearchAlbum)
	
	Set Results = SDB.NewStringList
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""
	
	Set FirstTrack = SDB.Tools.WebSearch.NewTracks.item(0)

	'searchURL = "http://www.metal-archives.com/search/ajax-advanced/searching/albums?bandName=" & URLEncodeUTF8(CleanSearchString(SavedSearchArtist)) & "&exactBandMatch=1&releaseTitle=" & URLEncodeUTF8(CleanSearchString(SavedSearchAlbum)) & "&exactReleaseMatch=1&releaseYearFrom=&releaseMonthFrom=&releaseYearTo=&releaseMonthTo=&country=&location=&releaseLabelName=&genre=#albums"
	searchURL = "http://www.metal-archives.com/search/ajax-advanced/searching/albums/?bandName=" & URLEncodeUTF8(CleanSearchString(SavedSearchArtist)) & "&exactBandMatch=1&releaseTitle=" & URLEncodeUTF8(CleanSearchString(SavedSearchAlbum)) & "&exactReleaseMatch=1&releaseYearFrom=&releaseMonthFrom=&releaseYearTo=&releaseMonthTo=&country=&location=&releaseLabelName=&genre="
	WriteLog searchURL
	oXMLHTTP.open "GET", searchURL, false
	oXMLHTTP.send()
	if oXMLHTTP.Status = 200 Then
		ResponseHTML = oXMLHTTP.responseText
		Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
		Dim f
		Set f = fso.OpenTextFile(SDB.ScriptsPath&"test.log", 2, true, -1)
		f.WriteLine ResponseHTML
		f.Close
		TXTBegin = InStr(ResponseHTML, "iTotalDisplayRecords")
		If TXTBegin > 1 Then
			ResponseHTML = Mid(ResponseHTML, TXTBegin + 23)
			TXTEnd = InStr(ResponseHTML, ",")
			MAReleases = Left(ResponseHTML, TXTEnd -1)
			WriteLog "Anzahl=" & MAReleases
			'WriteLog "Anzahl=" & ResponseHTML
			For i = 1 to MAReleases
				TXTBegin = InStr(ResponseHTML, "\" & Chr(34) & ">")
				If TXTBegin > 1 Then
					ResponseHTML = Mid(ResponseHTML, TXTBegin + 3)
					TXTEnd = InStr(responseHTML, "</a>")

					WriteLog "Band=" & Left(ResponseHTML, TXTEnd - 1)
					ReleaseDesc = Left(ResponseHTML, TXTEnd - 1)
					
					TXTBegin = InStr(ResponseHTML, "href=\")
					ResponseHTML = Mid(ResponseHTML, TXTBegin + 7)
					TXTEnd = InStr(ResponseHTML, "\" & Chr(34) & ">")
					WriteLog "Album-URL=" & Left(ResponseHTML, TXTEnd - 1)
					ResultsReleaseID.Add Left(ResponseHTML, TXTEnd - 1)
					
					TXTBegin = TXTEnd + 3
					ResponseHTML = Mid(ResponseHTML, TXTBegin)
					TXTEnd = InStr(responseHTML, "</a>")
					WriteLog "Album=" & Left(ResponseHTML, TXTEnd - 1)
					ReleaseDesc = ReleaseDesc & " / " & Left(ResponseHTML, TXTEnd - 1)

					ResponseHTML = Mid(ResponseHTML, TXTEnd)
					TXTBegin = InStr(ResponseHTML, ",")
					ResponseHTML = Mid(ResponseHTML, TXTBegin)
					TXTBegin = InStr(ResponseHTML, Chr(34))
					ResponseHTML = Mid(ResponseHTML, TXTBegin + 1)
					TXTEnd = InStr(responseHTML, Chr(34))
					WriteLog "Type=" & Left(ResponseHTML, TXTEnd - 1)
					ReleaseDesc = ReleaseDesc & " / " & Left(ResponseHTML, TXTEnd - 1)
					Results.Add ReleaseDesc
				End If
			Next
		End If
	End If
	SDB.Tools.WebSearch.SetSearchResults Results
	SDB.Tools.WebSearch.ResultIndex = 0

	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	End If
End Sub








Sub FindResults(SearchTerm)



	SearchTerm = LTrim(SearchTerm)
	WriteLog("Start FindResults")
	WriteLog("SearchTerm=" & SearchTerm)
	WriteLog("SavedSearchArtist=" & SavedSearchArtist)
	WriteLog("SavedSearchAlbum=" & SavedSearchAlbum)
	
	Dim ErrorMessage, FilterFound, a, searchURL, searchURL_F, searchURL_L

	Set Results = SDB.NewStringList
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""

	Set FirstTrack = SDB.Tools.WebSearch.NewTracks.item(0)

	If (InStr(SearchTerm," - [search by release id]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search by release id]")-1)
	End If

	If (InStr(SearchTerm," - [search by release url]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search by release url]")-1)
	End If

	If (InStr(SearchTerm," - [currently tagged with this release]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [currently tagged with this release]")-1)
	End If

	If (InStr(SearchTerm," - [search returned no results]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search returned no results]")-1)
	End If

	If (InStr(SearchTerm," - [search that yielded error]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search that yielded error]")-1)
	End If

	' Handle direct urls

	If (InStr(SearchTerm,"/master/") > 0) Then
		CurrentLoadType = "Master Release"
		LoadMasterResults Mid(SearchTerm,InStrRev(SearchTerm,"/")+1)
		Exit Sub
	End If

	'Will not be longer supported, cause the artist url at Discogs have no more artist-id
	Rem If (InStr(SearchTerm,"/artist/") > 0) Then
		Rem CurrentLoadType = "Releases of Artist"
		Rem WriteLog("Direct Artist url (ArtistId)=" & Mid(SearchTerm,InStrRev(SearchTerm,"/")+1))
		Rem LoadArtistResults Mid(SearchTerm,InStrRev(SearchTerm,"/")+1)
		Rem Exit Sub
	Rem End If

	If (InStr(SearchTerm,"/label/") > 0) Then
		CurrentLoadType = "Releases of Label"
		LoadLabelResults Mid(SearchTerm,InStrRev(SearchTerm,"/")+1)
		Exit Sub
	End If

	If SearchTerm = "" Then
		ErrorMessage = "No search term"
	ElseIf IsNumeric(SearchTerm) Then
		Results.Add SearchTerm & " - [search by release id]"
		ResultsReleaseID.Add SearchTerm
	ElseIf (InStr(SearchTerm,"/release/") > 0) Then
		Results.Add SearchTerm & " - [search by release url]"
		ResultsReleaseID.Add Mid(SearchTerm,InStrRev(SearchTerm,"/")+1)
	Else

		If IsNumeric(SavedReleaseId) Then
			Results.Add FirstTrack.Artist.Name & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
			ResultsReleaseID.Add get_release_ID(FirstTrack) 'get saved Release_ID from User-Defined Custom-Tag
		End If

		
		If SavedSearchArtist <> "" And SavedSearchAlbum <> "" Then
			searchURL = CleanSearchString(SearchTerm)
			searchURL_F = "http://api.discogs.com/database/search?q="
			searchURL_L = "&type=release&per_page=100"
		ElseIf SavedSearchArtist = "" And SavedSearchAlbum <> "" Then
			searchURL = CleanSearchString(SavedSearchAlbum)
			searchURL_F = "http://api.discogs.com/database/search?type=release&title="
			searchURL_L = "&per_page=100"
		ElseIf SavedSearchArtist <> "" And SavedSearchAlbum = "" Then
			searchURL = CleanSearchString(SavedSearchArtist)
			searchURL_F = "http://api.discogs.com/database/search?type=release&artist="
			searchURL_L = "&per_page=100"
		Else
			searchURL = CleanSearchString(SearchTerm)
			searchURL_F = "http://api.discogs.com/database/search?q="
			searchURL_L = "&type=release&per_page=100"
		End If

		WriteLog("Complete searchURL=" & searchURL_F & searchURL & searchURL_L)

		JSONParser_find_result searchURL, "results", searchURL_F, searchURL_L

		If ResultsReleaseID.Count = 0 Then
			FilterFound = False
			If FilterCountry = "Use Country Filter" Then
				For a = 1 To CountryList.Count - 1
					If CountryFilterList.Item(a) = "1" Then
						FilterFound = True
						Exit For
					End If
				Next
				If FilterFound = False Then
					ErrorMessage = "No Country Filter set !"
				Else
					ErrorMessage = "Search returned no results"
				End If
				Results.Add SearchTerm & " - [search returned no results]"
				ResultsReleaseID.Add SearchTerm
			End If
			FilterFound = False
			If FilterMediaType = "Use MediaType Filter" Then
				For a = 1 To MediaTypeList.Count - 1
					If MediaTypeFilterList.Item(a) = "1" Then
						FilterFound = True
						Exit For
					End If
				Next
				If FilterFound = False Then
					If ErrorMessage = "" Then
						ErrorMessage = "No MediaType Filter set !"
					Else
						ErrorMessage = ErrorMessage & vbCrLf & "No MediaType Filter set !"
					End If
				End If
			End If
			FilterFound = False
			If FilterMediaFormat = "Use MediaFormat Filter" Then
				For a = 1 To MediaFormatList.Count - 1
					If MediaFormatFilterList.Item(a) = "1" Then
						FilterFound = True
						Exit For
					End If
				Next
				If FilterFound = False Then
					If ErrorMessage = "" Then
						ErrorMessage = "No MediaFormat Filter set !"
					Else
						ErrorMessage = ErrorMessage & vbCrLf & "No MediaFormat Filter set !"
					End If
				End If
			End If
			FilterFound = False
			If FilterYear = "Use Year Filter" Then
				For a = 1 To YearList.Count - 1
					If YearFilterList.Item(a) = "1" Then
						FilterFound = True
						Exit For
					End If
				Next
				If FilterFound = False Then
					If ErrorMessage = "" Then
						ErrorMessage = "No Year Filter set !"
					Else
						ErrorMessage = ErrorMessage & vbCrLf & "No Year Filter set !"
					End If
				End If
			End If

			If ErrorMessage = "" Then
				ErrorMessage = "Search returned no results"
			End If
			Results.Add SearchTerm & " - [search returned no results]"
			ResultsReleaseID.Add SearchTerm
		End If

	End If

	SDB.Tools.WebSearch.SetSearchResults Results
	SDB.Tools.WebSearch.ResultIndex = 0

	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	End If

End Sub

Sub LoadMasterResults(MasterId)

	Dim ErrorMessage, masterURL
	WriteLog("MasterResult")

	Set Results = SDB.NewStringList
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""

	If MasterId = "" Then
		ErrorMessage = "Cannot load empty master release"
	Else
		If IsNumeric(SavedReleaseId) Then
			Set FirstTrack = SDB.Tools.WebSearch.NewTracks.item(0)
			Results.Add FirstTrack.Artist.Name & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
			ResultsReleaseID.Add get_release_ID(FirstTrack) 'get saved Release_ID from User-Defined Custom-Tag
		End If

		masterURL = MasterId
		JSONParser_find_result masterURL, "versions", "http://api.discogs.com/masters/", "/versions"
	End If

	SDB.Tools.WebSearch.SetSearchResults Results
	SDB.Tools.WebSearch.ResultIndex = 0

	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	End If

End Sub


Sub LoadArtistResults(ArtistId)

	Dim ErrorMessage
	Dim artistURL
	
	Set Results = SDB.NewStringList
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""

	If ArtistId = "" Then
		ErrorMessage = "Cannot load empty artist"
	Else
		If IsNumeric(SavedReleaseId) Then
			Set FirstTrack = SDB.Tools.WebSearch.NewTracks.item(0)
			Results.Add FirstTrack.Artist.Name & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
			ResultsReleaseID.Add get_release_ID(FirstTrack) 'get saved Release_ID from User-Defined Custom-Tag
		End If

		artistURL = ArtistId
		JSONParser_find_result artistURL, "releases", "http://api.discogs.com/artists/", "/releases?per_page=100"
	End If

	SDB.Tools.WebSearch.SetSearchResults Results
	SDB.Tools.WebSearch.ResultIndex = 0

	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	End If

End Sub


Sub LoadLabelResults(LabelId)

	Dim ErrorMessage, labelURL

	Set Results = SDB.NewStringList
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""

	If LabelId = "" Then
		ErrorMessage = "Cannot load empty label"
	Else
		If IsNumeric(SavedReleaseId) Then
			Set FirstTrack = SDB.Tools.WebSearch.NewTracks.item(0)
			Results.Add FirstTrack.Artist.Name & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
			ResultsReleaseID.Add get_release_ID(FirstTrack) 'get saved Release_ID from User-Defined Custom-Tag
		End If

		labelURL = LabelId
		JSONParser_find_result labelURL, "releases", "http://api.discogs.com/labels/", "/releases?per_page=100"
	End If

	SDB.Tools.WebSearch.SetSearchResults Results
	SDB.Tools.WebSearch.ResultIndex = 0

	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	End If

End Sub


'For reloading results
Sub ReloadResults

	Dim Tracks, TracksNum, DiscogsTracksNum, TracksCD, ArtistTitles, InvolvedArtists, Lyricists, Composers, Conductors, Producers, Durations
	Dim AlbumArtist, AlbumArtistTitle, AlbumLyricist, AlbumComposer, AlbumConductor, AlbumProducer, AlbumInvolved, AlbumFeaturing, AlbumTitle
	Dim track, currentTrack, position, artist, currentArtist, artistName, extraArtist, extra
	Dim currentImage, currentLabel, currentFormat, theMaster, i, g, l, s, f, d
	Dim ReleaseDate, ReleaseSplit, theLabels, theCatalogs, theCountry, theFormat
	Dim Genres, Styles, Comment, DataQuality
	Dim NoSubTrackUsing, oldSubTrackNumber

	Set Tracks = SDB.NewStringList
	Set TracksNum = SDB.NewStringList
	Set DiscogsTracksNum = SDB.NewStringList
	Set TracksCD = SDB.NewStringList
	Set ArtistTitles = SDB.NewStringList
	Set InvolvedArtists = SDB.NewStringList
	Set Lyricists = SDB.NewStringList
	Set Composers = SDB.NewStringList
	Set Conductors = SDB.NewStringList
	Set Producers = SDB.NewStringList
	Set Durations = SDB.NewStringList

	'----------------------------------DiscogsImages----------------------------------------
	Set SaveImage = SDB.NewStringList
	Set SaveImageType = SDB.NewStringList
	Set FileNameList = SDB.NewStringList
	ImagesCount = 0
	'----------------------------------DiscogsImages----------------------------------------


	If OptionsChanged = True Then
		OptionsChanged = False
		WriteOptions()
	End If

	SDB.Tools.WebSearch.ClearTracksData   ' Tell MM to disregard any previously set tracks' data
	If not isnull(CurrentRelease) Then

		AlbumArtist = ""
		AlbumArtistTitle = ""
		AlbumLyricist = ""
		AlbumComposer = ""
		AlbumConductor = ""
		AlbumProducer = ""
		AlbumInvolved = ""
		AlbumArtURL = ""
		AlbumArtThumbNail = ""
		AlbumFeaturing = ""
		LastDisc = ""

		Dim iTrackNum, iSubTrack, cSubTrack, subTrackTitle
		Dim trackName, t, pos
		Dim role, rolea, currentRole, NoSplit, zahl, zahltemp, zahl2, zahltemp2
		Dim CharSeparatorSubTrack
		ReDim Involved_R(0)
		Dim tmp
		Dim tmp2
		Dim rTrack
		Dim LeadingZeroTrackPosition
		ReDim TrackRoles(0)
		ReDim TrackArtist2(0)
		ReDim TrackPos(0)
		ReDim Title_Position(0)
		SavedArtistId = ""
		SavedLabelId = ""
		LeadingZeroTrackPosition = False

		'Get Track-List
		For Each track In CurrentRelease("tracklist")
			Set currentTrack = CurrentRelease("tracklist")(track)
			position = currentTrack("position")
			DiscogsTracksNum.Add position
			position = exchange_roman_numbers(position)
			ReDim Preserve Title_Position(UBound(Title_Position)+1)
			Title_Position(UBound(Title_Position)) = position
		Next


		'Check for leading zero in track-position
		LeadingZeroTrackPosition = CheckLeadingZeroTrackPosition(Title_Position(1))

		' Get artist title
		For Each artist in CurrentRelease("artists")
			Set currentArtist = CurrentRelease("artists")(artist)
			WriteLog("currentArtist=" & currentArtist("name"))
			If Not CheckUseAnv And currentArtist("anv") <> "" Then
				artistName = CleanArtistName(currentArtist("anv"))
				' !!!!!artistName <- currentArtist
			Else
				artistName = CleanArtistName(currentArtist("name"))
				' !!!!!artistName <- currentArtist
			End If
			If SavedArtistId = "" Then SavedArtistId = currentArtist("id")

			If (AlbumArtist = "") Then
				AlbumArtist = artistName
			End If

			Writelog("SavedArtistId=" & SavedArtistId)
			AlbumArtistTitle = AlbumArtistTitle & artistName

			If currentArtist("join") <> "" Then
				tmp = currentArtist("join")
				If tmp = "," Then
					AlbumArtistTitle = AlbumArtistTitle & ArtistSeparator
				ElseIf LookForFeaturing(tmp) And CheckFeaturingName Then
					If TxtFeaturingName = "," or TxtFeaturingName = ";" Then
						AlbumArtistTitle = AlbumArtistTitle & TxtFeaturingName & " "
					Else
						AlbumArtistTitle = AlbumArtistTitle & " " & TxtFeaturingName & " "
					End If
				Else
					AlbumArtistTitle = AlbumArtistTitle & " " & currentArtist("join") & " "
				End If
			End If
		Next

		Writelog("AlbumArtistTitle=" & AlbumArtistTitle)

		If Right(AlbumArtistTitle, 3) = " , " Then AlbumArtistTitle = Left(AlbumArtistTitle, Len(AlbumArtistTitle)-3)

		If (Not CheckAlbumArtistFirst) Then
			AlbumArtist = AlbumArtistTitle
		End If

		If AlbumArtist = "Various" And CheckVarious Then
			AlbumArtist = TxtVarious
		End If
		If AlbumArtistTitle = "Various" And CheckVarious Then
			AlbumArtistTitle = TxtVarious
		End If


		WriteLog " "
		WriteLog "ExtraArtists"
		If currentRelease.Exists("extraartists") Then
			For Each extraArtist In CurrentRelease("extraartists")
				WriteLog " "
				Set currentArtist = CurrentRelease("extraartists")(extraArtist)
				If currentArtist("tracks") = "" Then
					If (currentArtist("anv") <> "") And Not CheckUseAnv Then
						artistName = CleanArtistName(currentArtist("anv"))
					Else
						artistName = CleanArtistName(currentArtist("name"))
					End If
					WriteLog ("ArtistName=" & artistName)
					WriteLog "Without Track Info"
					role = currentArtist("role")
					NoSplit = False
					If InStr(role, ",") = 0 Then
						currentRole = role
						zahl = 1
						NoSplit = True
					Else
						rolea = CheckSpecialRole(role)
						REM rolea = Split(role, ",")
						zahl = UBound(rolea)
					End If

					WriteLog ("Role count=" & zahl)
					For zahltemp = 1 To zahl
						If NoSplit = False Then
							currentRole = Trim(rolea(zahltemp))
						End If
						WriteLog ("currentRole=" & currentRole)
						If LookForFeaturing(currentRole) Then
							WriteLog ("Featuring found")
							If InStr(AlbumFeaturing, artistName) = 0 Then
								If AlbumFeaturing = "" Then
									If CheckFeaturingName Then
										AlbumFeaturing = TxtFeaturingName & " " & artistName
									Else
										AlbumFeaturing = currentRole & " " & artistName
									End If
								Else
									AlbumFeaturing = AlbumFeaturing & Separator & artistName
								End If
							End If
						Else
							Do
								tmp = searchKeyword(LyricistKeywords, currentRole, AlbumLyricist, artistName)
								If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
									AlbumLyricist = tmp
									WriteLog ("AlbumLyricist=" & AlbumLyricist)
									Exit Do
								ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
									Exit Do
								End If
								tmp = searchKeyword(ConductorKeywords, currentRole, AlbumConductor, artistName)
								If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
									AlbumConductor = tmp
									WriteLog ("AlbumConductor=" & AlbumConductor)
									Exit Do
								ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
									Exit Do
								End If
								tmp = searchKeyword(ProducerKeywords, currentRole, AlbumProducer, artistName)
								If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
									AlbumProducer = tmp
									WriteLog ("AlbumProducer=" & AlbumProducer)
									Exit Do
								ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
									Exit Do
								End If
								tmp = searchKeyword(ComposerKeywords, currentRole, AlbumComposer, artistName)
								If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
									AlbumComposer = tmp
									WriteLog ("AlbumComposer=" & AlbumComposer)
									Exit Do
								ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
									Exit Do
								End If
								tmp2 = search_involved(Involved_R, currentRole)
								If tmp2 = -1 Then
									ReDim Preserve Involved_R(UBound(Involved_R)+1)
									Involved_R(UBound(Involved_R)) = currentRole & ": " & artistName
									WriteLog ("New Role: " & currentRole & ": " & artistName)
								Else
									If InStr(Involved_R(tmp2), artistName) = 0 Then
										Involved_R(tmp2) = Involved_R(tmp2) & ", " & artistName
										WriteLog ("Role updated: " & Involved_R(tmp2))
									Else
										WriteLog ("artist already inside role")
									End If
								End If
								Exit Do
							Loop While True
						End If
					Next
				Else
					If Not CheckUseAnv And currentArtist("anv") <> "" Then
						artistName = CleanArtistName(currentArtist("anv"))
					Else
						artistName = CleanArtistName(currentArtist("name"))
					End If
					WriteLog ("ArtistName=" & artistName)
					role = currentArtist("role")
					rTrack = currentArtist("tracks")
					WriteLog ("Track(s)=" & rTrack)
					WriteLog ("Role(s)=" & role)
					NoSplit = False
					If InStr(role, ",") <> 0 Then
						REM rolea = Split(role, ",")
						rolea = CheckSpecialRole(role)
						zahl = UBound(rolea)
					ElseIf InStr(role, " & ") <> 0 Then
						rolea = Split(role, "&")
						zahl = UBound(rolea)
					Else
						involvedRole = Trim(role)
						zahl = 1
						NoSplit = True
					End If
					For zahltemp = 1 To zahl
						If NoSplit = False Then
							involvedRole = Trim(rolea(zahltemp))
						End If
						WriteLog ("involvedRole=" & involvedRole)
						If InStr(rTrack, ",") = 0 And InStr(rTrack, " to ") = 0 And InStr(rTrack, " & ") = 0 Then
							currentTrack = rTrack
							Add_Track_Role currentTrack, artistName, involvedRole, TrackRoles, TrackArtist2, TrackPos
						End If
						If InStr(rTrack, ",") <> 0 Then
							tmp = Split(rTrack, ",")
							zahl2 = UBound(tmp)
							For zahltemp2 = 0 To zahl2
								currentTrack = Trim(tmp(zahltemp2))
								If InStr(currentTrack, " to ") <> 0 Then
									Track_from_to currentTrack, artistName, involvedRole, Title_Position, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition
								Else
									Add_Track_Role currentTrack, artistName, involvedRole, TrackRoles, TrackArtist2, TrackPos
								End If
							Next
						ElseIf InStr(rTrack, " to ") <> 0 Then
							currentTrack = Trim(rTrack)
							Track_from_to currentTrack, artistName, involvedRole, Title_Position, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition
						ElseIf InStr(rTrack, " & ") <> 0 Then
							tmp = Split(rTrack, " & ")
							zahl2 = UBound(tmp)
							For zahltemp2 = 0 To zahl2
								currentTrack = Trim(tmp(zahltemp2))
								Add_Track_Role currentTrack, artistName, involvedRole, TrackRoles, TrackArtist2, TrackPos
							Next
						End If
					Next
				End If
			Next
		End If
		' Get track titles and track artists

		iAutoTrackNumber = 1
		iAutoDiscNumber = 1
		iTrackNum = 0
		iSubTrack = 0
		cSubTrack = -1
		subTrackTitle = ""
		CharSeparatorSubTrack = 0
		Rem CharSeparatorSubTrack: 0 = nothing    1 = "."     2 = a-z
		Rem subTrackStart = 1 '0 = Song -1    1 = First Song

		'Workaround for using "." as separator at discogs -----------------------------------------------------------------------------------------------------------
		tmp = 0 : tmp2 = 0
		NoSubTrackUsing = False
		For Each t In CurrentRelease("tracklist")
			Set currentTrack = CurrentRelease("tracklist")(t)
			position = currentTrack("position")
			If position <> "" Then
				tmp2 = tmp2 + 1
			End If
			If InStr(position, ".") <> 0 Then tmp = tmp + 1
		Next
		If tmp = tmp2 Then NoSubTrackUsing = True	'all tracks have "." in position tag, this can't be a subtrack
		'Workaround for using "." as separator at discogs -----------------------------------------------------------------------------------------------------------


		For Each t In CurrentRelease("tracklist")
			Set currentTrack = CurrentRelease("tracklist")(t)

			position = currentTrack("position")
			If Right(position, 1) = "." Then position = Left(position, Len(position)-1)
			If NoSubTrackUsing = True Then position = Replace(position, ".", "-")
			trackName = PackSpaces(DecodeHtmlChars(currentTrack("title")))
			Durations.Add currentTrack("duration")
			position = exchange_roman_numbers(position)
			WriteLog " "
			WriteLog " "
			WriteLog "Position=" & position

			pos = 0
			If InStr(LCase(position), "-") > 0 Then
				pos = InStr(LCase(position), "-")
			End If
			' Here comes the new track/disc numbering methods

			If position <> "" Then
				If CheckTurnOffSubTrack = False Then
					If (cSubTrack <> -1 And InStr(LCase(position), ".") = 0 And CharSeparatorSubTrack = 1) Or (cSubTrack <> -1 And IsNumeric(Right(position, 1)) And CharSeparatorSubTrack = 2) Then
						WriteLog "End of Subtrack found"
						If SubTrackNameSelection = False Then
							Tracks.Item(cSubTrack) = Tracks.Item(cSubTrack) & " (" & subTrackTitle & ")"
						Else
							Tracks.Item(cSubTrack) = subTrackTitle
						End If
						cSubTrack = -1
						subTrackTitle = ""
						CharSeparatorSubTrack = 0
					End If

					If NoSubTrackUsing = False Then
						WriteLog "Calling Subtrack Function"
						'SubTrack Function ---------------------------------------------------------
						If InStr(LCase(position), ".") > 0 Then
							CharSeparatorSubTrack = 1
						ElseIf Not IsNumeric(Right(position, 1)) And Len(position) > 1 Then
							CharSeparatorSubTrack = 2
						End If
						If CharSeparatorSubTrack <> 0 Then
							If cSubTrack <> -1 Then 'more subtrack
								If CharSeparatorSubTrack = 1 Then
									tmp = Split(position, ".")
									If oldSubTrackNumber <> tmp(0) Then
										If SubTrackNameSelection = False Then
											Tracks.Item(cSubTrack) = Tracks.Item(cSubTrack) & " (" & subTrackTitle & ")"
										Else
											Tracks.Item(cSubTrack) = subTrackTitle
										End If
										cSubTrack = -1
										subTrackTitle = ""
										REM CharSeparatorSubTrack = 0
									End If
								ElseIf CharSeparatorSubTrack = 2 Then
									tmp2 = FindSubTrackSplit(position)
									If oldSubTrackNumber <> tmp2 Then
										If SubTrackNameSelection = False Then
											Tracks.Item(cSubTrack) = Tracks.Item(cSubTrack) & " (" & subTrackTitle & ")"
										Else
											Tracks.Item(cSubTrack) = subTrackTitle
										End If
										cSubTrack = -1
										subTrackTitle = ""
										REM CharSeparatorSubTrack = 0
									End If
								End If
							End If
							If cSubTrack = -1 Then 'new subtrack
								WriteLog("New SubTrack found")
								If SubTrackNameSelection = False Then
									cSubTrack = iTrackNum - 1
								Else
									cSubTrack = iTrackNum
								End If
								If CharSeparatorSubTrack = 1 Then
									tmp = Split(position, ".")
									oldSubTrackNumber = tmp(0)
								ElseIf CharSeparatorSubTrack = 2 Then
									oldSubTrackNumber = FindSubTrackSplit(position)
									If oldSubTrackNumber = "" Then oldSubTrackNumber = position
								End If
								WriteLog ("oldSubTrackNumber=" & oldSubTrackNumber)
							End If
							
							If subTrackTitle = "" Then
								subTrackTitle = trackName
								If SubTrackNameSelection = False Then
									UnselectedTracks(iTrackNum) = "x"
								Else
									UnselectedTracks(iTrackNum) = ""
								End If
							Else
								subTrackTitle = subTrackTitle & ", " & trackName
								UnselectedTracks(iTrackNum) = "x"
							End If

							'SubTrack Function ---------------------------------------------------------
						End If
					End If
				End If
				If pos > 0 And CheckNoDisc = False Then ' Disc Number Included
					If CheckForceNumeric Then
						If Left(position,2) = "CD" Then
							If Mid(position,3,1) = "-" Then
								iAutoDiscNumber = 1
							Else
								If iAutoDiscNumber <> Mid(position,3,1) Then
									iAutoTrackNumber = 1
								End If
							End If
						End If
						If Left(position,2) <> "CD" And IsInteger(Left(position,pos-1)) Then
							If Int(iAutoDiscNumber) <> Int(Left(position,pos-1)) Then
								iAutoTrackNumber = 1
							End If
						End If
						If UnselectedTracks(iTrackNum) <> "x" Then
							If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
								tracksNum.Add "0" & iAutoTrackNumber
							Else
								tracksNum.Add iAutoTrackNumber
							End If
							iAutoTrackNumber = iAutoTrackNumber + 1
						Else
							tracksNum.Add ""
						End If
					Else
						If pos > 0 Then
							If Len(Mid(position, pos+1)) > 1 Then	'minimum 2 Char after -  (1-1a, 1-II, 1-12)
								If IsInteger(Mid(position, pos+1, 1)) And Not IsInteger(Right(position, 1)) Then	'First is a Number, Char at the end (1-1a, 1-1b, 1-1c,...) = Sub-Track !
									If Mid(position,pos + 1, Len(position) - pos - 1) < 10 And CheckLeadingZero = True Then
										tracksNum.Add "0" & Right(position,Len(position)-pos)
									Else
										tracksNum.Add Right(position,Len(position)-pos)
									End If
								ElseIf IsInteger(Mid(position, pos+1)) Then		'no char at all (1-01, 1-02, 1-12)
									If CheckLeadingZero = True And Right(position,Len(position)-pos) < 10 Then
										tracksNum.Add "0" & Right(position,Len(position)-pos)
									Else
										tracksNum.Add Right(position,Len(position)-pos)
									End If
								Else
									tracksNum.Add Right(position,Len(position)-pos)
								End If
							ElseIf Len(Mid(position, pos+1)) = 1 Then	'1 Char after -  (1-1, 1-I, 1-2)
								If IsInteger(Mid(position, pos+1)) Then
									If CheckLeadingZero = True And Mid(position, pos+1) < 10 Then
										tracksNum.Add "0" & Mid(position, pos+1)
									Else
										tracksNum.Add Mid(position, pos+1)
									End If
								Else
									tracksNum.Add Mid(position, pos+1)
								End If
							End If
						End If
						If UnselectedTracks(iTrackNum) <> "x" Then
							If IsInteger(Right(position,len(position)-pos)) Then
								iAutoTrackNumber = Right(position,len(position)-pos) + 1
							Else
								iAutoTrackNumber = iAutoTrackNumber + 1
							End If
						End If
					End If
					If Left(position,2) = "CD" Then
						If Mid(position,3,1) = "-" Then
							'Or Mid(position,3,1) = "." Then
							iAutoDiscNumber = 1
						Else
							iAutoDiscNumber = Mid(position,3,1)
						End If
					End If
					If Left(position,2) <> "CD" Then iAutoDiscNumber = Left(position,pos-1)
					tracksCD.Add iAutoDiscNumber
				Else ' Apply Track Numbering Schemes
					If Not CheckSidesToDisc Or IsInteger(Left(position,1)) Then
						If CheckForceNumeric Then
							If UnselectedTracks(iTrackNum) <> "x" Then
								If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
									tracksNum.Add "0" & iAutoTrackNumber
								Else
									tracksNum.Add iAutoTrackNumber
								End If
								iAutoTrackNumber = iAutoTrackNumber + 1
							Else
								tracksNum.Add ""
							End If
						Else
							If CheckLeadingZero = True And IsInteger(position) Then
								If position < 10 Then
									tracksNum.Add "0" & position
								Else
									tracksNum.Add position
								End If
							Else
								tracksNum.Add position
							End If
							If UnselectedTracks(iTrackNum) <> "x" Then
								If IsInteger(position) Then
									iAutoTrackNumber = position + 1
								Else
									iAutoTrackNumber = iAutoTrackNumber + 1
								End If
							End If
						End If
						If CheckForceDisc Then
							tracksCD.Add iAutoDiscNumber
						Else
							tracksCD.Add ""
						End If
					Else
						If Len(position) = 1 Then ' Only side is specified
							If CheckLeadingZero = True Then
								tracksNum.Add "01"
							Else
								tracksNum.Add "1"
							End If
							If 	LastDisc <> position Then
								If 	LastDisc <> "" Then
									iAutoDiscNumber = iAutoDiscNumber + 1
								End If
								LastDisc = position
							End If
							If CheckForceNumeric Then
								tracksCD.Add iAutoDiscNumber
							Else
								tracksCD.Add position
							End If
						ElseIf Len(position) = 2 Then
							If IsInteger(Mid(position,2,1)) And Not IsInteger(Mid(position,1,1)) Then
								' First is Side Second is Track
								If CheckLeadingZero = True And Mid(position,2) < 10 Then
									tracksNum.Add "0" & Mid(position,2)
								Else
									tracksNum.Add Mid(position,2)
								End If
								If 	LastDisc <>  Left(position,1) Then
									If 	LastDisc <> "" Then
										iAutoDiscNumber = iAutoDiscNumber + 1
									End If
									LastDisc = Left(position,1)
								End If
								If CheckForceNumeric Then
									tracksCD.Add iAutoDiscNumber
								Else
									tracksCD.Add Left(position,1)
								End If
							Else ' Two byte side
								tracksNum.Add "1"
								If 	LastDisc <>  position Then
									If 	LastDisc <> "" Then
										iAutoDiscNumber = iAutoDiscNumber + 1
									End If
									LastDisc = position
								End If
								If CheckForceNumeric Then
									tracksCD.Add iAutoDiscNumber
								Else
									tracksCD.Add position
								End If
							End If
						Else ' More than 2 bytes
							If IsInteger(Mid(position,2)) And CheckNoDisc = False Then
							'First is Side Latter is Track
								tracksNum.Add Mid(position,2)
								If 	LastDisc <>  Left(position,1) Then
									If 	LastDisc <> "" Then
										iAutoDiscNumber = iAutoDiscNumber + 1
									End If
									LastDisc = Left(position,1)
								End If
								If CheckForceNumeric Then
									tracksCD.Add iAutoDiscNumber
								Else
									tracksCD.Add Left(position,1)
								End If
							ElseIf IsInteger(Mid(position,3)) And CheckNoDisc = False Then
								' Two Byte Side, Latter is Track
								tracksNum.Add Mid(position,3)
								If 	LastDisc <>  Left(position,2) Then
									If 	LastDisc <> "" Then
										iAutoDiscNumber = iAutoDiscNumber + 1
									End If
									LastDisc = Left(position,2)
								End If
								If CheckForceNumeric Then
									tracksCD.Add iAutoDiscNumber
								Else
									tracksCD.Add Left(position,2)
								End If
							Else ' More than two non numeric bytes!
								If CheckNoDisc = False Then
									tracksNum.Add position
									tracksCD.Add ""
								Else
									If CheckForceNumeric Then
										If UnselectedTracks(iTrackNum) <> "x" Then
											If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
												tracksNum.Add "0" & iAutoTrackNumber
											Else
												tracksNum.Add iAutoTrackNumber
											End If
											iAutoTrackNumber = iAutoTrackNumber + 1
										Else
											tracksNum.Add ""
										End If
									Else
										If UnselectedTracks(iTrackNum) <> "x" Then
											If IsInteger(position) Then
												tracksNum.Add iAutoTrackNumber
												iAutoTrackNumber = position + 1
											Else
												tracksNum.Add iAutoTrackNumber
												iAutoTrackNumber = iAutoTrackNumber + 1
											End If
										End If
									End If
									tracksCD.Add ""
								End If
							End If
						End If
					End If
				End If
			ElseIf currentTrack("duration") = "" And currentTrack("title") = "-" Then
				tracksNum.Add ""
				tracksCD.Add ""
				UnselectedTracks(iTrackNum) = "x"
			Else ' Nothing specified
				If CheckForceNumeric and UnselectedTracks(iTrackNum) <> "x" Then
					If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
						tracksNum.Add "0" & iAutoTrackNumber
					Else
						tracksNum.Add iAutoTrackNumber
					End If
					iAutoTrackNumber = iAutoTrackNumber + 1
				Else
					tracksNum.Add ""
				End If
				If CheckForceDisc Then
					tracksCD.Add iAutoDiscNumber
				Else
					tracksCD.Add ""
				End If
			End If

			Dim involvedArtist, involvedTemp, involvedRole
			Dim TrackInvolvedPeople, TrackComposers, TrackConductors, TrackProducers, TrackLyricists, TrackFeaturing
			ReDim Involved_R_T(0)
			Dim ret

			TrackInvolvedPeople = ""
			TrackComposers = ""
			TrackConductors = ""
			TrackProducers = ""
			TrackLyricists = ""
			TrackFeaturing = AlbumFeaturing

			If UBound(Involved_R) > 0 Then
				For tmp = 1 To UBound(Involved_R)
					ReDim Preserve Involved_R_T(tmp)
					Involved_R_T(tmp) = Involved_R(tmp)
				Next
			End If

			For tmp = 1 To UBound(TrackPos)
				If TrackPos(tmp) = position Then
					WriteLog "trackpos(" & tmp & ")=" & trackpos(tmp)
					REM msgbox tmp
					involvedRole = TrackRoles(tmp)
					involvedArtist = TrackArtist2(tmp)

					If LookForFeaturing(involvedRole) Then
						If InStr(TrackFeaturing, involvedArtist) = 0 Then
							If TrackFeaturing = "" Then
								If CheckFeaturingName Then
									TrackFeaturing = TxtFeaturingName & " " & involvedArtist
								Else
									TrackFeaturing = involvedRole & " " & involvedArtist
								End If
							Else
								TrackFeaturing = TrackFeaturing & Separator & involvedArtist
							End If
						End If
						WriteLog("TrackFeaturing=" & TrackFeaturing)
					Else
						Do
							ret = searchKeyword(LyricistKeywords, involvedRole, TrackLyricists, involvedArtist)
							If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
								TrackLyricists = ret
								WriteLog ("TrackLyricists=" & TrackLyricists)
								Exit Do
							ElseIf ret = "ALREADY_INSIDE_ROLE" Then
								Exit Do
							End If
							ret = searchKeyword(ConductorKeywords, involvedRole, TrackConductors, involvedArtist)
							If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
								TrackConductors = ret
								WriteLog ("TrackConductors=" & TrackConductors)
								Exit Do
							ElseIf ret = "ALREADY_INSIDE_ROLE" Then
								Exit Do
							End If
							ret = searchKeyword(ProducerKeywords, involvedRole, TrackProducers, involvedArtist)
							If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
								TrackProducers = ret
								WriteLog ("TrackProducers=" & TrackProducers)
								Exit Do
							ElseIf ret = "ALREADY_INSIDE_ROLE" Then
								Exit Do
							End If
							ret = searchKeyword(ComposerKeywords, involvedRole, TrackComposers, involvedArtist)
							If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
								TrackComposers = ret
								WriteLog ("TrackComposers=" & TrackComposers)
								Exit Do
							ElseIf ret = "ALREADY_INSIDE_ROLE" Then
								Exit Do
							End If
							tmp2 = search_involved(Involved_R_T, involvedRole)
							If tmp2 = -1 Then
								ReDim Preserve Involved_R_T(UBound(Involved_R_T)+1)
								Involved_R_T(UBound(Involved_R_T)) = involvedRole & ": " & TrackArtist2(tmp)
								WriteLog ("New Role: " & involvedRole & ": " & TrackArtist2(tmp))
							Else
								If InStr(Involved_R_T(tmp2), TrackArtist2(tmp)) = 0 Then
									Involved_R_T(tmp2) = Involved_R_T(tmp2) & ", " & TrackArtist2(tmp)
									WriteLog ("Role updated: " & Involved_R_T(tmp2))
								Else
									WriteLog ("artist already inside role")
								End If
							End If
							Exit Do
						Loop While True
					End If
				End If
			Next

			Dim trackArtist, artistList, FoundFeaturing, tmpJoin, tmpTrackArtist
			artistList = ""
			tmpJoin = ""

			WriteLog " "
			WriteLog("Search for TrackArtist")
			If currentTrack.Exists("artists") Then
				FoundFeaturing = False
				For Each artist in currentTrack("artists")
					WriteLog " "
					Set currentArtist = currentTrack("artists")(artist)
					If (currentArtist("anv") <> "") And Not CheckUseAnv Then
						tmpTrackArtist = CleanArtistName(currentArtist("anv"))
					Else
						tmpTrackArtist = CleanArtistName(currentArtist("name"))
					End If
					If FoundFeaturing = False Then
						artistList = artistList & tmpTrackArtist
					Else
						If TrackFeaturing = "" Then
							If CheckFeaturingName Then
								TrackFeaturing = TxtFeaturingName & " " & tmpTrackArtist
							Else
								TrackFeaturing = tmpJoin & " " & tmpTrackArtist
							End If
						Else
							TrackFeaturing = TrackFeaturing & ", " & tmpTrackArtist
						End If
						WriteLog("TrackFeaturing=" & TrackFeaturing)
					End If
					'TitleFeaturing
					If currentArtist("join") <> "" Then
						If LookForFeaturing(currentArtist("join")) Then
							FoundFeaturing = True
							tmpJoin = currentArtist("join")
						Else
							artistList = artistList & " " & currentArtist("join") & " "
							FoundFeaturing = False
						End If
					End If
					WriteLog("artistlist=" & artistlist)
				Next
			End If

			If artistList = "" Then artistList = AlbumArtistTitle

			If Right(artistList, 3) = " , " Then artistList = Left(artistList, Len(artistList)-3)

			If currentTrack.Exists("extraartists") Then
				WriteLog " "
				For Each extra In currentTrack("extraartists")
					Set currentArtist = CurrentTrack("extraartists")(extra)
					If (currentArtist("anv") <> "") And Not CheckUseAnv Then
						involvedArtist = CleanArtistName(currentArtist("anv"))
					Else
						involvedArtist = CleanArtistName(currentArtist("name"))
					End If
					If involvedArtist <> "" Then
						role = currentArtist("role")
						NoSplit = False
						If InStr(role, ",") = 0 Then
							involvedRole = role
							zahl = 1
							NoSplit = True
						Else
							REM rolea = Split(role, ", ")
							rolea = CheckSpecialRole(role)
							zahl = UBound(rolea)
						End If
						For zahltemp = 1 To zahl
							If NoSplit = False Then
								involvedRole = rolea(zahltemp)
							End If

							If LookForFeaturing(involvedRole) Then
								If InStr(artistList, involvedArtist) = 0 Then
									If TrackFeaturing = "" Then
										If CheckFeaturingName Then
											TrackFeaturing = TxtFeaturingName & " " & involvedArtist
										Else
											TrackFeaturing = involvedRole & " " & involvedArtist
										End If
									Else
										If InStr(TrackFeaturing, involvedArtist) = 0 Then
											TrackFeaturing = TrackFeaturing & ", " & involvedArtist
										End If
									End If
								End If
							Else
								Do
									tmp = searchKeyword(LyricistKeywords, involvedRole, TrackLyricists, involvedArtist)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										TrackLyricists = tmp
										WriteLog ("TrackLyricists=" & TrackLyricists)
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										Exit Do
									End If
									tmp = searchKeyword(ConductorKeywords, involvedRole, TrackConductors, involvedArtist)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										TrackConductors = tmp
										WriteLog ("TrackConductors=" & TrackConductors)
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										Exit Do
									End If
									tmp = searchKeyword(ProducerKeywords, involvedRole, TrackProducers, involvedArtist)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										TrackProducers = tmp
										WriteLog ("TrackProducers=" & TrackProducers)
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										Exit Do
									End If
									tmp = searchKeyword(ComposerKeywords, involvedRole, TrackComposers, involvedArtist)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										TrackComposers = tmp
										WriteLog ("TrackComposers=" & TrackComposers)
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										Exit Do
									End If
									tmp2 = search_involved(Involved_R_T, involvedRole)
									If tmp2 = -1 Then
										ReDim Preserve Involved_R_T(UBound(Involved_R_T)+1)
										Involved_R_T(UBound(Involved_R_T)) = involvedRole & ": " & involvedArtist
										WriteLog ("New Role: " & involvedRole & ": " & involvedArtist)
									Else
										If InStr(Involved_R_T(tmp2), involvedArtist) = 0 Then
											Involved_R_T(tmp2) = Involved_R_T(tmp2) & ", " & involvedArtist
											WriteLog ("Role updated: " & Involved_R_T(tmp2))
										Else
											WriteLog ("artist already inside role")
										End If
									End If
									Exit Do
								Loop While True
							End If
						Next
					End If
				Next
				WriteLog("TrackArtist end")
			End If

			If TrackFeaturing <> "" Then
				If CheckTitleFeaturing = True Then
					tmp = InStrRev(TrackFeaturing, ", ")
					If tmp = 0 Or ArtistLastSeparator = False Then
						trackName = trackName & " (" & TrackFeaturing & ")"
					Else
						trackName = trackName & " (" &  Left(TrackFeaturing, tmp-1) & " & " & Mid(TrackFeaturing, tmp+2) & ")"
					End If
				Else
					tmp = InStrRev(TrackFeaturing, ", ")
					If tmp = 0 Or ArtistLastSeparator = False Then
						artistList = artistList & " " & TrackFeaturing
					Else
						artistList = artistList & " " & Left(TrackFeaturing, tmp-1) & " & " & Mid(TrackFeaturing, tmp+2)
					End If
				End If
			End If

			If InStr(artistList, " & ") <> 0 And ArtistLastSeparator = False Then
				artistList = Replace(artistList, " & ", ArtistSeparator)
			End If
			If ArtistSeparator <> ", " Then
				artistList = Replace(artistList, ", ", ArtistSeparator)
				artistList = Replace(artistList, " " & ArtistSeparator, ArtistSeparator)
			Else
				artistList = Replace(artistList, " , ", ", ")
			End If
			ArtistTitles.Add artistList

			TrackLyricists = FindArtist(TrackLyricists, AlbumLyricist)
			If AlbumLyricist <> "" and TrackLyricists <> "" Then
				Lyricists.Add AlbumLyricist & "; " & TrackLyricists
			Else
				Lyricists.Add AlbumLyricist & TrackLyricists
			End If
			TrackComposers = FindArtist(TrackComposers, AlbumComposer)
			If AlbumComposer <> "" and TrackComposers <> "" Then
				Composers.Add AlbumComposer & "; " & TrackComposers
			Else
				Composers.Add AlbumComposer & TrackComposers
			End If
			TrackConductors = FindArtist(TrackConductors, AlbumConductor)
			If AlbumConductor <> "" and TrackConductors <> "" Then
				Conductors.Add AlbumConductor & "; " & TrackConductors
			Else
				Conductors.Add AlbumConductor & TrackConductors
			End If
			TrackProducers = FindArtist(TrackProducers, AlbumProducer)
			If AlbumProducer <> "" and TrackProducers <> "" Then
				Producers.Add AlbumProducer & "; " & TrackProducers
			Else
				Producers.Add AlbumProducer & TrackProducers
			End If

			If UBound(Involved_R_T) > 0 Then
				For tmp = 1 To UBound(involved_R_T)
					TrackInvolvedPeople = TrackInvolvedPeople & Involved_R_T(tmp) & "; "
				Next
				TrackInvolvedPeople = Left(TrackInvolvedPeople, Len(TrackInvolvedPeople)-2)
			Else
				TrackInvolvedPeople = ""
			End If

			InvolvedArtists.Add TrackInvolvedPeople
			Tracks.Add trackName
			iTrackNum = iTrackNum + 1
		Next

		If cSubTrack <> -1 Then
			If SubTrackNameSelection = False Then
				Tracks.Item(cSubTrack) = Tracks.Item(cSubTrack) & " (" & subTrackTitle & ")"
			Else
				Tracks.Item(cSubTrack) = subTrackTitle
			End If
			cSubTrack = -1
			subTrackTitle = ""
			CharSeparatorSubTrack = 0
		End If

		' Get album title
		AlbumTitle = currentRelease("title")

		' Get Album art URL
		If CurrentRelease.Exists("images") Then
			For Each i In CurrentRelease("images")
				Set currentImage = CurrentRelease("images")(i)

				If currentImage("type") = "primary" Or AlbumArtURL = "" Then
					AlbumArtURL = currentImage("uri")
					AlbumArtURL = Replace(AlbumArtURL, "http://api.discogs.com", "http://s.pixogs.com")
					AlbumArtThumbNail = currentImage("uri150")
					AlbumArtThumbNail = Replace(AlbumArtThumbnail, "http://api.discogs.com", "http://s.pixogs.com")
				End If
			Next
		End If

		'----------------------------------DiscogsImages----------------------------------------
		Set ImageList = SDB.NewStringList
		Set SaveImageType = SDB.NewStringList
		Set SaveImage = SDB.NewStringList
		ImagesCount = 0
		Dim tmpArt

		If CurrentRelease.Exists("images") Then
			ImagesCount = CurrentRelease("images").Count
			If CurrentRelease("images").Count > 1 Then
				For Each i In CurrentRelease("images")
					Set currentImage = CurrentRelease("images")(i)
					tmpArt = currentImage("uri")
					tmpArt = Replace(tmpArt, "http://api.discogs.com", "http://s.pixogs.com")
					If AlbumArtURL <> tmpArt Then
						ImageList.add tmpArt
						SaveImageType.add "other"
						SaveImage.add "0"
					End If	
				Next
			End If
		End If
		'----------------------------------DiscogsImages----------------------------------------

		' Get Master ID
		If CurrentRelease.Exists("master_id") Then
			theMaster = currentRelease("master_id")
			If SavedMasterId <> theMaster Then
				OriginalDate = ReloadMaster(theMaster)
				SavedMasterId = theMaster
			End If
		Else
			theMaster = ""
			SavedMasterId = theMaster
			OriginalDate = ""
		End If


		' Get release year/date
		If CurrentRelease.Exists("released") Then
			ReleaseDate = CurrentRelease("released")
			If Len(ReleaseDate) > 4 Then
				ReleaseSplit = Split(ReleaseDate,"-")
				If ReleaseSplit(2) = "00" Then
					ReleaseDate = Left(ReleaseDate, 4)
				Else
					ReleaseDate = ReleaseSplit(2) & "-" & ReleaseSplit(1) & "-" & ReleaseSplit(0)
				End If
				If CheckYearOnlyDate Then
					ReleaseDate = Right(ReleaseDate, 4)
				End If
			End If
		Else
			ReleaseDate = ""
		End If

		'Set OriginalDate
		If OriginalDate <> "" Then
			If Len(OriginalDate) > 4 Then
				ReleaseSplit = Split(OriginalDate,"-")
				If ReleaseSplit(2) = "00" Then
					OriginalDate = Left(OriginalDate, 4)
				Else
					OriginalDate = ReleaseSplit(2) & "-" & ReleaseSplit(1) & "-" & ReleaseSplit(0)
				End If
				If CheckYearOnlyDate Then
					OriginalDate = Right(OriginalDate, 4)
				End If
			End If
		End If

		' Get genres
		For Each g In CurrentRelease("genres")
			AddToField Genres, CurrentRelease("genres")(g)
		Next

		' Get styles/moods/themes
		If CurrentRelease.Exists("styles") Then
			For Each s In CurrentRelease("styles")
				AddToField Styles, CurrentRelease("styles")(s)
			Next
		End If

		' Get Label
		If CurrentRelease.Exists("labels") Then
			For Each l in CurrentRelease("labels")
				Set currentLabel = CurrentRelease("labels")(l)
				If SavedLabelId = "" Then
					If currentLabel.Exists("id") Then
						SavedLabelId = currentLabel("id")
					End If
				End If
				AddToField theLabels, CleanArtistName(currentLabel("name"))
				AddToField theCatalogs, currentLabel("catno")
			Next
		Else
			theLabels = ""
			theCatalogs = ""
		End If

		' Get Country
		If CurrentRelease.Exists("country") Then
			theCountry = CurrentRelease("country")
		Else
			theCountry = ""
		End If

		' Get Format
		If CurrentRelease.Exists("formats") Then
			For Each f in CurrentRelease("formats")
				Set currentFormat = CurrentRelease("formats")(f)
				AddToField theFormat, currentFormat("name")
				If currentFormat.Exists("descriptions") Then
					For Each d in currentFormat("descriptions")
						theFormat = theFormat & ", " & currentFormat("descriptions")(d)
					Next
				End If
			Next
		Else
			theFormat = ""
		End If

		' Get Comment
		If CurrentRelease.Exists("notes") Then
			Comment = CurrentRelease("notes")
		Else
			Comment = ""
		End If

		' Get data_quality
		If CurrentRelease.Exists("data_quality") Then
			DataQuality = CurrentRelease("data_quality")
		Else
			DataQuality = ""
		End If
	End If

	FormatSearchResultsViewer Tracks, TracksNum, TracksCD, Durations, AlbumArtist, AlbumArtistTitle, ArtistTitles, AlbumTitle, ReleaseDate, OriginalDate, Genres, Styles, theLabels, theCountry, AlbumArtThumbNail, CurrentResultID, theCatalogs, Lyricists, Composers, Conductors, Producers, InvolvedArtists, theFormat, theMaster, comment, DiscogsTracksNum, DataQuality

	Dim SelectedTracks, j
	Set SelectedTracks = SDB.NewStringList
	Set SelectedSongsGlobal = SDB.NewSongList
	For i = 0 To Tracks.Count - 1
		If UnselectedTracks(i) = "" Then
			SelectedTracks.Add Tracks.Item(i)
		End If
	Next
	For i = 0 To SDB.Tools.WebSearch.NewTracks.Count -1
		If UnselectedTracks(i) = "" Then
			SelectedSongsGlobal.Add SDB.Tools.WebSearch.NewTracks.item(i)
		End If
	Next

	SDB.Tools.WebSearch.SmartUpdateTracks SelectedTracks

	If CheckCover Then
		SDB.Tools.WebSearch.AlbumArtURL = AlbumArtURL
	ElseIf SmallCover Then
		SDB.Tools.WebSearch.AlbumArtURL = AlbumArtThumbNail
	Else
		SDB.Tools.WebSearch.AlbumArtURL = ""
	End If


	For i = 0 To SDB.Tools.WebSearch.NewTracks.Count - 1

		If CheckArtist Then SDB.Tools.WebSearch.NewTracks.Item(i).ArtistName = AlbumArtistTitle

		For j = 0 To Tracks.Count - 1
			If Tracks.Item(j) = SDB.Tools.WebSearch.NewTracks.Item(i).Title Then
				If UnselectedTracks(j) = "" Then
					If CheckArtist Then SDB.Tools.WebSearch.NewTracks.Item(i).ArtistName = ArtistTitles.Item(j)
					If CheckTrackNum Then SDB.Tools.WebSearch.NewTracks.Item(i).TrackOrderStr = TracksNum.Item(j)
					If CheckDiscNum Then SDB.Tools.WebSearch.NewTracks.Item(i).DiscNumberStr = TracksCD.Item(j)
					If CheckInvolved Then SDB.Tools.WebSearch.NewTracks.Item(i).InvolvedPeople = InvolvedArtists.Item(j)
					If CheckLyricist Then SDB.Tools.WebSearch.NewTracks.Item(i).Lyricist = Lyricists.Item(j)
					If CheckComposer Then SDB.Tools.WebSearch.NewTracks.Item(i).Author = Composers.Item(j)
					If CheckConductor Then SDB.Tools.WebSearch.NewTracks.Item(i).Conductor = Conductors.Item(j)
					If CheckProducer Then SDB.Tools.WebSearch.NewTracks.Item(i).Producer = Producers.Item(j)
				End If
			End If
		Next
		If CheckAlbumArtist Then SDB.Tools.WebSearch.NewTracks.Item(i).AlbumArtistName = AlbumArtist
		If CheckAlbum Then SDB.Tools.WebSearch.NewTracks.Item(i).AlbumName = AlbumTitle

		If CheckDate Then
			If Len(ReleaseDate) > 4 Then
				SDB.Tools.WebSearch.NewTracks.Item(i).Year = Mid(ReleaseDate,7,4)
				SDB.Tools.WebSearch.NewTracks.Item(i).Month = Mid(ReleaseDate,4,2)
				SDB.Tools.WebSearch.NewTracks.Item(i).Day = Mid(ReleaseDate,1,2)
			ElseIf IsNumeric(ReleaseDate) Then
				SDB.Tools.WebSearch.NewTracks.Item(i).Year = ReleaseDate
			ElseIf ReleaseDate = "" Then
				SDB.Tools.WebSearch.NewTracks.Item(i).Year = -1
			End If
		End If
		If CheckOrigDate Then
			If Len(OriginalDate) > 4 Then
				SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = Mid(OriginalDate,7,4)
				SDB.Tools.WebSearch.NewTracks.Item(i).OriginalMonth = Mid(OriginalDate,4,2)
				SDB.Tools.WebSearch.NewTracks.Item(i).OriginalDay = Mid(OriginalDate,1,2)
			ElseIf IsNumeric(OriginalDate) Then
				SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = OriginalDate
			ElseIf OriginalDate = "" Then
				SDB.Tools.WebSearch.NewTracks.Item(i).OriginalYear = -1
			End If
		End If

		If CheckStyleField = "Default (stored with Genre)" Then
			If CheckGenre And CheckStyle Then
				SDB.Tools.WebSearch.NewTracks.Item(i).Genre = Genres & Separator & Styles
				If Genres = "" Then SDB.Tools.WebSearch.NewTracks.Item(i).Genre = Styles
				If Styles = "" Then SDB.Tools.WebSearch.NewTracks.Item(i).Genre = Genres
			ElseIf CheckGenre Then
				SDB.Tools.WebSearch.NewTracks.Item(i).Genre = Genres
			ElseIf CheckStyle Then
				SDB.Tools.WebSearch.NewTracks.Item(i).Genre = Styles
			End If
		Else
			If CheckGenre Then
				SDB.Tools.WebSearch.NewTracks.Item(i).Genre = Genres
			End If
			If CheckStyle Then
				If CheckStyleField = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = Styles
				If CheckStyleField = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = Styles
				If CheckStyleField = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = Styles
				If CheckStyleField = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = Styles
				If CheckStyleField = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = Styles
			End If
		End If
		If CheckLabel Then SDB.Tools.WebSearch.NewTracks.Item(i).Publisher = theLabels

		If CheckComment Then SDB.Tools.WebSearch.NewTracks.Item(i).Comment = comment

		If CheckRelease Then
			If ReleaseTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = CurrentResultID
			If ReleaseTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = CurrentResultID
			If ReleaseTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = CurrentResultID
			If ReleaseTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = CurrentResultID
			If ReleaseTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = CurrentResultID
			If ReleaseTag = "Grouping" Then SDB.Tools.WebSearch.NewTracks.Item(i).Grouping = CurrentResultID
			If ReleaseTag = "ISRC" Then SDB.Tools.WebSearch.NewTracks.Item(i).ISRC = CurrentResultID
			If ReleaseTag = "Encoding" Then SDB.Tools.WebSearch.NewTracks.Item(i).Encodiung = CurrentResultID
			If ReleaseTag = "Copyright" Then SDB.Tools.WebSearch.NewTracks.Item(i).Copyright = CurrentResultID
		End If

		If CheckCatalog Then
			If CatalogTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theCatalogs
			If CatalogTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theCatalogs
			If CatalogTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theCatalogs
			If CatalogTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theCatalogs
			If CatalogTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theCatalogs
		End If

		If CheckCountry Then
			If CountryTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theCountry
			If CountryTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theCountry
			If CountryTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theCountry
			If CountryTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theCountry
			If CountryTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theCountry
		End If

		If CheckFormat Then
			If FormatTag = "Custom1" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom1 = theFormat
			If FormatTag = "Custom2" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom2 = theFormat
			If FormatTag = "Custom3" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom3 = theFormat
			If FormatTag = "Custom4" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom4 = theFormat
			If FormatTag = "Custom5" Then SDB.Tools.WebSearch.NewTracks.Item(i).Custom5 = theFormat
		End If
	Next
	SDB.Tools.WebSearch.RefreshViews   ' Tell MM that we have made some changes
End Sub

Function CheckSpecialRole(Role)

	Dim tmp, tmp2, tmp3
	WriteLog "Start CheckSpecialRole"
	ReDim SingleRole(0)
	Do While 1 = 1
		tmp = InStr(Role, ",")
		tmp2 = InStr(Role, "[")
		tmp3 = InStr(Role, "]")
		If tmp = 0 Then
			ReDim Preserve SingleRole(UBound(SingleRole)+1)
			SingleRole(UBound(SingleRole)) = Role
			Exit Do
		End If
		If tmp < tmp2 Then
			ReDim Preserve SingleRole(UBound(SingleRole)+1)
			SingleRole(UBound(SingleRole)) = Left(Role, tmp-1)
			Role = LTrim(Mid(Role, tmp+1))
		End If
		If tmp > tmp2 And tmp > tmp3 Then
			ReDim Preserve SingleRole(UBound(SingleRole)+1)
			SingleRole(UBound(SingleRole)) = Left(Role, tmp-1)
			Role = LTrim(Mid(Role, tmp+1))
		End If
		If tmp > tmp2 And tmp < tmp3 Then
			tmp = InStr(tmp3, Role, ",", 0)
			ReDim Preserve SingleRole(UBound(SingleRole)+1)
			If tmp <> 0 Then
				SingleRole(UBound(SingleRole)) = Left(Role, tmp-1)
				Role = LTrim(Mid(Role, tmp+1))
			Else
				SingleRole(UBound(SingleRole)) = Role
				Exit Do
			End If
		End If
	Loop
	WriteLog "End CheckSpecialRole"
	CheckSpecialRole = SingleRole

End Function


Function FindSubTrackSplit(position)

	Dim tmp
	For tmp = 1 To Len(position)-1
		If Not IsNumeric(Mid(position, tmp, 1)) Then
			FindSubTrackSplit = Left(position, tmp-1)
			Exit For
		End If
	Next

End Function


Function FindArtist(ArtistList1, ArtistList2)

	Dim tmpArtist, i
	ReDim newArtistList1(0)
	tmpArtist = Split(ArtistList1, "; ")
	For i = 0 To UBound(tmpArtist)
		If InStr(ArtistList2,tmpArtist(i)) = 0 Then
			ReDim Preserve newArtistList1(UBound(newArtistList1)+1)
			newArtistList1(UBound(newArtistList1)) = tmpArtist(i)
		End If
	Next
	For i = 0 To UBound(newArtistList1)
		If FindArtist = "" Then
			FindArtist = newArtistList1(i)
		Else
			FindArtist = FindArtist & "; " & newArtistList1(i)
		End If
	Next

End Function


Sub Track_from_to (currentTrack, currentArtist, involvedRole, Title_Position, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition)

	Dim tmp3, tmp4, tmpSide1, tmpSide2, tmpSideD, Vinyl_Pos1, Vinyl_Pos2, zahltemp3, ret
	WriteLog "currentTrack=" & currentTrack
	tmp3 = Split(currentTrack, " ")
	tmpSide1 = ""
	tmpSide2 = ""
	tmpSideD = ""

	tmp3(0) = exchange_roman_numbers(tmp3(0))
	tmp3(2) = exchange_roman_numbers(tmp3(2))

	If InStr(tmp3(0), "-") <> 0 Then
		tmp4 = Split(tmp3(0), "-")
		tmpSide1 = Trim(tmp4(0))
		tmp3(0) = Trim(tmp4(1))
		tmp3(0) = exchange_roman_numbers(tmp3(0))
		tmpSideD = "-"
	End If
	If InStr(tmp3(2), "-") <> 0 Then
		tmp4 = Split(tmp3(2), "-")
		tmpSide2 = Trim(tmp4(0))
		tmp3(2) = Trim(tmp4(1))
		tmp3(2) = exchange_roman_numbers(tmp3(2))
		tmpSideD = "-"
	End If
	If InStr(tmp3(0), ".") <> 0 Then
		tmp4 = Split(tmp3(0), ".")
		tmpSide1 = Trim(tmp4(0))
		tmp3(0) = Trim(tmp4(1))
		tmp3(0) = exchange_roman_numbers(tmp3(0))
		tmpSideD = "."
	End If
	If InStr(tmp3(2), ".") <> 0 Then
		tmp4 = Split(tmp3(2), ".")
		tmpSide2 = Trim(tmp4(0))
		tmp3(2) = Trim(tmp4(1))
		tmp3(2) = exchange_roman_numbers(tmp3(2))
		tmpSideD = "."
	End If
	If Left(tmp3(0), 2) = "CD" Then
		tmpSide1 = "CD"
		tmp3(0) = Mid(tmp3(0), 3)
	End If
	If Left(tmp3(2), 2) = "CD" Then
		tmpSide2 = "CD"
		tmp3(2) = Mid(tmp3(2), 3)
	End If
	If Left(tmp3(0), 3) = "DVD" Then
		tmpSide1 = "DVD"
		tmp3(0) = Mid(tmp3(0), 4)
	End If
	If Left(tmp3(2), 3) = "DVD" Then
		tmpSide2 = "DVD"
		tmp3(2) = Mid(tmp3(2), 4)
	End If
	If IsNumeric(Right(tmp3(0),1)) = False And Len(tmp3(0)) > 1 Then
		tmp3(0) = Left(tmp3(0), Len(tmp3(0))-1)
	End If
	If IsNumeric(Right(tmp3(2),1)) = False And Len(tmp3(2)) > 1 Then
		tmp3(2) = Left(tmp3(2), Len(tmp3(2))-1)
	End If
	If IsNumeric(tmp3(0)) = False Then
		If Len(tmp3(0)) > 1 Then
			tmpSide1 = Left(tmp3(0), 1)
			tmp3(0) = Mid(tmp3(0), 2)
		Else
			tmpSide1 = tmp3(0)
			tmp3(0) = 1
		End If
	End If
	If IsNumeric(tmp3(2)) = False Then
		If Len(tmp3(2)) > 1 Then
			tmpSide2 = Left(tmp3(2), 1)
			tmp3(2) = Mid(tmp3(2), 2)
		Else
			tmpSide2 = tmp3(2)
			tmp3(2) = 1
		End If
	End If

	If tmpSide1 <> tmpSide2 Then
		Vinyl_Pos1 = tmpSide1
		Vinyl_Pos2 = tmp3(0)
		Do
			If LeadingZeroTrackPosition = True And Vinyl_Pos2 < 10 Then
				Vinyl_Pos2 = "0" & Vinyl_Pos2
			End If
			tmp4 = Vinyl_Pos1 & tmpSideD & Vinyl_Pos2
			ret = search_involved(Title_Position, tmp4)
			If ret = -1 Then
				If IsNumeric(Vinyl_Pos1) = True Then
					If Vinyl_Pos1 > 101 Then Exit Do
					Vinyl_Pos1 = Vinyl_Pos1 + 1
				Else
					If Chr(Asc(Vinyl_Pos1)) = "Z" Then Exit Do
					Vinyl_Pos1 = Chr(Asc(Vinyl_Pos1) + 1)
				End If
				Vinyl_Pos2 = "1"
			Else
				ReDim Preserve TrackRoles(UBound(TrackRoles)+1)
				ReDim Preserve TrackArtist2(UBound(TrackArtist2)+1)
				ReDim Preserve TrackPos(UBound(TrackPos)+1)
				TrackArtist2(UBound(TrackArtist2)) = currentArtist
				TrackRoles(UBound(TrackRoles)) = involvedRole
				TrackPos(UBound(TrackPos)) = Vinyl_Pos1 & tmpSideD & Vinyl_Pos2
				WriteLog "  currentTrack=" & Vinyl_Pos1 & tmpSideD & Vinyl_Pos2
				If cStr(Vinyl_Pos1) = cStr(tmpSide2) And cStr(Vinyl_Pos2) = cStr(tmp3(2)) Then Exit Do
				Vinyl_Pos2 = Vinyl_Pos2 + 1
			End If
		Loop While True
	Else
		For zahltemp3 = tmp3(0) To tmp3(2)
			If LeadingZeroTrackPosition = True And zahltemp3 < 10 Then
				zahltemp3 = "0" & zahltemp3
			End If
			ReDim Preserve TrackRoles(UBound(TrackRoles)+1)
			ReDim Preserve TrackArtist2(UBound(TrackArtist2)+1)
			ReDim Preserve TrackPos(UBound(TrackPos)+1)
			TrackArtist2(UBound(TrackArtist2)) = currentArtist
			TrackRoles(UBound(TrackRoles)) = involvedRole
			TrackPos(UBound(TrackPos)) = tmpSide1 & tmpSideD & zahltemp3
			WriteLog "  currentTrack=" & tmpSide1 & tmpSideD & zahltemp3
		Next
	End If

End Sub


Sub Add_Track_Role(currentTrack, currentArtist, involvedRole, TrackRoles, TrackArtist2, TrackPos)

	WriteLog "currentTrack=" & currentTrack
	currentTrack = exchange_roman_numbers(currentTrack)
	ReDim Preserve TrackRoles(UBound(TrackRoles)+1)
	ReDim Preserve TrackArtist2(UBound(TrackArtist2)+1)
	ReDim Preserve TrackPos(UBound(TrackPos)+1)
	TrackArtist2(UBound(TrackArtist2)) = currentArtist
	TrackRoles(UBound(TrackRoles)) = involvedRole
	TrackPos(UBound(TrackPos)) = currentTrack

End Sub


' ShowResult is called every time the search result is changed from the drop
' down at the top of the window
Sub ShowResult(ResultID)

	Dim ReleaseID, searchURL, oXMLHTTP, ResponseHTML, TXTBegin, TXTEnd, Title, SelectedTracks, searchURL_F

	WebBrowser.SetHTMLDocument ""                 ' Deletes visible search result

	If UseMetalArchives = True Then
		searchURL = ResultsReleaseID.Item(ResultID)
		Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
		Call oXMLHTTP.open("GET", searchURL, False)
		Call oXMLHTTP.send()
		If oXMLHTTP.Status = 200 Then
			ResponseHTML = oXMLHTTP.responseText
			TXTBegin = InStr(ResponseHTML, "tbody")
			TXTEnd = InStr(ResponseHTML, "/tbody")
			ResponseHTML = Mid(ResponseHTML, TXTBegin, TXTEnd - TXTBegin - 1)
			Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
			Dim f
			Set f = fso.OpenTextFile(SDB.ScriptsPath&"test2.log", 2, true, -1)
			f.WriteLine ResponseHTML
			f.Close
			
			Set SelectedTracks = SDB.NewStringList

			SDB.Tools.WebSearch.ClearTracksData

			Do While 1 = 1
				TXTBegin = InStr(ResponseHTML, "<td class=" & Chr(34) & "wrapWords")
				If TXTBegin = 0 Then Exit Do
				ResponseHTML = Mid(ResponseHTML, TXTBegin + 23)
				TXTEnd = InStr(ResponseHTML, "</td>")
				Title = Left(ResponseHTML, TXTEnd - 1)
				Title = Replace(Title, Chr(10), "")
				Title = Trim(Replace(Title, Chr(9), " "))
				WriteLog "Title=" & Title & chr(34)
				SelectedTracks.add Title
			Loop
			SDB.Tools.WebSearch.SmartUpdateTracks SelectedTracks

			SDB.Tools.WebSearch.RefreshViews
		End If










	Else

		WriteLog "Start ShowResult"
		ReleaseID = ResultsReleaseID.Item(ResultID)
		If Right(Results.Item(ResultID), 1) = "*" Then  'Master-Release
			searchURL = ReleaseID
			searchURL_F = "http://api.discogs.com/masters/"
		Else
			searchURL = ReleaseID
			searchURL_F = "http://api.discogs.com/releases/"
		End If

		Set oXMLHTTP = CreateObject("Msxml2.XMLHttp.6.0")   
		oXMLHTTP.Open "POST", "http://www.germanc64.de/mm/oauth/check_new.php", False
		oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		oXMLHTTP.setRequestHeader "User-Agent","MediaMonkeyDiscogsAutoTagBatch/2.0 +http://mediamonkey.com"
		oXMLHTTP.send ("at=" & AccessToken & "&ats=" & AccessTokenSecret & "&searchURL=" & searchURL & "&searchURL_F=" & searchURL_F & "&searchURL_L=")
		WriteLog "Post durchgeführt"
		
		' use json api with vbsjson class at start of file now
		REM Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")

		Dim json
		Set json = New VbsJson

		Dim response

		If oXMLHTTP.Status = 200 Then
			Set CurrentRelease = json.Decode(oXMLHTTP.responseText)

			CurrentResultID = ReleaseID

			ReloadResults
		End If
	End If

End Sub


' This does the final clean up, so that our script doesn't leave any unwanted traces
Sub FinishSearch(Panel)

	Dim ret, res, RndFileName, i, itm, path, j, k, ImageSelected
	If IsObject(ImageList) Then
		If ImageList.Count > 0 Then
			ImageSelected = False
			For i = 0 to ImageList.Count - 1
				If SaveImage.Item(i) = 1 Then ImageSelected = True
			Next
			If ImageSelected = True Then
				res = SDB.MessageBox("Save the selected image(s) ?", mtConfirmation, Array(mbYes, mbNo))
				If res = 6 Then
					For i = 0 to ImageList.Count - 1
						res = 0
						If SaveImage.Item(i) = 1 Then
							Set itm = SelectedSongsGlobal.item(0)
							path = Mid(itm.Path,1,InStrRev(itm.Path,"\")-1)
							If CoverStorage = 1 Or CoverStorage = 3 Then
								If SDB.Tools.FileSystem.FileExists(path & "\" & FileNameList.Item(i)) = True Then
									res = SDB.MessageBox("The file " & FileNameList.Item(i) & " already exist. Overwrite it ?", mtConfirmation, Array(mbYes, mbNo))
									If res = 6 Then
										SDB.Tools.FileSystem.DeleteFile(path & "\" & FileNameList.Item(i))
										ret = getimages(ImageList.Item(i), path & "\" & FileNameList.Item(i))
									End If
								Else
									ret = getimages(ImageList.Item(i), path & "\" & FileNameList.Item(i))
									If ret = "" Then DebugOut("ERROR:Image Download failed !")
								End If
							End If
							If CoverStorage = 0 Then
								Dim max, min
								max=100000
								min=10000
								Randomize
								RndFileName = Int((max-min+1)*Rnd+min) & ".jpg"
								ret = getimages(ImageList.Item(i), path & "\" & RndFileName)
							End If

							If res <> 7 Then 'don't overwrite file
								For j = 0 To SelectedSongsGlobal.Count - 1
									Set itm = SelectedSongsGlobal.item(j)
									Dim pics : Set pics = itm.AlbumArt
									If pics Is Nothing Then
										Exit Sub
									End If
									Dim img, ImageTagCount
									ImageTagCount = pics.Count

									Set img = pics.AddNew
									img.Description = ""

									If CoverStorage = 1 Or CoverStorage = 3 Then
										img.PicturePath = path & "\" & FileNameList.Item(i)
										img.ItemStorage = 1
									Else
										img.PicturePath = path & "\" & RndFileName
										img.ItemStorage = 0
									End If
									For k = 0 to ImageTypeList.Count - 1
										If SaveImageType.Item(i) = ImageTypeList.Item(k) Then
											If k = 0 Then k = -2
											If k > 14 Then k = k + 1
											img.ItemType = k + 2
											pics.UpdateDB
											Exit For
										End If
									Next
									Set pics = itm.AlbumArt
									If ImageTagCount + 1 = pics.Count Then
									Else
									End If
									If CoverStorage = 3 Then
										Set pics = itm.AlbumArt
										ImageTagCount = pics.Count
										Set img = pics.AddNew
										img.Description = ""
										img.PicturePath = path & "\" & FileNameList.Item(i)
										img.ItemStorage = 0
										For k = 0 to ImageTypeList.Count - 1
											If SaveImageType.Item(i) = ImageTypeList.Item(k) Then
												If k = 0 Then k = -2
												If k > 14 Then k = k + 1
												img.ItemType = k + 2
												pics.UpdateDB
												Exit For
											End If
										Next
										Set pics = itm.AlbumArt
										If ImageTagCount + 1 = pics.Count Then
										Else
										End If
									End If
								Next
							End If
							If CoverStorage = 0 Then
								SDB.Tools.FileSystem.DeleteFile(path & "\" & RndFileName)
							End If
						End If
					Next
				End If
			End If
		End If
	End If

	WebBrowser.Common.DestroyControl      ' Destroy the external control
	Set WebBrowser = Nothing              ' Release global variable
	SDB.Objects("WebBrowser") = Nothing

	Set ini = Nothing
	Set ResultsReleaseID = Nothing
	Script.UnregisterAllEvents

End Sub


Function GetHeader()

	Dim templateHTML, i
	templateHTML = "<HTML>"
	templateHTML = templateHTML &  "<HEAD>"
	templateHTML = templateHTML &  "<style type=""text/css"" media=""screen"">"
	templateHTML = templateHTML &  ".tabletext { font-family: Arial, Helvetica, sans-serif; font-size: 8pt;}"
	templateHTML = templateHTML &  "option.tabletext{background-color:#3E7CBB;}"

	templateHTML = templateHTML &  "</style>"
	templateHTML = templateHTML &  "</HEAD>"
	templateHTML = templateHTML &  "<body bgcolor=""#FFFFFF"">"
	templateHTML = templateHTML &  "<table border=0 width=100% cellspacing=0 cellpadding=1 class=tabletext>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td align=left><a href=""http://www.discogs.com"" target=""_blank""><img src=""http://s.discogss.com/images/discogs-white-2.png"" border=""0""/ alt=""Discogs Homepage""></a><b>" & VersionStr & "</b></td>"
	templateHTML = templateHTML &  "<td colspan=3 align=right valign=top>"

	templateHTML = templateHTML &  "<table border=0 cellspacing=0 cellpadding=2 class=tabletext>"
	templateHTML = templateHTML &  "<tr><td colspan=2></td><td><b>Filter Results: </b></td><td colspan=3> </td></tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><b>Load:</b></td>"
	templateHTML = templateHTML &  "<td><b>Quick Search:</b></td>"
	templateHTML = templateHTML &  "<td align=left><button type=button class=tabletext id=""showmediatypefilter"">Set Type Filter</button></td>"
	templateHTML = templateHTML &  "<td align=left><button type=button class=tabletext id=""showmediaformatfilter"">Set Format Filter</button></td>"
	templateHTML = templateHTML &  "<td align=left><button type=button class=tabletext id=""showcountryfilter"">Set Country Filter</button></td>"
	templateHTML = templateHTML &  "<td align=left><button type=button class=tabletext id=""showyearfilter"">Set Year Filter</button></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td>"
	templateHTML = templateHTML &  "<select id=""load"" class=tabletext title=""Search Result=Search with Artist and Album Title" & vbCrLf & "Master Release=Show all releases from the master"">"

	For i = 0 To LoadList.Count - 1
		If LoadList.Item(i) <> CurrentLoadType Then
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(LoadList.Item(i)) & """>" & LoadList.Item(i) & "</option>"
		Else
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(LoadList.Item(i)) & """ selected>" & LoadList.Item(i) & "</option>"
		End If
	Next
	templateHTML = templateHTML &  "</select>"
	templateHTML = templateHTML &  "</td>"
	'Alternative Searches Begin
	templateHTML = templateHTML &  "<td>"
	templateHTML = templateHTML &  "<select id=""alternative"" class=tabletext>"
	For i = 0 To AlternativeList.Count - 1
		If AlternativeList.Item(i) <> SavedSearchTerm Then
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(AlternativeList.Item(i)) & """>" & AlternativeList.Item(i) & "</option>"
		Else
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(AlternativeList.Item(i)) & """ selected>" & AlternativeList.Item(i) & "</option>"
		End If
	Next
	templateHTML = templateHTML &  "</select>"
	templateHTML = templateHTML &  "</td>"
	'Alternative Searches End
	'Filters Begin
	templateHTML = templateHTML &  "<td>"
	templateHTML = templateHTML &  "<select id=""filtermediatype"" class=tabletext>"

	If FilterMediaType = "None" Then
		templateHTML = templateHTML &  "<option value=""None"">No MediaType Filter</option>"
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use MediaType Filter"">Use MediaType Filter</option>"
	ElseIf FilterMediaType = "Use MediaType Filter" Then
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use MediaType Filter"">Use MediaType Filter</option>"
		templateHTML = templateHTML &  "<option value=""None"">No MediaType Filter</option>"
	End If
	If FilterMediaType <> "None" And FilterMediaType <> "Use MediaType Filter" Then
		templateHTML = templateHTML &  "<option value=""None"">No MediaType Filter</option>"
		templateHTML = templateHTML &  "<option value=""Use MediaType Filter"">Use MediaType Filter</option>"
	End If
	For i = 1 To MediaTypeList.Count - 1
		If FilterMediaType <> MediaTypeList.Item(i) Or FilterMediaType = "None" Or FilterMediaType = "Use MediaType Filter" Then
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(MediaTypeList.Item(i)) & """>" & MediaTypeList.Item(i) & "</option>"
		Else
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(MediaTypeList.Item(i)) & """ selected>" & MediaTypeList.Item(i) & "</option>"
		End If
	Next
	templateHTML = templateHTML &  "</select>"
	templateHTML = templateHTML &  "</td>"

	templateHTML = templateHTML &  "<td>"
	templateHTML = templateHTML &  "<select id=""filtermediaformat"" class=tabletext>"

	If FilterMediaFormat = "None" Then
		templateHTML = templateHTML &  "<option value=""None"">No MediaFormat Filter</option>"
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use MediaFormat Filter"">Use MediaFormat Filter</option>"
	ElseIf FilterMediaFormat = "Use MediaFormat Filter" Then
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use MediaFormat Filter"">Use MediaFormat Filter</option>"
		templateHTML = templateHTML &  "<option value=""None"">No MediaFormat Filter</option>"
	End If
	If FilterMediaFormat <> "None" And FilterMediaFormat <> "Use MediaFormat Filter" Then
		templateHTML = templateHTML &  "<option value=""None"">No MediaFormat Filter</option>"
		templateHTML = templateHTML &  "<option value=""Use MediaFormat Filter"">Use MediaFormat Filter</option>"
	End If
	For i = 1 To MediaFormatList.Count - 1
		If FilterMediaFormat <> MediaFormatList.Item(i) Or FilterMediaFormat = "None" Or FilterMediaFormat = "Use MediaFormat Filter" Then
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(MediaFormatList.Item(i)) & """>" & MediaFormatList.Item(i) & "</option>"
		Else
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(MediaFormatList.Item(i)) & """ selected>" & MediaFormatList.Item(i) & "</option>"
		End If
	Next
	templateHTML = templateHTML &  "</select>"
	templateHTML = templateHTML &  "</td>"

	templateHTML = templateHTML &  "<td>"
	templateHTML = templateHTML &  "<select id=""filtercountry"" class=tabletext>"

	If FilterCountry = "None" Then
		templateHTML = templateHTML &  "<option value=""None"">No Country Filter</option>"
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use Country Filter"">Use Country Filter</option>"
	ElseIf FilterCountry = "Use Country Filter" Then
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use Country Filter"">Use Country Filter</option>"
		templateHTML = templateHTML &  "<option value=""None"">No Country Filter</option>"
	End If
	If FilterCountry <> "None" And FilterCountry <> "Use Country Filter" Then
		templateHTML = templateHTML &  "<option value=""None"">No Country Filter</option>"
		templateHTML = templateHTML &  "<option value=""Use Country Filter"">Use Country Filter</option>"
	End If
	For i = 1 To CountryList.Count - 1
		If FilterCountry <> CountryList.Item(i) Or FilterCountry = "None" Or FilterCountry = "Use Country Filter" Then
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(CountryList.Item(i)) & """>" & CountryList.Item(i) & "</option>"
		Else
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(CountryList.Item(i)) & """ selected>" & CountryList.Item(i) & "</option>"
		End If
	Next

	templateHTML = templateHTML &  "</select>"
	templateHTML = templateHTML &  "</td>"

	templateHTML = templateHTML &  "<td>"
	templateHTML = templateHTML &  "<select id=""filteryear"" class=tabletext>"

	If FilterYear = "None" Then
		templateHTML = templateHTML &  "<option value=""None"">No Year Filter</option>"
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use Year Filter"">Use Year Filter</option>"
	ElseIf FilterYear = "Use Year Filter" Then
		templateHTML = templateHTML &  "<option style=""background-color:#F4113F;"" value=""Use Year Filter"">Use Year Filter</option>"
		templateHTML = templateHTML &  "<option value=""None"">No Year Filter</option>"
	End If
	If FilterYear <> "None" And FilterYear <> "Use Year Filter" Then
		templateHTML = templateHTML &  "<option value=""None"">No Year Filter</option>"
		templateHTML = templateHTML &  "<option value=""Use Year Filter"">Use Year Filter</option>"
	End If
	For i = 1 To YearList.Count - 1
		If FilterYear <> YearList.Item(i) Or FilterYear = "None" Or FilterYear = "Use Year Filter" Then
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(YearList.Item(i)) & """>" & YearList.Item(i) & "</option>"
		Else
			templateHTML = templateHTML &  "<option value=""" & EncodeHtmlChars(YearList.Item(i)) & """ selected>" & YearList.Item(i) & "</option>"
		End If
	Next

	templateHTML = templateHTML &  "</select>"
	templateHTML = templateHTML &  "</td>"
	'Filters End
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</td>"
	templateHTML = templateHTML &  "</tr>"

	GetHeader = templateHTML

End Function


Function GetFooter()

	Dim templateHTML
	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</body>"
	templateHTML = templateHTML &  "</HTML>"

	GetFooter = templateHTML

End Function



' We use this procedure to reformat results as soon as they are downloaded
Sub FormatSearchResultsViewer(Tracks, TracksNum, TracksCD, Durations, AlbumArtist, AlbumArtistTitle, ArtistTitles, AlbumTitle, ReleaseDate, OriginalDate, Genres, Styles, theLabels, theCountry, theArt, releaseID, Catalog, Lyricists, Composers, Conductors, Producers, InvolvedPeople, theFormat, theMaster, comment, DiscogsTracksNum, DataQuality)

	Dim templateHTML, checkBox, text, listBox, submitButton, tmp
	Dim SelectedTracksCount, UnSelectedTracksCount
	Dim SubTrackFlag
	Dim i, theTracks, currentCD, theGenres
	templateHTML = ""
	templateHTML = templateHTML &  GetHeader()

	' Titles Begin
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td align=left bgcolor=""#CCCCCC""><b>Album Art:</b></td>"
	templateHTML = templateHTML &  "<td align=left bgcolor=""#CCCCCC""><b>Release Information:</b></td>"
	templateHTML = templateHTML &  "<td align=left bgcolor=""#CCCCCC""><b>Tracklisting:</b></td>"
	templateHTML = templateHTML &  "</tr>"
	' Titles End
	templateHTML = templateHTML &  "<tr>"
	' Release Cover Begin
	templateHTML = templateHTML &  "<td align=left valign=top>"
	templateHTML = templateHTML &  "<table border=0 cellspacing=0 cellpadding=1 class=tabletext>"
	If theArt <> "" Then
		templateHTML = templateHTML &  "<tr><td colspan=2><a href=""http://www.discogs.com/viewimages?release=<!RELEASEID!>"" target=""_blank""><img src=""<!COVER!>"" border=""0""/></a></td></tr>"
	Else
		templateHTML = templateHTML &  "<tr><td colspan=2><table width=150 height=150 border=1><tr><td><center>No Image<br>Available</center></td></tr></table></td></tr>"
	End If
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""cover"" >Large <input type=checkbox id=""smallcover"" >Small (150px)</td></tr>"
	If ImagesCount > 1 Then
		templateHTML = templateHTML &  "<tr><td colspan=2 align=center><button type=button class=tabletext id=""moreimages"">More Images</button></td></tr>"
	End If
	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><br></td></tr>"
	' Release Cover End

	' Options Begin
	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><button type=button class=tabletext id=""saveoptions"">Save Options</button></td></tr>"
	templateHTML = templateHTML &  "<tr><td align=center colspan=2><b>Options:</b></td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""lyricist"" >Save Lyricist</td></tr>"
	Rem " & SDB.Localize("Save") & " Lyricist</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""composer"" >Save Composer</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""conductor"" >Save Conductor</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""producer"" >Save Producer</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""involved"" >Save Involved People</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""comments"" >Save Comment</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""useanv"" title=""Artist Name Variation - Using no name variation (e.g. nickname)"" >Don't Use ANV's</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""yearonlydate"" title=""If checked only the Year will be saved (e.g. 14.01.1982 -> 1982)"" >Only Year Of Date</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""titlefeaturing"" title=""If checked the feat. Artist appears in the title tag (e.g. Aaliyah (ft. Timbaland) - We Need a Resolution  ->  Aaliyah - We Need a Resolution (ft. Timbaland) )"" >feat. Artist behind Title</td></tr>"

	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""FeaturingName"" title=""Rename 'feat.' to the given word"" >Rename 'feat.' to:</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=text id=""TxtFeaturingName"" ></td></tr>"

	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""various"" title=""Rename 'Various' Artist to the given word"" >Rename 'Various' Artist to:</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=text id=""txtvarious"" ></td></tr>"

	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""involvedpeoplesingle"" title=""Print every involved people in a single line"" >Every invol. people single line</td></tr>"

	templateHTML = templateHTML &  "<tr><td align=center colspan=2><br></td></tr>"
	templateHTML = templateHTML &  "<tr><td align=center colspan=2><b>Disc/Track Numbering:</b></td></tr>"

	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""TurnOffSubTrack"" title=""If checked the Sub-Track detection is turned off"" >No Sub-Track detection</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""SubTrackNameSelection"" title=""If checked the Sub-Track will be named like 'Sub-Track 1, Sub-Track 2, Sub Track 3'  if not checked the Sub-Tracks will be named like 'Track Name (Sub-Track 1, Sub-Track 2, Sub Track 3)'"" >Other Sub-Track Naming</td></tr>"

	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""forcenumeric"" title=""Always use numbers instead of letters (Vinyl-releases use A1, A2,..., B1, B2 as track numbering)"" >Force To Numeric</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""sidestodisc"" title=""Save the Vinyl sides to the disc tag"" >Sides To Disc</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""forcedisc"" title=""Always add a disc-number"" >Force Disc Usage</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""nodisc"" title=""Prevent the script from interpret sub tracks as disc-numbers"" >Force NO Disc Usage</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""leadingzero"" title=""Track Position: 1 -> 01   2 -> 02 ..."" >Add Leading Zero</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><br></td></tr>"

	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</td>"
	' Options End
	
	' Release Information Begin
	templateHTML = templateHTML &  "<td align=left valign=top>"
	templateHTML = templateHTML &  "<table border=0 cellspacing=0 cellpadding=1 class=tabletext>"
	
	iMaxTracks = Tracks.Count
	If TracksCD.Count < iMaxTracks Then
		iMaxTracks = TracksCD.Count
	End If
	
	'Check for different Track number
	SelectedTracksCount = 0
	UnSelectedTracksCount = 0
	SubTrackFlag = False
	For i = 0 To iMaxTracks - 1
		If (UnselectedTracks(i) = "") Then
			If instr(DiscogsTracksNum.Item(i), ".") <> 0 Then
				If SubTrackFlag = False Then
					SubTrackFlag = True
					SelectedTracksCount = SelectedTracksCount + 1
				End If
			Else
				If SubTrackFlag = True Then SubTrackFlag = False
				SelectedTracksCount = SelectedTracksCount + 1
			End If
		Else
			UnSelectedTracksCount = UnSelectedTracksCount + 1
		End If
	Next
	If (iMaxTracks - UnSelectedTracksCount) <> SDB.Tools.WebSearch.NewTracks.Count Then
		templateHTML = templateHTML &  "<tr><td colspan=3 align=center><b><span style=""color:#FF0000"">There are different numbers of tracks !</span></b></td></tr>"
		templateHTML = templateHTML &  "<tr><td colspan=3 align=center><br></td></tr>"
	End If

	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""releaseid"" ></td>"
	templateHTML = templateHTML &  "<td>Release:</td>"
	If (theMaster <> "") Then
		templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/release/<!RELEASEID!>"" target=""_blank""><!RELEASEID!></a> (Master: <a href=""http://www.discogs.com/master/<!MASTERID!>"" target=""_blank""><!MASTERID!></a>)</td>"
	Else
		templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/release/<!RELEASEID!>"" target=""_blank""><!RELEASEID!></a> (Master: N/A)</td>"
	End If
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""artist"" ></td>"
	templateHTML = templateHTML &  "<td>Artist:</td>"
	templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/artist/<!ARTIST!>"" target=""_blank""><!ARTIST!></a></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""album"" ></td>"
	templateHTML = templateHTML &  "<td>Album:</td>"
	templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/release/<!RELEASEID!>"" target=""_blank""><!ALBUMTITLE!></a></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""albumartist"" ><input type=checkbox id=""albumartistfirst"" ></td>"
	templateHTML = templateHTML &  "<td>Album Artist:</td>"
	templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/artist/<!ALBUMARTIST!>"" target=""_blank""><!ALBUMARTIST!></a></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""label"" ></td>"
	templateHTML = templateHTML &  "<td>Label:</td>"
	templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/label/<!LABEL!>"" target=""_blank""><!LABEL!></a></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""catalog"" ></td>"
	templateHTML = templateHTML &  "<td>Catalog#:</td>"
	templateHTML = templateHTML &  "<td><!CATALOG!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""format"" ></td>"
	templateHTML = templateHTML &  "<td>Format:</td>"
	templateHTML = templateHTML &  "<td><!FORMAT!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""country"" ></td>"
	templateHTML = templateHTML &  "<td>Country:</td>"
	templateHTML = templateHTML &  "<td><!COUNTRY!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox title=""If option set, the release date of this Discogs release will be saved"" id=""date"" ></td>"
	templateHTML = templateHTML &  "<td>Date:</td>"
	templateHTML = templateHTML &  "<td><!RELEASEDATE!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox title=""If option set, the release date of the Discogs master release will be saved"" id=""origdate"" ></td>"
	templateHTML = templateHTML &  "<td>Original Date:</td>"
	templateHTML = templateHTML &  "<td><!ORIGDATE!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""genre"" ><input type=checkbox id=""style"" ></td>"
	templateHTML = templateHTML &  "<td>Genre:</td>"
	templateHTML = templateHTML &  "<td><!GENRE!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td colspan=2>Release Data Quality:</td>"
	templateHTML = templateHTML &  "<td><!DATAQUALITY!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</td>"
	' Release Information End
	' Tracklisting Begin
	templateHTML = templateHTML & "<td align=left valign=top>"
	templateHTML = templateHTML & "<table border=0 cellspacing=0 cellpadding=1 class=tabletext>"
	templateHTML = templateHTML & "<tr>"

	If CheckOriginalDiscogsTrack Then
		templateHTML = templateHTML & "<td align=left><b>Discogs</b></td>"
	Else
		templateHTML = templateHTML & "<td> </td>"
	End If
	templateHTML = templateHTML & "<td><input type=checkbox id=""selectall""></td>"
	templateHTML = templateHTML & "<td align=center><input type=checkbox id=""discnum""></td>"
	templateHTML = templateHTML & "<td align=center><input type=checkbox id=""tracknum"" title=""If option NOT set, track numbers will not set automatically (useful when you didn't select all tracks from a release""></td>"
	templateHTML = templateHTML & "<td align=right><b>Artist</b></td>"
	templateHTML = templateHTML & "<td> </td>"
	templateHTML = templateHTML & "<td align=left><b>Title</b></td>"
	templateHTML = templateHTML & "<td align=right><b>Duration</b></td>"
	templateHTML = templateHTML & "</tr>"

	theTracks = ""
	currentCD = 0

	For i=0 To iMaxTracks - 1
		templateHTML = templateHTML &  "<tr>"
		
		If CheckOriginalDiscogsTrack Then
			templateHTML = templateHTML & "<td align=center>" & DiscogsTracksNum.Item(i) & "</td>"
		Else
			templateHTML = templateHTML & "<td> </td>"
		End If
		If(UnselectedTracks(i) = "") Then
			templateHTML = templateHTML & "<td><input type=checkbox id=""unselected["&i&"]"" checked></td>"
		Else
			templateHTML = templateHTML & "<td><input type=checkbox id=""unselected["&i&"]""></td>"
		End If
		templateHTML = templateHTML & "<td align=center>" & TracksCD.Item(i) & "</td>"
		templateHTML = templateHTML & "<td align=center>" & TracksNum.Item(i) & "</td>"
		templateHTML = templateHTML & "<td align=right>" & ArtistTitles.Item(i) & "</td>"
		templateHTML = templateHTML & "<td align=center><b>-</b></td>"
		templateHTML = templateHTML & "<td align=left>" & Tracks.Item(i) & "</td>"
		templateHTML = templateHTML & "<td align=right>" & Durations.Item(i) & "</td>"
		templateHTML = templateHTML & "</tr>"
		If(CheckLyricist and Lyricists.Item(i) <> "") Then templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>Lyrics: "& Lyricists.Item(i) &"</td></tr>"
		If(CheckComposer and Composers.Item(i) <> "") Then templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>Composer: "& Composers.Item(i) &"</td></tr>"
		If(CheckConductor and Conductors.Item(i) <> "") Then templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>Conductor: "& Conductors.Item(i) &"</td></tr>"
		If(CheckProducer and Producers.Item(i) <> "") Then templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>Producer: "& Producers.Item(i) &"</td></tr>"

		If(CheckInvolved and InvolvedPeople.Item(i) <> "") Then
			templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left><b>Involved People:</b></td></tr>"
			'SDB.Localize("Involved People")
			If CheckInvolvedPeopleSingleLine = True And InStr(InvolvedPeople.Item(i), ";") <> 0 Then
				Dim x
				tmp = Split(InvolvedPeople.Item(i), "; ")
				For each x in tmp
					templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>"& x &"</td></tr>"
				Next
			Else
				templateHTML = templateHTML & "<tr><td colspan=6></td><td colspan=2 align=left>"& InvolvedPeople.Item(i) &"</td></tr>"
			End If
		End If
	Next

	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</td>"
	' Tracklisting End


	templateHTML = templateHTML &  GetFooter()

	templateHTML = Replace(templateHTML, "<!RELEASEID!>", releaseID)
	templateHTML = Replace(templateHTML, "<!MASTERID!>", theMaster)
	templateHTML = Replace(templateHTML, "<!ARTIST!>", AlbumArtistTitle)
	templateHTML = Replace(templateHTML, "<!ALBUMARTIST!>",  AlbumArtist)
	templateHTML = Replace(templateHTML, "<!ALBUMTITLE!>", AlbumTitle)
	templateHTML = Replace(templateHTML, "<!RELEASEDATE!>", ReleaseDate)
	templateHTML = Replace(templateHTML, "<!ORIGDATE!>", OriginalDate)
	templateHTML = Replace(templateHTML, "<!LABEL!>", theLabels)
	templateHTML = Replace(templateHTML, "<!COUNTRY!>", theCountry)
	templateHTML = Replace(templateHTML, "<!COVER!>", theArt)
	templateHTML = Replace(templateHTML, "<!CATALOG!>", Catalog)
	templateHTML = Replace(templateHTML, "<!FORMAT!>", theFormat)
	templateHTML = Replace(templateHTML, "<!DATAQUALITY!>", DataQuality)

	theGenres = ""

	If Genres <> "" Then
		If CheckGenre Then
			theGenres = Genres
		Else
			theGenres = "<s>" + Genres + "</s>"
		End If
	End If

	If Styles <> "" Then
		If theGenres <> "" Then
			If CheckGenre Then
				theGenres = theGenres & Separator
			Else
				theGenres = theGenres & "<s>" & Separator & "</s>"
			End If
		End If
		If CheckStyle Then
			theGenres = theGenres & Styles
		Else
			theGenres = theGenres & "<s>" & Styles & "</s>"
		End If
	End If
	templateHTML = Replace(templateHTML, "<!GENRE!>", theGenres)

	
	REM Dim filesys, filetxt, logdatei
	REM 'Const ForReading = 1, ForWriting = 2, ForAppending = 8
	REM logdatei = SDB.ScriptsPath & "HTML.htm"
	REM Set filesys = CreateObject("Scripting.FileSystemObject")
	REM Set filetxt = filesys.OpenTextFile(logdatei, 2, True)
	REM filetxt.WriteLine(templateHTML)
	REM filetxt.Close

	WebBrowser.SetHTMLDocument templateHTML

	Dim templateHTMLDoc
	Set templateHTMLDoc = WebBrowser.Interf.Document

	Set checkBox = templateHTMLDoc.getElementById("album")
	checkBox.Checked = CheckAlbum
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("artist")
	checkBox.Checked = CheckArtist
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("albumartist")
	checkBox.Checked = CheckAlbumArtist
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("albumartistfirst")
	checkBox.Checked = CheckAlbumArtistFirst
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("date")
	checkBox.Checked = CheckDate
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("origdate")
	checkBox.Checked = CheckOrigDate
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("label")
	checkBox.Checked = CheckLabel
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("country")
	checkBox.Checked = CheckCountry
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("genre")
	checkBox.Checked = CheckGenre
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("style")
	checkBox.Checked = CheckStyle
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("cover")
	checkBox.Checked = CheckCover
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("smallcover")
	checkBox.Checked = SmallCover
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("catalog")
	checkBox.Checked = CheckCatalog
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("releaseid")
	checkBox.Checked = CheckRelease
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("involved")
	checkBox.Checked = CheckInvolved
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("lyricist")
	checkBox.Checked = CheckLyricist
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("composer")
	checkBox.Checked = CheckComposer
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("conductor")
	checkBox.Checked = CheckConductor
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("producer")
	checkBox.Checked = CheckProducer
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("discnum")
	checkBox.Checked = CheckDiscNum
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("tracknum")
	checkBox.Checked = CheckTrackNum
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("format")
	checkBox.Checked = CheckFormat
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("useanv")
	checkBox.Checked = CheckUseAnv
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("yearonlydate")
	checkBox.Checked = CheckYearOnlyDate
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("forcenumeric")
	checkBox.Checked = CheckForceNumeric
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("sidestodisc")
	checkBox.Checked = CheckSidesToDisc
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("forcedisc")
	checkBox.Checked = CheckForceDisc
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("nodisc")
	checkBox.Checked = CheckNoDisc
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("leadingzero")
	checkBox.Checked = CheckLeadingZero
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("titlefeaturing")
	checkBox.Checked = CheckTitleFeaturing
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set text = templateHTMLDoc.getElementById("TxtFeaturingName")
	text.value = TxtFeaturingName
	Script.RegisterEvent text, "onchange", "Update"
	Set checkbox = templateHTMLDoc.getElementById("FeaturingName")
	checkBox.Checked = CheckFeaturingName
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("comments")
	checkBox.Checked = CheckComment
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set text = templateHTMLDoc.getElementById("txtvarious")
	text.value = TxtVarious
	Script.RegisterEvent text, "onchange", "Update"
	Set checkBox = templateHTMLDoc.getElementById("various")
	checkBox.Checked = CheckVarious
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("SubTrackNameSelection")
	checkBox.Checked = SubTrackNameSelection
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("TurnOffSubTrack")
	checkBox.Checked = CheckTurnOffSubTrack
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("involvedpeoplesingle")
	checkBox.Checked = CheckInvolvedPeopleSingleLine
	Script.RegisterEvent checkBox, "onclick", "Update"
	

	Set listBox = templateHTMLDoc.getElementById("filtermediatype")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtermediaformat")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtercountry")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filteryear")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("load")
	Script.RegisterEvent listBox, "onchange", "Filter"

	For i=0 To iMaxTracks - 1
		Set checkBox = templateHTMLDoc.getElementById("unselected["&i&"]")
		Script.RegisterEvent checkBox, "onclick", "Unselect"
	Next

	Set checkBox = templateHTMLDoc.getElementById("selectall")
	checkBox.Checked = SelectAll
	Script.RegisterEvent checkBox, "onclick", "SwitchAll"

	Set listBox = templateHTMLDoc.getElementById("alternative")
	Script.RegisterEvent listBox, "onchange", "Alternative"

	Set submitButton = templateHTMLDoc.getElementById("saveoptions")
	Script.RegisterEvent submitButton, "onclick", "SaveOptions"

	Set submitButton = templateHTMLDoc.getElementById("showcountryfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowCountryFilter"

	Set submitButton = templateHTMLDoc.getElementById("showmediatypefilter")
	Script.RegisterEvent submitButton, "onclick", "ShowMediaTypeFilter"

	Set submitButton = templateHTMLDoc.getElementById("showmediaformatfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowMediaFormatFilter"

	Set submitButton = templateHTMLDoc.getElementById("showyearfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowYearFilter"

	Set submitButton = templateHTMLDoc.getElementById("moreimages")
	Script.RegisterEvent submitButton, "onclick", "MoreImages"

End Sub


Sub Update()

	Dim templateHTMLDoc, checkBox, text
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document

	Set checkBox = templateHTMLDoc.getElementById("album")
	CheckAlbum = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("artist")
	CheckArtist = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("albumartist")
	CheckAlbumArtist = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("albumartistfirst")
	CheckAlbumArtistFirst = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("date")
	CheckDate = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("origdate")
	CheckOrigDate = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("label")
	CheckLabel = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("genre")
	CheckGenre = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("style")
	CheckStyle = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("country")
	CheckCountry = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("cover")
	If Not CheckCover And checkBox.Checked Then
		SmallCover = False
		CheckCover = checkBox.Checked
	Else
		CheckCover = checkBox.Checked
		Set checkBox = templateHTMLDoc.getElementById("smallcover")
		If Not SmallCover And checkBox.Checked Then
			CheckCover = False
		End If
		SmallCover = checkBox.Checked
	End If
	Set checkBox = templateHTMLDoc.getElementById("catalog")
	CheckCatalog = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("releaseid")
	CheckRelease = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("involved")
	CheckInvolved = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("lyricist")
	CheckLyricist = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("composer")
	CheckComposer = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("conductor")
	CheckConductor = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("producer")
	CheckProducer = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("discnum")
	CheckDiscNum = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("tracknum")
	CheckTrackNum = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("format")
	CheckFormat = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("useanv")
	CheckUseAnv = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("yearonlydate")
	CheckYearOnlyDate = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("forcenumeric")
	CheckForceNumeric = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("sidestodisc")
	CheckSidesToDisc = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("forcedisc")
	If Not CheckForceDisc And checkBox.Checked Then
		CheckNoDisc = False
		CheckForceDisc = checkBox.Checked
	Else
		CheckForceDisc = checkBox.Checked
		Set checkBox = templateHTMLDoc.getElementById("nodisc")
		If Not CheckNoDisc And checkBox.Checked Then
			CheckForceDisc = False
		End If
		CheckNoDisc = checkBox.Checked
	End If
	Set checkBox = templateHTMLDoc.getElementById("leadingzero")
	CheckLeadingZero = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("titlefeaturing")
	CheckTitleFeaturing = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("FeaturingName")
	CheckFeaturingName = checkBox.Checked
	Set text = templateHTMLDoc.getElementById("TxtFeaturingName")
	TxtFeaturingName = text.Value
	Set checkBox = templateHTMLDoc.getElementById("comments")
	CheckComment = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("various")
	CheckVarious = checkBox.Checked
	Set text = templateHTMLDoc.getElementById("txtvarious")
	TxtVarious = text.Value
	Set checkBox = templateHTMLDoc.getElementById("SubTrackNameSelection")
	SubTrackNameSelection = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("TurnOffSubTrack")
	CheckTurnOffSubTrack = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("involvedpeoplesingle")
	CheckInvolvedPeopleSingleLine = checkBox.Checked
	
	OptionsChanged = True

	ReloadResults

End Sub





Sub Filter()

	Dim templateHTMLDoc, listBox
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document

	Set listBox = templateHTMLDoc.getElementById("filtermediatype")
	FilterMediaType = listBox.Value
	If FilterMediaType = "None" Then
		MediaTypeFilterList.Item(0) = "0"
	ElseIf FilterMediaType = "Use MediaType Filter" Then
		MediaTypeFilterList.Item(0) = "1"
	Else
		MediaTypeFilterList.Item(0) = FilterMediaType
	End If
	Set listBox = templateHTMLDoc.getElementById("filtermediaformat")
	FilterMediaFormat = listBox.Value
	If FilterMediaFormat = "None" Then
		MediaFormatFilterList.Item(0) = "0"
	ElseIf FilterMediaFormat = "Use MediaFormat Filter" Then
		MediaFormatFilterList.Item(0) = "1"
	Else
		MediaFormatFilterList.Item(0) = FilterMediaFormat
	End If
	Set listBox = templateHTMLDoc.getElementById("filtercountry")
	FilterCountry = listBox.Value
	If FilterCountry = "None" Then
		CountryFilterList.Item(0) = "0"
	ElseIf FilterCountry = "Use Country Filter" Then
		CountryFilterList.Item(0) = "1"
	Else
		CountryFilterList.Item(0) = FilterCountry
	End If
	Set listBox = templateHTMLDoc.getElementById("filteryear")
	FilterYear = listBox.Value
	If FilterYear = "None" Then
		YearFilterList.Item(0) = "0"
	ElseIf FilterYear = "Use Year Filter" Then
		YearFilterList.Item(0) = "1"
	Else
		YearFilterList.Item(0) = FilterYear
	End If

	Set listBox = templateHTMLDoc.getElementById("load")
	CurrentLoadType = listBox.Value

	If(CurrentLoadType = "Master Release") Then
		LoadMasterResults(SavedMasterId)
	ElseIf(CurrentLoadType = "Releases of Artist") Then
		LoadArtistResults(SavedArtistId)
	ElseIf(CurrentLoadType = "Releases of Label") Then
		LoadLabelResults(SavedLabelId)
	Else
		FindResults(SavedSearchTerm)
	End If

End Sub


Sub Alternative()

	Dim templateHTMLDoc
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	SavedSearchTerm =  templateHTMLDoc.getElementById("alternative").Value
	CurrentLoadType = "Search Results"
	FindResults(SavedSearchTerm)
	
End Sub


Sub SaveOptions()

	Dim a, tmp
	' save options if ini exists
	If Not (ini Is Nothing) Then
		ini.BoolValue("DiscogsAutoTagWeb","CheckAlbum") = CheckAlbum
		ini.BoolValue("DiscogsAutoTagWeb","CheckArtist") = CheckArtist
		ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtist") = CheckAlbumArtist
		ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst") = CheckAlbumArtistFirst
		ini.BoolValue("DiscogsAutoTagWeb","CheckLabel") = CheckLabel
		ini.BoolValue("DiscogsAutoTagWeb","CheckDate") = CheckDate
		ini.BoolValue("DiscogsAutoTagWeb","CheckOrigDate") = CheckOrigDate
		ini.BoolValue("DiscogsAutoTagWeb","CheckGenre") = CheckGenre
		ini.BoolValue("DiscogsAutoTagWeb","CheckStyle") = CheckStyle
		ini.BoolValue("DiscogsAutoTagWeb","CheckCountry") = CheckCountry
		ini.BoolValue("DiscogsAutoTagWeb","CheckCatalog") = CheckCatalog
		ini.BoolValue("DiscogsAutoTagWeb","CheckRelease") = CheckRelease
		ini.BoolValue("DiscogsAutoTagWeb","CheckInvolved") = CheckInvolved
		ini.BoolValue("DiscogsAutoTagWeb","CheckLyricist") = CheckLyricist
		ini.BoolValue("DiscogsAutoTagWeb","CheckComposer") = CheckComposer
		ini.BoolValue("DiscogsAutoTagWeb","CheckConductor") = CheckConductor
		ini.BoolValue("DiscogsAutoTagWeb","CheckProducer") = CheckProducer
		ini.BoolValue("DiscogsAutoTagWeb","CheckDiscNum") = CheckDiscNum
		ini.BoolValue("DiscogsAutoTagWeb","CheckTrackNum") = CheckTrackNum
		ini.BoolValue("DiscogsAutoTagWeb","CheckFormat") = CheckFormat
		ini.BoolValue("DiscogsAutoTagWeb","CheckUseAnv") = CheckUseAnv
		ini.BoolValue("DiscogsAutoTagWeb","CheckYearOnlyDate") = CheckYearOnlyDate
		ini.BoolValue("DiscogsAutoTagWeb","CheckForceNumeric") = CheckForceNumeric
		ini.BoolValue("DiscogsAutoTagWeb","CheckSidesToDisc") = CheckSidesToDisc
		ini.BoolValue("DiscogsAutoTagWeb","CheckForceDisc") = CheckForceDisc
		ini.BoolValue("DiscogsAutoTagWeb","CheckNoDisc") = CheckNoDisc
		ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZero") = CheckLeadingZero
		ini.StringValue("DiscogsAutoTagWeb","ReleaseTag") = ReleaseTag
		ini.StringValue("DiscogsAutoTagWeb","CatalogTag") = CatalogTag
		ini.StringValue("DiscogsAutoTagWeb","CountryTag") = CountryTag
		ini.StringValue("DiscogsAutoTagWeb","FormatTag") = FormatTag
		ini.BoolValue("DiscogsAutoTagWeb","CheckVarious") = CheckVarious
		ini.StringValue("DiscogsAutoTagWeb","TxtVarious") = TxtVarious
		ini.BoolValue("DiscogsAutoTagWeb","CheckTitleFeaturing") = CheckTitleFeaturing
		ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName") = TxtFeaturingName
		ini.BoolValue("DiscogsAutoTagWeb","CheckFeaturingName") = CheckFeaturingName
		ini.BoolValue("DiscogsAutoTagWeb","CheckComment") = CheckComment
		ini.BoolValue("DiscogsAutoTagWeb","SubTrackNameSelection") = SubTrackNameSelection
		ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = ArtistSeparator
		ini.BoolValue("DiscogsAutoTagWeb","CheckTurnOffSubTrack") = CheckTurnOffSubTrack
		ini.BoolValue("DiscogsAutoTagWeb","CheckInvolvedPeopleSingleLine") = CheckInvolvedPeopleSingleLine

		tmp = CountryFilterList.Item(0)
		For a = 1 To CountryList.Count - 1
			tmp = tmp & "," & CountryFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter") = tmp
		tmp = MediaTypeFilterList.Item(0)
		For a = 1 To MediaTypeList.Count - 1
			tmp = tmp & "," & MediaTypeFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter") = tmp
		tmp = MediaFormatFilterList.Item(0)
		For a = 1 To MediaFormatList.Count - 1
			tmp = tmp & "," & MediaFormatFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter") = tmp
		tmp = YearFilterList.Item(0)
		For a = 1 To YearList.Count - 1
			tmp = tmp & "," & YearFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = tmp
	End If

End Sub

' Format Error Message
Sub FormatErrorMessage(ErrorMessage)

	Dim templateHTML, listBox, templateHTMLDoc, submitButton
	templateHTML = ""
	templateHTML = templateHTML &  GetHeader()
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td colspan=4 align=center><p><b>" & ErrorMessage & "</b></p></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  GetFooter()

	WebBrowser.SetHTMLDocument templateHTML

	Set templateHTMLDoc = WebBrowser.Interf.Document

	Set listBox = templateHTMLDoc.getElementById("alternative")
	Script.RegisterEvent listBox, "onchange", "Alternative"

	Set listBox = templateHTMLDoc.getElementById("filtermediatype")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtermediaformat")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtercountry")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filteryear")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("load")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set submitButton = templateHTMLDoc.getElementById("showcountryfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowCountryFilter"
	Set submitButton = templateHTMLDoc.getElementById("showmediatypefilter")
	Script.RegisterEvent submitButton, "onclick", "ShowMediaTypeFilter"
	Set submitButton = templateHTMLDoc.getElementById("showmediaformatfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowMediaFormatFilter"
	Set submitButton = templateHTMLDoc.getElementById("showyearfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowYearFilter"

	SDB.Tools.WebSearch.ClearTracksData

End Sub


Function JSONParser_find_result(searchURL, ArrayName, searchURL_F, searchURL_L)

	Dim oXMLHTTP, r, f, a
	WriteLog ("Start JSONParser_find_result")
	WriteLog("Arrayname=" & ArrayName)
	WriteLog("Complete searchURL=" & searchURL_F & searchURL & searchURL_L)
	' use json api with vbsjson class at start of file now
	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
	
	Dim json
	Set json = New VbsJson

	Dim response
	Dim format, title, country, v_year, label, artist, Rtype, catNo, main_release, tmp, ReleaseDesc, FilterFound, SongCount, SongCountMax, isRelease, listCount
	Dim Page, SongPages

	oXMLHTTP.Open "POST", "http://www.germanc64.de/mm/oauth/check_new.php", False
	oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	REM oXMLHTTP.setRequestHeader "Content-Type","application/json"
	oXMLHTTP.setRequestHeader "User-Agent","MediaMonkeyDiscogsAutoTagWeb/2.0 +http://mediamonkey.com"
	oXMLHTTP.send ("at=" & AccessToken & "&ats=" & AccessTokenSecret & "&searchURL=" & searchURL & "&searchURL_F=" & searchURL_F & "&searchURL_L=" & searchURL_L)

	If oXMLHTTP.Status = 200 Then

		Set response = json.Decode(oXMLHTTP.responseText)
		'check if any results
		'and add titles to drop down
		'msgbox response(ArrayName)(0)("title")

		SongCount = 0
		SongCountMax = response("pagination")("items")
		WriteLog ("SongCountMax=" & SongCountMax)

		isRelease = False
		If Results.Count = 1 Then isRelease = True
		SongPages = response("pagination")("pages")
		WriteLog ("SongPages=" & SongPages)
		For Page = 1 to SongPages
			If Page <> 1 Then
				oXMLHTTP.Open "POST", "http://www.germanc64.de/mm/oauth/check_new.php", False
				oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"  
				oXMLHTTP.setRequestHeader "User-Agent","MediaMonkeyDiscogsAutoTagBatch/2.0 +http://mediamonkey.com"
				WriteLog "SearchURL=" & SearchURL & "&page=" & Page
				oXMLHTTP.send ("at=" & AccessToken & "&ats=" & AccessTokenSecret & "&searchURL=" & searchURL & "&searchURL_F=" & searchURL_F & "&searchURL_L=" & searchURL_L)
				Set response = json.Decode(oXMLHTTP.responseText)
			End If
			For Each r In response(ArrayName)
				format = ""
				title = ""
				country = ""
				v_year = ""
				artist = ""
				label = ""
				Rtype = ""
				catNo = ""
				main_release = ""

				title = response(ArrayName)(r)("title")
				Set tmp = response(ArrayName)(r)
				If tmp.Exists("artist") Then
					artist = tmp("artist")
				End If
				If tmp.Exists("main_release") Then
					main_release = tmp("main_release")
				End If
				If ArrayName = "results" Then
					If tmp.Exists("format") Then
						For Each f In response(ArrayName)(r)("format")
							format = format & response(ArrayName)(r)("format")(f) & ", "
						Next
						If Len(format) <> 0 Then format = Left(format, Len(format)-2)
					End If
				Else
					If tmp.Exists("format") Then
						format = response(ArrayName)(r)("format")
					End If
				End If

				country = response(ArrayName)(r)("country")
				If ArrayName = "versions" Then
					If tmp.Exists("released") Then
						v_year = response(ArrayName)(r)("released")
					End If
				Else
					If tmp.Exists("year") Then
						v_year = response(ArrayName)(r)("year")
					End If
				End If
				If tmp.Exists("catno") Then
					catNo = response(ArrayName)(r)("catno")
				End If
				If tmp.Exists("type") Then
					Rtype = response(ArrayName)(r)("type")
				End If
				If ArrayName = "results" Then
					If tmp.Exists("label") Then
						For Each f In response(ArrayName)(r)("label")
							If label <> "" Then
								If Left(label, Len(label)-2) <> response(ArrayName)(r)("label")(f) Then
									label = label & response(ArrayName)(r)("label")(f) & ", "
								End If
							Else
								label = response(ArrayName)(r)("label")(f) & ", "
							End If
						Next
						If Len(label) <> 0 Then label = Left(label, Len(label)-2)
					End If
				Else
					If tmp.Exists("label") Then label = response(ArrayName)(r)("label")
				End If
				ReleaseDesc = ""
				Do
					If FilterMediaType = "Use MediaType Filter" And Format <> "" Then
						FilterFound = False
						For a = 1 To MediaTypeList.Count - 1
							If InStr(Format, MediaTypeList.Item(a)) <> 0 And MediaTypeFilterList.Item(a) = "1" Then FilterFound = True
						Next
						If FilterFound = False Then Exit Do
					End If
					If(FilterMediaType <> "None" And FilterMediaType <> "Use MediaType Filter" And InStr(format, FilterMediaType) = 0 And format <> "") Then Exit Do

					If FilterMediaFormat = "Use MediaFormat Filter" And format <> "" Then
						FilterFound = False
						For a = 1 To MediaFormatList.Count - 1
							If InStr(format, MediaFormatList.Item(a)) <> 0 And MediaFormatFilterList.Item(a) = "1" Then FilterFound = True
						Next
						If FilterFound = False Then Exit Do
					End If
					If(FilterMediaFormat <> "None" And FilterMediaFormat <> "Use MediaFormat Filter" And InStr(format, FilterMediaFormat) = 0 And Format <> "") Then Exit Do

					If FilterCountry = "Use Country Filter" And country <> "" Then
						FilterFound = False
						For a = 1 To CountryList.Count - 1
							If InStr(country, CountryList.Item(a)) <> 0 And CountryFilterList.Item(a) = "1" Then FilterFound = True
						Next
						If FilterFound = False Then Exit Do
					End If
					If(FilterCountry <> "None" And FilterCountry <> "Use Country Filter" And InStr(country, FilterCountry) = 0 And country <> "") Then Exit Do

					If FilterYear = "Use Year Filter" And v_year <> "" Then
						FilterFound = False
						For a = 1 To YearList.Count - 1
							If InStr(v_year, YearList.Item(a)) <> 0 And YearFilterList.Item(a) = "1" Then FilterFound = True
						Next
						If FilterFound = False Then Exit Do
					End If
					If(FilterYear <> "None" And FilterYear <> "Use Year Filter" And InStr(v_year, FilterYear) = 0 And v_year <> "") Then Exit Do

					If artist <> "" Then ReleaseDesc = ReleaseDesc & " " & artist End If
					If artist <> "" and title <> "" Then ReleaseDesc = ReleaseDesc & " -" End If
					If title <> "" Then ReleaseDesc = ReleaseDesc & " " & title End If
					If format <> "" Then ReleaseDesc = ReleaseDesc & " [" & format & "]" End If
					If label <> "" Then ReleaseDesc = ReleaseDesc & " " & label End If
					If country <> "" Then ReleaseDesc = ReleaseDesc & " / " & country End If
					If v_year <> "" Then ReleaseDesc = ReleaseDesc & " (" & v_year & ")" End If
					If catNo <> "" Then ReleaseDesc = ReleaseDesc & " catNo:" & catNo End If
					If Rtype = "master" Then ReleaseDesc = ReleaseDesc & " *" End If

					Results.Add ReleaseDesc
					ResultsReleaseID.Add response(ArrayName)(r)("id")
					SongCount = SongCount + 1
					If SongCount = 250 Then Exit Do
				Loop While False
				If SongCount = 250 Then Exit For
			Next
			If SongCount = 250 Then Exit For
		Next
	End If
	ListCount = 1
	For r = 1 to Results.Count
		If r= 1 and isRelease = True Then
			Results.Item(0) = "(" & SongCountMax & ") " & Results.Item(0)
		Else
			If SongCount <> SongCountMax Then
				Results.Item(r-1) = "(" & ListCount & "/" & SongCount & "/" & SongCountMax & ") " & Results.Item(r-1)
			Else
				Results.Item(r-1) = "(" & ListCount & "/" & SongCountMax & ") " & Results.Item(r-1)
			End IF
			ListCount = ListCount + 1
		End If
	Next
	WriteLog ("End JSONParser_find_result")
End Function


Function ReloadMaster(SavedMasterId)

	Dim oXMLHTTP, masterURL
	masterURL = SavedMasterId
	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")

	Dim json
	Set json = New VbsJson
	Dim response

	oXMLHTTP.Open "POST", "http://www.germanc64.de/mm/oauth/check_new.php", False
	oXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	REM oXMLHTTP.setRequestHeader "Content-Type","application/json"
	oXMLHTTP.setRequestHeader "User-Agent","MediaMonkeyDiscogsAutoTagWeb/2.0 +http://mediamonkey.com"
	oXMLHTTP.send ("at=" & AccessToken & "&ats=" & AccessTokenSecret & "&searchURL=" & masterURL & "&searchURL_F=http://api.discogs.com/masters/&searchURL_L=")

	If oXMLHTTP.Status = 200 Then
		Set response = json.Decode(oXMLHTTP.responseText)
		If response.Exists("year") Then
			OriginalDate = response("year")
		Else
			OriginalDate = ""
		End If
	End If

	ReloadMaster = OriginalDate

End Function



Function get_release_ID(FirstTrack)

	CurrentResultID = ""
	If ReleaseTag = "Custom1" Then CurrentResultID = FirstTrack.Custom1
	If ReleaseTag = "Custom2" Then CurrentResultID = FirstTrack.Custom2
	If ReleaseTag = "Custom3" Then CurrentResultID = FirstTrack.Custom3
	If ReleaseTag = "Custom4" Then CurrentResultID = FirstTrack.Custom4
	If ReleaseTag = "Custom5" Then CurrentResultID = FirstTrack.Custom5
	If ReleaseTag = "Grouping" Then CurrentResultID = FirstTrack.Grouping
	If ReleaseTag = "ISRC" Then CurrentResultID = FirstTrack.ISRC
	If ReleaseTag = "Encoding" Then CurrentResultID = FirstTrack.Encoding
	If ReleaseTag = "Copyright" Then CurrentResultID = FirstTrack.Copyright

	get_release_ID = CurrentResultID

End Function



Sub Unselect()

	Dim templateHTMLDoc, i, checkBox

	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document

	For i=0 To iMaxTracks - 1
		Set checkBox = templateHTMLDoc.getElementById("unselected["&i&"]")
		If checkBox.Checked Then
			UnselectedTracks(i) = ""
		Else
			UnselectedTracks(i) = "x"
		End If
	Next

	ReloadResults

End Sub


Sub ShowCountryFilter

	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, WebBrowser2, countrybutton, FilterCountry, FilterFound
	Dim i, a
	Set Form = UI.NewForm
	Form.Common.Width = 675
	Form.Common.Height = 600
	Form.FormPosition = 4
	Form.Caption = "Choose the country's to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = CountryList
	CountColumn = (CountryList.Count - 1) / 94
	iWidth = (CountryList.Count - 1) / 94 * 200
	filterHTML = GetFilterHTML(iWidth, 93, CountColumn)

	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 30
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True

	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True

	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 15
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "Btn3Click"

	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "Btn4Click"

	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100

	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront

	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document

	For i = 1 To CountryList.Count - 1
		Set countrybutton = filterHTMLDoc.getElementById("Filter" & i)
		If CountryFilterList.Item(i) = "1" Then
			countrybutton.checked = True
		End If
	Next

	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To CountryList.Count - 1
			Set countrybutton = filterHTMLDoc.getElementById("Filter" & a)
			If countrybutton.checked = True Then
				CountryFilterList.Item(a) = "1"
				FilterFound = True
			Else
				CountryFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterCountry = "None"
			CountryFilterList.Item(0) = "0"
		Else
			FilterCountry = "Use Country Filter"
			CountryFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults(SavedSearchTerm)
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If

End Sub


Sub ShowMediaFormatFilter

	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, WebBrowser2, MediaFormatButton, FilterMediaFormat, FilterFound
	Dim i, a
	Set Form = UI.NewForm
	Form.Common.Width = 380
	Form.Common.Height = 700
	Form.FormPosition = 4
	Form.Caption = "Choose the MediaFormat to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = MediaFormatList
	iWidth = (MediaFormatList.Count - 1) / 24 * 150
	CountColumn = (MediaFormatList.Count - 1) / 24

	filterHTML = GetFilterHTML(iWidth, 23, CountColumn)

	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 20
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True

	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True

	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 5
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "Btn3Click"

	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "Btn4Click"

	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100

	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront

	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document

	For i = 1 To MediaFormatList.Count - 1
		Set MediaFormatButton = filterHTMLDoc.getElementById("Filter" & i)
		If MediaFormatFilterList.Item(i) = "1" Then
			MediaFormatButton.checked = True
		End If
	Next

	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To MediaFormatList.Count - 1
			Set MediaFormatButton = filterHTMLDoc.getElementById("Filter" & a)
			If MediaFormatButton.checked = True Then
				MediaFormatFilterList.Item(a) = "1"
				FilterFound = True
			Else
				MediaFormatFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterMediaFormat = "None"
			MediaFormatFilterList.Item(0) = "0"
		Else
			FilterMediaFormat = "Use MediaFormat Filter"
			MediaFormatFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults(SavedSearchTerm)
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If

End Sub


Sub ShowMediaTypeFilter

	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, WebBrowser2, MediaTypeButton, FilterMediaType, FilterFound
	Dim i, a
	Set Form = UI.NewForm
	Form.Common.Width = 420
	Form.Common.Height = 600
	Form.FormPosition = 4
	Form.Caption = "Choose the MediaType to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = MediaTypeList
	iWidth = (MediaTypeList.Count - 1) / 19 * 175
	CountColumn = (MediaTypeList.Count - 1) / 19

	filterHTML = GetFilterHTML(iWidth, 18, CountColumn)

	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 30
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True

	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True

	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 15
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "Btn3Click"

	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "Btn4Click"

	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100

	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront

	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document

	For i = 1 To MediaTypeList.Count - 1
		Set MediaTypeButton = filterHTMLDoc.getElementById("Filter" & i)
		If MediaTypeFilterList.Item(i) = "1" Then
			MediaTypeButton.checked = True
		End If
	Next

	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To MediaTypeList.Count - 1
			Set MediaTypeButton = filterHTMLDoc.getElementById("Filter" & a)
			If MediaTypeButton.checked = True Then
				MediaTypeFilterList.Item(a) = "1"
				FilterFound = True
			Else
				MediaTypeFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterMediaType = "None"
			MediaTypeFilterList.Item(0) = "0"
		Else
			FilterMediaType = "Use MediaType Filter"
			MediaTypeFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults(SavedSearchTerm)
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If

End Sub


Sub ShowYearFilter

	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, WebBrowser2, YearButton, FilterYear, FilterFound
	Dim i, a, row
	Set Form = UI.NewForm
	Form.Common.Width = 550
	Form.Common.Height = 550
	Form.FormPosition = 4
	Form.Caption = "Choose the Year to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = YearList
	'CountColumn = 6
	If ((YearList.Count - 1) / 6) = Int((YearList.Count - 1) / 6) Then
		iWidth = (YearList.Count - 1) / 6 * 25
		row = Int((YearList.Count - 1) / 6)
		CountColumn = 6
	Else
		row = Int((YearList.Count - 1) / 6) + 1
		iWidth = (YearList.Count - 1) / 6 * 25
		CountColumn = 6
	End If

	filterHTML = GetFilterHTML(iWidth, row-1, CountColumn)

	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 30
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True

	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True

	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 15
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "Btn3Click"

	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "Btn4Click"

	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100

	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront

	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document

	For i = 1 To YearList.Count - 1
		Set Yearbutton = filterHTMLDoc.getElementById("Filter" & i)
		If YearFilterList.Item(i) = "1" Then
			Yearbutton.checked = True
		End If
	Next

	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To YearList.Count - 1
			Set Yearbutton = filterHTMLDoc.getElementById("Filter" & a)
			If Yearbutton.checked = True Then
				YearFilterList.Item(a) = "1"
				FilterFound = True
			Else
				YearFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterYear = "None"
			YearFilterList.Item(0) = "0"
		Else
			FilterYear = "Use Year Filter"
			YearFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults(SavedSearchTerm)
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If

End Sub


Sub Btn3Click

	Dim WebBrowser2, FilterList, filterHTMLDoc, a, filterbutton
	Set WebBrowser2 = SDB.Objects("WebBrowser2")
	Set FilterList = SDB.Objects("Filter")
	Set filterHTMLDoc = WebBrowser2.Interf.Document
	For a = 1 To FilterList.Count - 1
		Set filterbutton = filterHTMLDoc.getElementById("Filter" & a)
		filterbutton.checked = True
	Next

End Sub


Sub Btn4Click

	Dim WebBrowser2, FilterList, filterHTMLDoc, a, filterbutton
	Set WebBrowser2 = SDB.Objects("WebBrowser2")
	Set FilterList = SDB.Objects("Filter")
	Set filterHTMLDoc = WebBrowser2.Interf.Document
	For a = 1 To FilterList.Count - 1
		Set filterbutton = filterHTMLDoc.getElementById("Filter" & a)
		filterbutton.checked = False
	Next

End Sub


Function GetFilterHTML(Width, Row, CountColumn)

	Dim FilterList, filterHTML, i, a
	Set FilterList = SDB.Objects("Filter")
	filterHTML = "<HTML>"
	filterHTML = filterHTML & "<HEAD>"
	filterHTML = filterHTML & "<style type=""text/css"" media=""screen"">"
	filterHTML = filterHTML & ".tabletext { font-family: Arial, Helvetica, sans-serif; font-size: 8pt;}"
	filterHTML = filterHTML & "</style>"
	filterHTML = filterHTML & "</HEAD>"
	filterHTML = filterHTML & "<table border=0 width=" & Width & " cellspacing=0 cellpadding=1 class=tabletext>"
	For i = 0 To Row
		filterHTML = filterHTML &  "<tr>"
		For a = 1 To CountColumn
			If FilterList.Count = a + (i * CountColumn) Then Exit For
			filterHTML = filterHTML &  "<td><input type=checkbox id=""Filter" & a + (i * CountColumn) & """ >" & FilterList.Item(a + (i * CountColumn))
			filterHTML = filterHTML &  "</td>"
		Next
		filterHTML = filterHTML &  "</tr>"
	Next
	filterHTML = filterHTML &  "</table>"
	filterHTML = filterHTML &  "</body>"
	filterHTML = filterHTML &  "</HTML>"
	GetFilterHTML = filterHTML

End Function


Sub MoreImages()

	Set ImageTypeList = SDB.NewStringList

	ImageTypeList.Add SDB.Localize("Not specified")		'0
	ImageTypeList.Add SDB.Localize("Cover (front)")		'3
	ImageTypeList.Add SDB.Localize("Cover (back)")		'4
	ImageTypeList.Add SDB.Localize("Leaflet Page")
	ImageTypeList.Add SDB.Localize("Media Label")
	ImageTypeList.Add SDB.Localize("Lead Artist")
	ImageTypeList.Add SDB.Localize("Artist")
	ImageTypeList.Add SDB.Localize("Conductor")
	ImageTypeList.Add SDB.Localize("Band")
	ImageTypeList.Add SDB.Localize("Composer")
	ImageTypeList.Add SDB.Localize("Lyricist")
	ImageTypeList.Add SDB.Localize("Recording Location")
	ImageTypeList.Add SDB.Localize("During Recording")
	ImageTypeList.Add SDB.Localize("During Performance")
	ImageTypeList.Add SDB.Localize("Video Screen Capture")
	ImageTypeList.Add SDB.Localize("Illustration")
	ImageTypeList.Add SDB.Localize("Band Logotype")
	ImageTypeList.Add SDB.Localize("Publisher Logotype")		'20

	Dim iWidth, imageHTML, imageHTMLDoc, i, j
	imageHTML = "<HTML>"
	imageHTML = imageHTML &  "<HEAD>"
	imageHTML = imageHTML &  "<style type=""text/css"" media=""screen"">"
	imageHTML = imageHTML &  ".tabletext { font-family: Arial, Helvetica, sans-serif; font-size: 8pt;}"
	imageHTML = imageHTML &  "</style>"
	imageHTML = imageHTML &  "</HEAD>"
	imageHTML = imageHTML &  "<body bgcolor=""#FFFFFF"">"
	iWidth = (ImageList.Count - 1) * 200
	imageHTML = imageHTML &  "<table border=0 width=" & iWidth & " cellspacing=0 cellpadding=1 class=tabletext>"
	imageHTML = imageHTML &  "<tr>"

	For i = 0 To ImageList.Count - 1
		imageHTML = imageHTML &  "<td><img border=""0"" src=""" & ImageList.Item(i) & """ width=""180"" height=""180""></td>"
	Next
	imageHTML = imageHTML &  "</tr><tr>"
	For i = 0 To ImageList.Count - 1
		imageHTML = imageHTML &  "<td>"
		imageHTML = imageHTML &  "<select id=""ImageType" & i & """ class=tabletext>"
		For j = 0 To ImageTypeList.Count - 1
			If SaveImageType.Item(i) <> ImageTypeList.Item(j) Then
				imageHTML = imageHTML &  "<option value=""" & ImageTypeList.Item(j) & """>" & ImageTypeList.Item(j) & "</option>"
			Else
				imageHTML = imageHTML &  "<option value=""" & ImageTypeList.Item(j) & """ selected>" & ImageTypeList.Item(j) & "</option>"
			End If
		Next
		imageHTML = imageHTML &  "</select></td>"
	Next

	imageHTML = imageHTML &  "</tr><tr>"
	For i = 0 To ImageList.Count - 1
		imageHTML = imageHTML &  "<td><input type=checkbox id=""SaveImage" & i & """ >Save Image"
		imageHTML = imageHTML &  "</td>"
	Next

	imageHTML = imageHTML &  "</tr><tr>"
	For i = 0 To ImageList.Count - 1
		If CoverStorage = 1 Or CoverStorage = 3 Then
			imageHTML = imageHTML &  "<td><input type=text id=""FileName" & i & """ >"
		Else
			imageHTML = imageHTML &  "<td>Store in tag"
		End If
		imageHTML = imageHTML &  "</td>"
	Next

	imageHTML = imageHTML &  "</tr>"

	imageHTML = imageHTML &  "</table>"
	imageHTML = imageHTML &  "</body>"
	imageHTML = imageHTML &  "</HTML>"



	Dim Form
	Set Form = UI.NewForm 
	Form.Common.ClientWidth = 800
	Form.Common.ClientHeight = 400
	Form.FormPosition = 4
	'Form.SavePositionName = "TheDiscogsWindow" 
	Form.Caption = "Choose additional Images for the Release"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("ImageForm") = Form

	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35

	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width -30
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True  

	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width -5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1 
	Btn2.Default = True

	Set WebBrowser3 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser3.Common.Align = 5      ' Fill whole client rectangle
	WebBrowser3.Common.ControlName = "WebBrowser3"
	WebBrowser3.Common.Top = 100
	WebBrowser3.Common.Left = 100

	SDB.Objects("WebBrowser3") = WebBrowser3
	WebBrowser3.Interf.Visible = True
	WebBrowser3.Common.BringToFront

	WebBrowser3.SetHTMLDocument imageHTML
	Set imageHTMLDoc = WebBrowser3.Interf.Document

	Dim saveimagebutton, text
	For i = 0 To ImageList.Count - 1
		Set saveimagebutton = imageHTMLDoc.getElementById("SaveImage" & i)
		If SaveImage.Item(i) = 1 Then
			saveimagebutton.checked = 1
		End If
		If CoverStorage = 1 Or CoverStorage = 3 Then
			Set text = imageHTMLDoc.getElementById("FileName" & i)
			text.value = "folder" & i & ".jpg"
		End If
	Next

	If Form.ShowModal = 1 Then
		SetSaveImages()
	End If

	WebBrowser3.SetHTMLDocument ""
	SDB.Objects("ImageForm") = Nothing
	WebBrowser3.Common.DestroyControl      ' Destroy the external control
	Set WebBrowser3 = Nothing              ' Release global variable
	SDB.Objects("WebBrowser3") = Nothing

End Sub


Sub SetSaveImages()

	Dim imageHTMLDoc, i, checkbox, listbox, text

	Set SaveImage = SDB.NewStringList
	Set SaveImageType = SDB.NewStringList
	Set FileNameList = SDB.NewStringList

	Set imageHTMLDoc = WebBrowser3.Interf.Document
	For i = 0 To ImageList.Count - 1
		Set checkBox = imageHTMLDoc.getElementById("SaveImage" & i)
		If checkBox.Checked Then
			SaveImage.Add "1"
			Set listBox = imageHTMLDoc.getElementById("ImageType" & i)
			SaveImageType.Add listBox.Value
			If CoverStorage = 1 Or CoverStorage = 3 Then
				Set text = imageHTMLDoc.getElementById("FileName" & i)
				FileNameList.Add text.Value
			Else
				FileNameList.Add "nothing"
			End If
		Else
			SaveImage.Add "0"
			SaveImageType.Add "other"
			FileNameList.Add "nothing"
		End If
	Next

End Sub


function getimages(DownloadDest, LocalFile)

	dim xmlhttp
	set xmlhttp=createobject("MSXML2.XMLHTTP.3.0")

	xmlhttp.Open "GET", DownloadDest, False

	xmlhttp.Send
	If xmlhttp.Status = 200 Then
		Dim objStream
		set objStream = CreateObject("ADODB.Stream")
		objStream.Type = 1 'adTypeBinary
		objStream.Open
		objStream.Write xmlhttp.responseBody
		objStream.SaveToFile LocalFile
		objStream.Close
		set objStream = Nothing
		set xmlhttp=Nothing
		getimages = LocalFile
	Else
		set xmlhttp=Nothing
		getimages = ""
	End If

End function


Sub WriteOptions()

	WriteLog " "
	WriteLog "Options:"
	WriteLog "CheckAlbum=" & CheckAlbum
	WriteLog "CheckArtist=" & CheckArtist
	WriteLog "CheckAlbumArtist=" & CheckAlbumArtist
	WriteLog "CheckAlbumArtistFirst=" & CheckAlbumArtistFirst
	WriteLog "CheckLabel=" & CheckLabel
	WriteLog "CheckDate=" & CheckDate
	WriteLog "CheckOrigDate=" & CheckOrigDate
	WriteLog "CheckGenre=" & CheckGenre
	WriteLog "CheckStyle=" & CheckStyle
	WriteLog "CheckCountry=" & CheckCountry
	REM WriteLog "CheckCover=" & CheckCover
	WriteLog "CheckSmallCover=" & CheckSmallCover
	WriteLog "CheckCatalog=" & CheckCatalog
	WriteLog "CheckRelease=" & CheckRelease
	WriteLog "CheckInvolved=" & CheckInvolved
	WriteLog "CheckLyricist=" & CheckLyricist
	WriteLog "CheckComposer=" & CheckComposer
	WriteLog "CheckConductor=" & CheckConductor
	WriteLog "CheckProducer=" & CheckProducer
	WriteLog "CheckDiscNum=" & CheckDiscNum
	WriteLog "CheckTrackNum=" & CheckTrackNum
	WriteLog "CheckFormat=" & CheckFormat
	WriteLog "CheckUseAnv=" & CheckUseAnv
	WriteLog "CheckYearOnlyDate=" & CheckYearOnlyDate
	WriteLog "CheckForceNumeric=" & CheckForceNumeric
	WriteLog "CheckSidesToDisc=" & CheckSidesToDisc
	WriteLog "CheckForceDisc=" & CheckForceDisc
	WriteLog "CheckNoDisc=" & CheckNoDisc
	WriteLog "CheckLeadingZero=" & CheckLeadingZero
	WriteLog "ReleaseTag=" & ReleaseTag
	WriteLog "CatalogTag=" & CatalogTag
	WriteLog "CountryTag=" & CountryTag
	WriteLog "FormatTag=" & FormatTag
	WriteLog "CheckVarious=" & CheckVarious
	WriteLog "TxtVarious=" & TxtVarious
	WriteLog "CheckTitleFeaturing=" & CheckTitleFeaturing
	WriteLog "TxtFeaturingName=" & TxtFeaturingName
	WriteLog "CheckFeaturingName=" & CheckFeaturingName
	WriteLog "CheckComment=" & CheckComment
	WriteLog "SubTrackNameSelection=" & SubTrackNameSelection
	WriteLog "CheckTurnOffSubTrack=" & CheckTurnOffSubTrack
	WriteLog "CheckInvolvedPeopleSingleLine=" & CheckInvolvedPeopleSingleLine
	WriteLog "ArtistSeparator=" & ArtistSeparator

End Sub


Sub WriteLogInit

	Dim logdatei
	logdatei = SDB.ScriptsPath & "Discogs_Script.log"
	If SDB.Tools.FileSystem.FileExists(logdatei) = True Then
		SDB.Tools.FileSystem.DeleteFile(logdatei)
	End If
	WriteLog "Start DiscogsTagger " & VersionStr
	WriteOptions()

End Sub


Sub WriteLog(Text)

	Dim filesys, filetxt, logdatei, tmpText
	'Const ForReading = 1, ForWriting = 2, ForAppending = 8
	logdatei = SDB.ScriptsPath & "Discogs_Script.log"
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set filetxt = filesys.OpenTextFile(logdatei, 8, True)
	tmpText = Time & Chr(9) & SDB.ToAscii(Text)
	filetxt.WriteLine(tmpText)
	filetxt.Close

End Sub



'+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-


Function searchKeyword(Keywords, Role, AlbumRole, artistName)

	WriteLog "Start searchKeyword"
	Dim tmp, x, RE, searchPattern
	tmp = Split(Keywords, ",")
	Set RE = New RegExp
	RE.IgnoreCase = True
	For Each searchPattern In tmp
		WriteLog "searchPattern=" & searchPattern
		If InStr(searchPattern, "*") <> 0 Then
			searchPattern = Replace(searchPattern, "*", ".*")
			RE.Pattern = "^" & searchPattern & "$"
			If RE.Test(Role) Then
				If InStr(AlbumRole, artistName) = 0 Then
					If AlbumRole = "" Then
						AlbumRole = artistName
					Else
						AlbumRole = AlbumRole & Separator & artistName
					End If
					searchKeyword = AlbumRole
				Else
					searchKeyword = "ALREADY_INSIDE_ROLE"
				End If
				Exit For
			End If
		Else
			If Trim(LCase(Role)) = Trim(LCase(searchPattern)) Then
				If InStr(AlbumRole, artistName) = 0 Then
					If AlbumRole = "" Then
						AlbumRole = artistName
					Else
						AlbumRole = AlbumRole & Separator & artistName
					End If
					searchKeyword = AlbumRole
				Else
					searchKeyword = "ALREADY_INSIDE_ROLE"
				End If
				Exit For
			End If
		End If
	Next

End Function

Class VbsJson
	'Author: Demon
	'Date: 2012/5/3
	'Website: http://demon.tw
	Private Whitespace, NumberRegex, StringChunk
	Private b, f, r, n, t

	Private Sub Class_Initialize
		Whitespace = " " & vbTab & vbCr & vbLf
		b = ChrW(8)
		f = vbFormFeed
		r = vbCr
		n = vbLf
		t = vbTab

		Set NumberRegex = New RegExp
		NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
		NumberRegex.Global = False
		NumberRegex.MultiLine = True
		NumberRegex.IgnoreCase = True

		Set StringChunk = New RegExp
		StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
		StringChunk.Global = False
		StringChunk.MultiLine = True
		StringChunk.IgnoreCase = True
	End Sub

	'Return a JSON string representation of a VBScript data structure
	'Supports the following objects and types
	'+-------------------+---------------+
	'| VBScript          | JSON          |
	'+===================+===============+
	'| Dictionary        | object        |
	'+-------------------+---------------+
	'| Array             | array         |
	'+-------------------+---------------+
	'| String            | string        |
	'+-------------------+---------------+
	'| Number            | number        |
	'+-------------------+---------------+
	'| True              | true          |
	'+-------------------+---------------+
	'| False             | false         |
	'+-------------------+---------------+
	'| Null              | null          |
	'+-------------------+---------------+
	Public Function Encode(ByRef obj)
		Dim buf, i, c, g
		Set buf = CreateObject("Scripting.Dictionary")
		Select Case VarType(obj)
			Case vbNull
				buf.Add buf.Count, "null"
			Case vbBoolean
				If obj Then
					buf.Add buf.Count, "true"
				Else
					buf.Add buf.Count, "false"
				End If
			Case vbInteger, vbLong, vbSingle, vbDouble
				buf.Add buf.Count, obj
			Case vbString
				buf.Add buf.Count, """"
				For i = 1 To Len(obj)
					c = Mid(obj, i, 1)
					Select Case c
						Case """" buf.Add buf.Count, "\"""
						Case "\"  buf.Add buf.Count, "\\"
						Case "/"  buf.Add buf.Count, "/"
						Case b    buf.Add buf.Count, "\b"
						Case f    buf.Add buf.Count, "\f"
						Case r    buf.Add buf.Count, "\r"
						Case n    buf.Add buf.Count, "\n"
						Case t    buf.Add buf.Count, "\t"
						Case Else
							If AscW(c) >= 0 And AscW(c) <= 31 Then
								c = Right("0" & Hex(AscW(c)), 2)
								buf.Add buf.Count, "\u00" & c
							Else
								buf.Add buf.Count, c
							End If
					End Select
				Next
				buf.Add buf.Count, """"
			Case vbArray + vbVariant
				g = True
				buf.Add buf.Count, "["
				For Each i In obj
					If g Then g = False Else buf.Add buf.Count, ","
					buf.Add buf.Count, Encode(i)
				Next
				buf.Add buf.Count, "]"
			Case vbObject
				If TypeName(obj) = "Dictionary" Then
					g = True
					buf.Add buf.Count, "{"
					For Each i In obj
						If g Then g = False Else buf.Add buf.Count, ","
						buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
					Next
					buf.Add buf.Count, "}"
				Else
					Err.Raise 8732,,"None dictionary object"
				End If
			Case Else
				buf.Add buf.Count, """" & CStr(obj) & """"
		End Select
		Encode = Join(buf.Items, "")
	End Function

	'Return the VBScript representation of ``str(``
	'Performs the following translations in decoding
	'+---------------+-------------------+
	'| JSON          | VBScript          |
	'+===============+===================+
	'| object        | Dictionary        |
	'+---------------+-------------------+
	'| array         | Array             |
	'+---------------+-------------------+
	'| string        | String            |
	'+---------------+-------------------+
	'| number        | Double            |
	'+---------------+-------------------+
	'| true          | True              |
	'+---------------+-------------------+
	'| false         | False             |
	'+---------------+-------------------+
	'| null          | Null              |
	'+---------------+-------------------+
	Public Function Decode(ByRef str)
		'return base object
		Set Decode = ParseObject(str, 1)
	End Function

	Private Function ParseValue(ByRef str, ByRef idx)
		Dim c, ms

		idx = NextToken(str, idx)
		c = Mid(str, idx, 1)

		If c = "{" Then
			Set ParseValue = ParseObject(str, idx)
			Exit Function
		ElseIf c = "[" Then
			Set ParseValue = ParseArray(str, idx)
			Exit Function
		ElseIf c = """" Then
			idx = idx + 1
			ParseValue = ParseString(str, idx)
			Exit Function
		ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ParseValue = Null
			Exit Function
		ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ParseValue = True
			Exit Function
		ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
			idx = idx + 5
			ParseValue = False
			Exit Function
		Else
			Set ms = NumberRegex.Execute(Mid(str, idx))
			If ms.Count = 1 Then
				idx = idx + ms(0).Length
				SetLocale "en-US"
				ParseValue = CDbl(ms(0))
				SetLocale 0
				Exit Function
			End If
		End If

		Err.Raise 8732,,"No JSON object could be ParseValued"
	End Function

	Private Function ParseObject(ByRef str, ByRef idx)
		Dim c, key, value
		Set ParseObject = CreateObject("Scripting.Dictionary")

		idx = NextToken(str, idx)

		c = Mid(str, idx, 1)

		If c = "{" Then
			idx = NextToken(str,idx+1)
		Else
			Err.Raise 8732,,"Expected { to begin Object"
		End If

		c = Mid(str, idx, 1)

		Do
			If c <> """" And c <> "}" Then

				Err.Raise 8732,,"Expecting property name or } near: " & Mid(str,idx)

			ElseIf c = """" Then

				idx = idx + 1
				key = ParseString(str, idx)

				idx = NextToken(str, idx)
				If Mid(str, idx, 1) <> ":" Then
					Err.Raise 8732,,"Expecting : delimiter near: " & Mid(str,idx)
				End If

				' skip : and whitespace after key
				idx = NextToken(str, idx + 1)

				' check for object or array value
				If Mid(str,idx,1) = "{" Or Mid(str,idx,1) = "[" Then
					Set value = ParseValue(str, idx)
				Else
					value = ParseValue(str,idx)
				End If

				ParseObject.Add key, value
			End If

			c = Mid(str,idx,1)

			If c = "}" Then
				idx = NextToken(str,idx+1)
				Exit Function
			End If

			If c <> "," Then

				Err.Raise 8732,,"Expecting , delimiter near: " & Mid(str,idx)

			End If

			'skip , and whitespace after value
			idx = NextToken(str, idx + 1)
			c = Mid(str, idx, 1)
			If c <> """" Then
				Err.Raise 8732,,"Expecting property name"
			End If
		Loop
	End Function

	Private Function ParseArray(ByRef str, ByRef idx)
		Dim c, values, value
		Set ParseArray = CreateObject("Scripting.Dictionary")

		idx = NextToken(str, idx)
		c = Mid(str, idx, 1)

		If c = "[" Then
			idx = NextToken(str,idx+1)
		Else
			Err.Raise 8732,,"Expected [ to begin Array"
		End If

		Do
			c = Mid(str, idx, 1)

			If c = "]" Then
				idx = NextToken(str,idx+1)
				Exit Function
			End If

			ParseArray.Add ParseArray.Count, ParseValue(str, idx)

			c = Mid(str, idx, 1)

			If c = "]" Then
				idx = NextToken(str, idx+1)
				Exit Function
			End If

			If c <> "," Then
				Err.Raise 8732,,"Expecting , delimiter near: " & Mid(str,idx)
			End If

			idx = NextToken(str,idx+1)

		Loop
	End Function

	Private Function ParseString(ByRef str, ByRef idx)
		Dim chunks, content, terminator, ms, esc, char
		Set chunks = CreateObject("Scripting.Dictionary")

		Do
			Set ms = StringChunk.Execute(Mid(str, idx))
			If ms.Count = 0 Then
				Err.Raise 8732,,"Unterminated string starting"
			End If

			content = ms(0).Submatches(0)
			terminator = ms(0).Submatches(1)
			If Len(content) > 0 Then
				chunks.Add chunks.Count, content
			End If

			idx = idx + ms(0).Length

			If terminator = """" Then
				Exit Do
			ElseIf terminator <> "\" Then
				Err.Raise 8732,,"Invalid control character"
			End If

			esc = Mid(str, idx, 1)

			If esc <> "u" Then
				Select Case esc
					Case """" char = """"
					Case "\"  char = "\"
					Case "/"  char = "/"
					Case "b"  char = b
					Case "f"  char = f
					Case "n"  char = n
					Case "r"  char = r
					Case "t"  char = t
					Case Else Err.Raise 8732,,"Invalid escape"
				End Select
				idx = idx + 1
			Else
				char = ChrW("&H" & Mid(str, idx + 1, 4))
				idx = idx + 5
			End If

			chunks.Add chunks.Count, char
		Loop

		ParseString = Join(chunks.Items, "")
	End Function

	Private Function NextToken(ByRef str, ByVal idx)
		Do While idx <= Len(str) And InStr(Whitespace, Mid(str, idx, 1)) > 0
			idx = idx + 1
		Loop
		NextToken = idx
	End Function

End Class

Function AddToField(ByRef field, ByVal ftext)

	' for adding data to multi-valued fields
	If field = "" Then
		field = ftext
	Else
		field = field & Separator & ftext
	End If

End Function


Function LookForFeaturing(Text)

	Dim i, tmp, x
	tmp = Split(FeaturingKeywords, ",")
	For each x in tmp
		If LCase(Text) = LCase(x) Then
			LookForFeaturing = true
			Exit Function
		End If
	Next
	LookForFeaturing = false

End Function


Function CheckLeadingZeroTrackPosition(TrackPosition)

	Dim tmpSplit, tmpTrack
	If InStr(TrackPosition, "-") <> 0 Then
		tmpSplit = Split(TrackPosition, "-")
		TrackPosition = tmpSplit(1)
	End If
	If InStr(TrackPosition, ".") <> 0 Then
		tmpSplit = Split(TrackPosition, ".")
		TrackPosition = tmpSplit(1)
	End If
	If Left(TrackPosition, 1) = "0" Then
		CheckLeadingZeroTrackPosition = True
	Else
		CheckLeadingZeroTrackPosition = False
	End If

End Function



Function exchange_roman_numbers(Text)

	If Text = "I" Then Text = 1
	If Text = "II" Then Text = 2
	If Text = "III" Then Text = 3
	If Text = "IV" Then Text = 4
	If Text = "V" Then Text = 5
	If Text = "VI" Then Text = 6
	If Text = "VII" Then Text = 7
	If Text = "VIII" Then Text = 8
	If Text = "IX" Then Text = 9
	If Text = "X" Then Text = 10
	If Text = "XI" Then Text = 11
	If Text = "XII" Then Text = 12
	If Text = "XIII" Then Text = 13
	If Text = "XIV" Then Text = 14
	If Text = "XV" Then Text = 15
	If Text = "XVI" Then Text = 16
	If Text = "XVII" Then Text = 17
	If Text = "XVIII" Then Text = 18
	If Text = "XIX" Then Text = 19
	If Text = "XX" Then Text = 20
	exchange_roman_numbers = Text

End Function

Function DecodeHtmlChars(Text)

	DecodeHtmlChars = Text
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&quot;",	"""")
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&lt;",	"<")
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&gt;",	">")
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&amp;",	"&")

End Function


Function EncodeHtmlChars(Text)

	EncodeHtmlChars= Text
	EncodeHtmlChars= Replace(EncodeHtmlChars, "&",	"&amp;")
	EncodeHtmlChars= Replace(EncodeHtmlChars,"""",	"&quot;")
	EncodeHtmlChars= Replace(EncodeHtmlChars,"<",	"&lt;")
	EncodeHtmlChars= Replace(EncodeHtmlChars, ">",	"&gt;")

End Function


Function CleanSearchString(Text)

	CleanSearchString = Text
	CleanSearchString = Replace(CleanSearchString,")", " ") 'remove paranthesis to avoid search errors (discogs bug)
	CleanSearchString = Replace(CleanSearchString,"(", " ") 'also clean other unneccessary characters
	CleanSearchString = Replace(CleanSearchString,"[", " ")
	CleanSearchString = Replace(CleanSearchString,"]", " ")
	CleanSearchString = Replace(CleanSearchString,".", " ")
	CleanSearchString = Replace(CleanSearchString,"@", " ")
	CleanSearchString = Replace(CleanSearchString,"_", " ")
	CleanSearchString = Replace(CleanSearchString,"?", " ")

End Function


Function CleanArtistName(artistname)

	CleanArtistName = DecodeHtmlChars(artistname)
	If InStr(CleanArtistName, " (") > 0 Then CleanArtistName = Left(CleanArtistName, InStrRev(CleanArtistName, " (") - 1)
	If InStr(CleanArtistName, ", The") > 0 Then CleanArtistName = "The " & Left(CleanArtistName, InStrRev(CleanArtistName, ", The") - 1)

End Function


Function AddAlternative(Alternative)

	Dim i
	If Trim(Alternative) <> "" Then
		For i = 0 To AlternativeList.Count - 1
			If AlternativeList.Item(i) = Trim(Alternative) Then
				Exit Function
			End If
		Next
		AlternativeList.Add Trim(Alternative)
	End If

End Function


Function AddAlternatives(Song)

	Dim SavedArtist, SavedTitle, SavedAlbum, SavedAlbumArtist, SavedFolderName, SavedFileName, Custom
	SavedArtist = Song.ArtistName
	SavedTitle = Song.Title
	SavedAlbum = Song.AlbumName
	SavedAlbumArtist = Song.AlbumArtistName
	SavedFolderName = Mid(Song.Path, 1, InStrRev(Song.Path,"\")-1)
	SavedFolderName = Mid(SavedFolderName, InStrRev(SavedFolderName,"\")+1)
	SavedFileName = Mid(Song.Path, 1, InStrRev(Song.Path,".")-1)
	SavedFileName = Mid(SavedFileName, InStrRev(SavedFileName,"\")+1)

	AddAlternative SavedFolderName
	If(InStr(SavedFolderName,"(") > 0) Then
		Custom = Mid(SavedFolderName,1,InStr(SavedFolderName,"(")-1)
		AddAlternative Custom
	End If
	If(InStr(SavedFolderName,"[") > 0) Then
		Custom = Mid(SavedFolderName,1,InStr(SavedFolderName,"[")-1)
		AddAlternative Custom
	End If
	AddAlternative SavedFileName
	If(InStr(SavedFileName,"(") > 0) Then
		Custom = Mid(SavedFileName,1,InStr(SavedFileName,"(")-1)
		AddAlternative Custom
	End If
	If(InStr(SavedFileName,"[") > 0) Then
		Custom = Mid(SavedFileName,1,InStr(SavedFileName,"[")-1)
		AddAlternative Custom
	End If
	AddAlternative Custom
	AddAlternative SavedArtist
	AddAlternative SavedTitle
	AddAlternative SavedAlbum
	AddAlternative SavedAlbumArtist
	If(InStr(SavedTitle,"(") > 0) Then
		Custom = Mid(SavedTitle,1,InStr(SavedTitle,"(")-1)
		AddAlternative Custom
	End If
	If(InStr(SavedTitle,"[") > 0) Then
		Custom = Mid(SavedTitle,1,InStr(SavedTitle,"[")-1)
		AddAlternative Custom
	End If
	AddAlternative SavedArtist & " " & SavedAlbum
	AddAlternative SavedAlbumArtist & " " & SavedAlbum
	AddAlternative SavedArtist & " " & SavedTitle
	AddAlternative SavedAlbumArtist & " " & SavedTitle

End Function


Function IsInteger(Str)

	Dim i, d
	IsInteger = True
	For i = 1 To Len(str)
		d = Mid(str, i, 1)
		If Asc(d) < 48 Or Asc(d) > 57 Then
			IsInteger = False
			Exit For
		End If
	Next

End Function


Function PackSpaces(Text)

	PackSpaces = Text
	PackSpaces = Replace(PackSpaces,"  ", " ") 'pack spaces
	PackSpaces = Replace(PackSpaces,"  ", " ") 'pack spaces left

End Function


Function search_involved(Text, SearchText)

	Dim i
	For i = 1 To UBound(Text)
		If Left(Text(i), Len(SearchText)) = SearchText Then
			search_involved = i
			Exit Function
		End If
	Next
	search_involved = -1

End Function



Function URLEncodeUTF8(ByRef input)

	' urlencode a string with UTF8 encoding - yes, it is cryptic but it works!
	Dim i, result, CurrentChar
	Dim FirstByte, SecondByte, ThirdByte

	result = ""
	For i = 1 To Len(input)
		CurrentChar = Mid(input, i, 1)
		CurrentChar = AscW(CurrentChar)

		If (CurrentChar < 0) Then
			CurrentChar = CurrentChar + 65536
		End If

		If (CurrentChar >= 0) And (CurrentChar < 128) Then
			' 1 byte
			If(CurrentChar = 32) Then
				' replace space with "+"
				result = result & "+"
			Else
				' replace punctuation chars with "%hex"
				result = result & Escape(Chr(CurrentChar))
			End If
		End If

		If (CurrentChar >= 128) And (CurrentChar < 2048) Then
			' 2 bytes
			FirstByte  = &HC0 Xor ((CurrentChar And &HFFFFFFC0) \ &H40&)
			SecondByte = &H80 Xor (CurrentChar And &H3F)
			result = result & "%" & Hex(FirstByte) & "%" & Hex(SecondByte)
		End If

		If (CurrentChar >= 2048) And (CurrentChar < 65536) Then
			' 3 bytes
			FirstByte  = &HE0 Xor (((CurrentChar And &HFFFFF000) \ &H1000&) And &HF)
			SecondByte = &H80 Xor (((CurrentChar And &HFFFFFFC0) \ &H40&) And &H3F)
			ThirdByte  = &H80 Xor (CurrentChar And &H3F)
			result = result & "%" & Hex(FirstByte) & "%" & Hex(SecondByte) & "%" & Hex(ThirdByte)
		End If
	Next
	URLEncodeUTF8 = result

End Function


Sub SwitchAll()

	Dim templateHTMLDoc, i, checkBox
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	Set checkBox = templateHTMLDoc.getElementById("selectall")
	SelectAll = checkBox.Checked

	For i = 0 To iMaxTracks - 1
		If SelectAll Then
			UnselectedTracks(i) = ""
		Else
			UnselectedTracks(i) = "x"
		End If
	Next

	ReloadResults

End Sub


