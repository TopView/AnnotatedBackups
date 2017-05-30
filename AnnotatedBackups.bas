Option Explicit	'BASIC	###### AnnotatedBackups ######

'Editor=Wide load 4:  Set your wide load editor to 4 column tabs, fixed size font.  Suggest Kate (Linux) or Notepad++ (windows).

	Const sProgramsVersion	= "1.5.19"	'AnnotatedBackups current version
	Const sSettingsVersion	= "1"		'AnnotatedBackupsSettings minimum required version


' === Global constants used for MsgBox() ==========================================================
	'Buttons displayed
	Const sbOkOnly				=  0
	Const sbOkCancel			=  1
	Const sbAbortRetryIgnore	=  2
	Const sbYesNoCancel			=  3
	Const sbYesNo				=  4
	Const sbRetryCancel			=  5

	'Icons displayed
	Const sbStop				= 16
	Const sbQuestion			= 32
	Const sbExclamation			= 48

	'Default button
	Const sbDefaultButton1		=128	'first button is default
	Const sbDefaultButton2		=256	'2nd   button is default
	Const sbDefaultButton3		=512	'3nd   button is default

	'Answers returned
	Const sbOK					=  1
	Const sbCancel				=  2
	Const sbAbort				=  3
	Const sbRetry				=  4
	Const sbIgnore				=  5
	Const sbYes					=  6
	Const sbNo					=  7



' === MAIN PROGRAM - CALLED FROM MENU BUTTON ======================================================
Sub AnnotatedBackups()			'was: Sub AnnotatedBackups(Optional oDoc As Object)


	' --- NAME AND PURPOSE -----------------------------------------------------------------------------------------------------
	'
	'	AnnotatedBackups  - On-demand Save and Backup for LibreOffice. 
	'		First check for settings, and install or upgrade if necessary,
	'		Check a few other things.  Also close any Forms or Reports.
	'		Do one or more backups of the current or given file, possibly removing older backups when done.


	' --- DOCUMENTATION INCLUDING SETUP INSTRUCTIONS ---------------------------------------------------------------------------
	'
	'	https://github.com/TopView/AnnotatedBackups


	' --- CREDITS --------------------------------------------------------------------------------------------------------------
	'
	'	Based partially on 'AutomaticBackup' by squenson, later extended by Ratslinger:
	'
	'	  squenson: https://forum.openoffice.org/en/forum/memberlist.php?mode=viewprofile&u=2781&sid=78e2eae7c08fba145326798ec04077b8
	'	  [Basic] Save a document and create a timestamped copy: https://forum.openoffice.org/en/forum/viewtopic.php?f=21&t=23531
	'
	'	  ratslinger: https://ask.libreoffice.org/en/question/88856/suggeston-for-location-of-backup-files/?answer=89030#post-id-89030
	'	  see also: https://ask.libreoffice.org/en/question/75460/libo-515-on-debian-85-writer-close-without-asking-to-save-in-need-of-an-automatic-incremental-saving-function/


	' --- LICENSE - Creative Commons - Attribution-ShareAlike / CC BY-SA -------------------------------------------------------
	'	The previous work that this is based on was not oficially licensed, but still freely offered for use and further development.  Because I was asked
	'	to select a license when uploading this to LibreOffice Wiki, I chose this which seems in keeping with the intent of the the other authors.
	'
	'	This license lets others remix, tweak, and build upon your work even for commercial purposes, as long as they credit you and license 
	'	their new creations under the identical terms. This license is often compared to “copyleft” free and open source software licenses. 
	'	All new works based on yours will carry the same license, so any derivatives will also allow commercial use. This is the license used by 
	'	Wikipedia, and is recommended for materials that would benefit from incorporating content from Wikipedia and similarly licensed projects. 

	'	USE AT YOUR OWN RISK.  No claims or warranties implied or otherwise as to its’ performance or accuracy.
	'	Please send updates, corrections, or suggestions to: EasyTrieve <CustomerService@OutWestClassifieds.org>



	'=== Useful Constants =================================================================================
	'-Base library names:
	Const  sLibraryName		= "Standard"					'Name of settings Library
	Const sProgramsName		= "AnnotatedBackups"			'Name of programs Module
	Const sSettingsName 	= "AnnotatedBackupsSettings"	'Name of settings Module (prefix)

	'-End of line characters, (can't make these Const),  :-( tip: these don't get passed to subs
	Dim CR As String	: CR = chr(10)
	Dim C2 As String	: C2 = chr(10)&chr(10)
	

	' === First get or create settings and if old possibly update =========================================
	'-Search my modules for the highest possible settings file version#
	'	Settings module names over time:
	'		AnnotatedBackupsSettings		version 1 (orignial)
	'		AnnotatedBackupsSettingsV#		versions 2,3,4...
'	mri BasicLibraries							:stop	'Get BOTH LibreOffice Macros & Dialogs AND My Macros & Dialogs
'	mri DialogLibraries							:stop	'(not useful here)
'	mri BasicLibraries.GetByName(sLibraryName)	:stop	'Get Standard libraries
'	mri ThisComponent.BasicLibraries			:stop	'Get only this document's libs
	Dim iVersion 		As Integer	:iVersion	= 0													'Default if no version found
	Dim sElement		As String
	Dim oLib			As Object	:oLib		= BasicLibraries.getByName(sLibraryName)
	For Each sElement In oLib.ElementNames
		if 		left(sElement,len(sSettingsName)	) = sSettingsName 		Then 					'Some settings found..
			If len(sElement) = len(sSettingsName) Then
'				'Copy older named AnnotatedBackupsSettings to newer name AnnotatedBackupsSettingsV1
'				oLib.insertByName(sSettingsName & "V1", oLib.getByName(sSettingsName))
'				oLib.removeByName(sSettingsName)	'works, but crashes, so took out				'Remove older settings (danger)
				iVersion 	= Max(iVersion,1)														'Older style version 1 found
			ElseIf	left(sElement,len(sSettingsName)+1	) = sSettingsName & "V"	Then
				iVersion 	= Max(iVersion,Right(sElement, len(sElement)-(len(sSettingsName)+1) ))	'Newer style version 2,3,4,etc. found
			End If
		End If
	Next sElement
'	MsgBox "Latest version found: " & iVersion :stop


	'-Check sSettingsVersion sanity
	If sSettingsVersion < iVersion Then MsgBox(_
		 "sSettingsVersion(=" & sSettingsVersion & ") should not be less than discovered iVersion(=" & iVersion & ")" & C2 _
		&"(If you wanted to undo a version then you must also delete the version library.)" _
		,sbOkOnly+sbExclamation _
		,"FATAL CONFIGURATION ERROR"):stop


	'-Make newer settings version if older version found
	If iVersion < sSettingsVersion Then 			'No settings found - must create settings

		CreateSettingsModule(oLib, sProgramsName, SettingsName(sSettingsName,sSettingsVersion), sProgramsVersion)

		'for test:
			'iVersion = 0
			'sSettingsVersion = 2
		If MsgBox(	iif(iVersion=0, "A ", "An updated") & " settings module was just created for you."_
			&C2 _
			_
			_
			&		iif(iVersion=0, "", "Tip: Your previous setting are still found here:"_
			&C2 _	
			&		"     My Macros & Dialogs | " & sLibraryName &" | "& SettingsName(sSettingsName,iVersion)_
			&C2 _
					)_
			_
			_
			&		iif(iVersion=0, "Default settings should probably be fine for now.",_
									"Although default settings will work, you might want to first edit "&_
									"your settings before proceeding with the backup, possibly migrating "&_
									"your previous custom settings.")_
			&C2 _
			_
			_
			&		iif(iVersion=0, "", "Your new setting are found here:"_
			&C2 _
			&		"     My Macros & Dialogs | " & sLibraryName & " | " & SettingsName(sSettingsName,sSettingsVersion)_
			&C2 _
					)_
			_
			&		"Click OK to continue with your backup, or"_
			&CR &	"CANCEL to abort this backup, so you can first edit "_
			& 						iif(iVersion=0, "", "or migrate ") & "your settings."_
			_
			,sbOkCancel+sbQuestion+sbDefaultButton1 _
			,iif(iVersion=0, "A ", "A NEWER ") & " SETTINGS MODULE WAS JUST INSTALLED"_
			) = sbCancel _
		Then stop	'setup: 1=Ok/Cancel + 32=question mark + 128=first button is default	results: 2=Cancel
		
	End If


	'-Get settings
'	sSettingsName = sSettingsName & "V" & iVersion
	Dim sPath 		As String		'Relative path from documents to backups.
	Dim iMaxCopies 	As Integer		'Max number of timestamped backup files to be retained (per file).  See GetiMaxCopies.
	Dim sB(200) 	As String		'Array of file types to possibly backup


	'This is to simulate something like this JavaScript syntax:  [iVersion].GETsPath()" which allows variable object notation, and which Basic doesn't support.
	Dim sSettingsVer	As String :sSettingsVer	= sSettingsVersion			'AnnotatedBackupsSettings minimum required version - need a variable for the Select Case!
	Select Case sSettingsVer

		Case 1																'Simple old style moudle name, w/o version suffix
			sPath 		= AnnotatedBackupsSettings.GETsPath()
			iMaxCopies	= AnnotatedBackupsSettings.GETiMaxCopies()
			sB()		= AnnotatedBackupsSettings.GETsB()

		Case 2																'New style module names, w/ version suffix
			sPath 		= AnnotatedBackupsSettingsV2.GETsPath()
			iMaxCopies	= AnnotatedBackupsSettingsV2.GETiMaxCopies()
			sB()		= AnnotatedBackupsSettingsV2.GETsB()

		Case 3
			sPath 		= AnnotatedBackupsSettingsV3.GETsPath()
			iMaxCopies	= AnnotatedBackupsSettingsV3.GETiMaxCopies()
			sB()		= AnnotatedBackupsSettingsV3.GETsB()

		Case 4
			sPath 		= AnnotatedBackupsSettingsV4.GETsPath()
			iMaxCopies	= AnnotatedBackupsSettingsV4.GETiMaxCopies()
			sB()		= AnnotatedBackupsSettingsV4.GETsB()

		'May have to extend this above someday with more Case statements
		
		'Failsafe if version updated w/o extending the above case statements
		Case Is > 4
			MsgBox("OOPS, select statement is too short for iVersion = " & iVersion & "." &C2 &_
					"Increase the number of case statements to fix this." ,sbOkOnly+sbExclamation ,"FATAL ERROR"):stop
			
	End Select
			


	'=== Now do one or more backups of the current or given file, possibly removing older backups =========

	'--- Check for reasonable iMaxCopies ------------------------------------------------
	Dim iMinCopies	As Integer	:iMinCopies	= 10
	Dim iMsgBoxResult		As Integer	:iMsgBoxResult	= sbNo
	If iMaxCopies < iMinCopies Then iMsgBoxResult = MsgBox(_
				"iMaxCopies(" & iMaxCopies & ") is lower than iMinCopies (" & iMinCopies & "). Ignore?"_
	 	&C2 &	"YES: Ignore and keep ALL backups."_
	 	&C2 &	"NO: Use iMaxCopies as is, (this is for testing only)"_
		&C2 &	"CANCEL to stop, so you can update iMaxCopies in:"_
		&CR &	"    "  &SettingsName(sSettingsName,iVersion)_
		,sbYesNoCancel+sbExclamation+sbDefaultButton1 _
		,"SETUP WARNING: iMaxCopies IS UNEXPECTEDLY LOW")
	If iMsgBoxResult=sbCancel Then stop
	'=sbNo if ok to use iMaxCopies
'MsgBox("answer" & iMsgBoxResult): stop


	'--- Check for non-empty backup path ------------------------------------------------
	if sPath = "" Then MsgBox(_
				"sPath="""""_
		&C2 &	"An empty path might cause accidental deletion of your document."_
		&C2 &	"Fix in this Module:"_
		&C2 &	"   My Macros & Dialogs"_
		&CR	&	"       Standard"_
		&CR	&	"           AnnotatedBackups"_
		&CR	&	"               AnnotatedBackupsSettings"_
		,sbOkOnly+sbExclamation _
		,"FATAL SETUP ERROR"):stop


	'--- Check that no slashes in backup path -------------------------------------------
	if instr(sPath,"/") + instr(sPath,"\") Then MsgBox(_
				"sPath=""" & sPath & """"_
		&C2 &	"This relative path should not contain a / or \ (slash or backslash)."_
		,sbOkOnly+sbExclamation _
		,"FATAL SETUP ERROR"):stop


	'--- Get optional comment and honor abort request -----------------------------------
	'(do this early, so as not to save or close anything if canceled)															'
	Dim sComment	As String
	sComment = InputBox(_
				"(MaxCopies=" & iMaxCopies _
		& 		"; rel. backup path=./" & sPath _
		& 		"; v" & sProgramsVersion _
		& 		", settings v" & sSettingsVersion &")"_
		_
		&C2 &	"Optional backup description, (this gets appended to backup's filename):"_
		,"OPTIONAL FILENAME ANNOTATION"_
		,"none")		'"none" is needed because an empty string and a cancel button are the same thing.
	If sComment = "" Then Exit Sub

	sComment = iif(sComment = "none", "", " " & sComment)		'Remove word 'none', and add leading space to comments


	'--- Get document -------------------------------------------------------------------															'
	On Error GoTo URL_error
'		If Not isObject(ThisComponent) Then MsgBox("ERROR: missing ThisComponent"):stop		'??Strange error
		Dim oDoc 		As Object	:oDoc		= ThisComponent		'Was: "If IsMissing(oDoc) Then Dim oDoc As Object :oDoc = ThisComponent" But, not sure what the If was for as it doesn't work.
		DIM sUrl_From	AS STRING	:sUrl_From	= oDoc.URL			'Default From URL
	On Error GoTo 0


	'--- Adjust oDoc if called from a Base Form, and also close any open Base Forms and Reports, but without chaning oDoc used by other LO modules
'	MsgBox "Title: " & oDoc.Title &CR & "URL: " & oDoc.URL &CR & "db: " & oDoc.supportsService("com.sun.star.sdb.OfficeDatabaseDocument") & "    parent: " & iif(isnull(oDoc.parent),"null","not"): stop
	'app	form?	state	title										db	parent	url
	'------	-------	-------	-------------------------------------------	---	-------	---------------------------------------------------------------------------
	'base			saved	Lookup.odb									t	missing	file:///home/howard/Shared/Data/LO/odb/1.8.0/Demonstrations/Lookup/Lookup.odb
	'base			saved	New Database.odb							t	missing	file:///home/howard/Documents/New%20Database.odb

	'base	form			New Database.odb : Form1					f	t		""													'forms don't have URL's
	'base	form			Lookup.odb : Sample search and edit form	f	t		""													'forms don't have URL's

	'calc 			unsaved	Untitled 2									f	void	""													'unsaved: no url; & no filename extension!
	'				saved	Untitled 2.ods								f	void	file:///home/howard/Documents/Untitled%202.ods

	'draw			unsaved	Untitled 2									f	void	""
	'				saved	Untitled 2.odg								f	void	file:///home/howard/Documents/Untitled%202.odg

	'impress		unsaved	Untitled 2									f	void	""
	'				saved	Untitled 2.odp								f	void	file:///home/howard/Documents/Untitled%202.odp

	'math			unsaved	Untitled 2									f	void	""
	'				saved	Untitled 2.odf								f	void	file:///home/howard/Documents/Untitled%202.odf

	'writer			unsaved	Untitled 2									f	void	""
	'				saved	Untitled 2.odt								f	void	file:///home/howard/Documents/Untitled%202.odt



	'Get out of any Base Form, and make sure all Base forms are closed, but without messing up oDoc.URL for other LO Modules
	if oDoc.supportsService("com.sun.star.sdb.OfficeDatabaseDocument") then	'If Base (main/outer dialog)	(Note: parent doesn't always exist to test, like in Base it isn't there).
			iBaseFormsClosed(oDoc)											'so make sure that all forms are closed

	Else 																	' other: A Base Form, or another LO module, i.e. Calc, Draw, Impress, Math, or Writer
		If not isnull(oDoc.parent) then 									'If a Base form?
			'Unravel - allow this to be run from within a Base Form		-Note!  A new, non-base docuemnt's URL is also empty.
			DO WHILE sUrl_From = ""		: oDoc = oDoc.Parent	:sUrl_From	= oDoc.URL	:LOOP
			iBaseFormsClosed(oDoc)											'so make sure that all forms are closed
		End If
	End If



	'--- Make sure document is saved before proceeding ----------------------------------															'
	' If a new document is open (i.e. unsaved), then first save it so we can get its path and filter type
	If Not(oDoc.hasLocation) Then
		Dim oDocNew		As Object	:oDocNew 	 = oDoc.CurrentController.Frame
		Dim oDispatcher As Object	:oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
		oDispatcher.executeDispatch(oDocNew, ".uno:SaveAs", "", 0, array())		'Create a new file
	End If

	If Not(oDoc.hasLocation) Then MsgBox(_
				"Your document was not saved, so backup was not done. "_
		&C2 &	"(Perhaps you tried saving to an unwritable folder or to a file already in use.)"_
		,sbOkOnly+sbExclamation _
		,"FATAL ERROR - DOCUMENT WAS NOT SAVED"):stop



	' --- Set up pathnames (document to backup and place to back it up) -----------------

	' - Get document to backup's name w/ full path
	' Retrieve the document name and path from the URL
	Dim sDocURL					As String	:sDocURL 				= oDoc.getLocation()		'used once below
	Dim sDocNameWithFullPath 	As String	:sDocNameWithFullPath	= ConvertFromURL(sDocURL)	'Source path/filename


	' --- Detect from path if we have "/" (Linux) or "\" (Windows) ----------------------
	Dim sSlash 		As String
	Dim sOtherSlash As String
	If Instr(1, sDocNameWithFullPath, "/") > 0 _
		Then :sSlash = "/"	:sOtherSlash = "\"		'Linux
		Else :sSlash = "\"	:sOtherSlash = "/"		'Windows
	End If


	' --- extract filename --------------------------------------------------------------
	Dim sDocName	As String	:sDocName = GetFileName(sDocNameWithFullPath, sSlash)				'Source filename


	' --- Backup folder -----------------------------------------------------------------
	' sPath is relative.
	'	"foo	relative - put backups where document is stored		.../basedir/foo/.
	'	"/foo	relative - put backups where document is stored		.../basedir/foo/.		Note: Leading slash is ok too
	Dim i			As Integer	:i			=  Len(sDocNameWithFullPath)
	'note: Star Basic does not have a instrrev(), so this is the workaround:
	While 									   Mid(sDocNameWithFullPath, i,  1) <> sSlash	:i=i-1	:Wend	'strip off doc filename to get abs path
	Dim sAbsPath	As String	:sAbsPath	= Left(sDocNameWithFullPath, i) & sPath & sSlash & sDocName & sSlash	'/DocumentPath/sPath/DocName/

	
	' --- Save current document changes -------------------------------------------------
	' Save the current document only if it has changed, is not new (has been saved before) and is not read only
	If oDoc.isModified and oDoc.hasLocation and Not(oDoc.isReadOnly) Then oDoc.store()


	' --- get timestamp -----------------------------------------------------------------
	'  the timestamp, so it will be identical for all the backup copies
	Dim s_	As String	:s_ = iif(sSlash = "/",":","_")		'Allowable hour:time:seconds delimiter for linux or windows
	Dim sTimeStamp 	As String: sTimeStamp	= "--" & 	Format(Year(	Now), "0000-"	) & _
														Format(Month(	Now), "00-"		) & _
														Format(Day(		Now), "00\_"	) & _
														Format(Hour(	Now), "00"&s_	) & _
														Format(Minute(	Now), "00"&s_	) & _
														Format(Second(	Now), "00"		)


	' --- Change illegal file name characters to dashes in comment ----------------------
	'(Can't abort now because already passed cancel button, but didn't have filename when we had to ask to proceed)
	If sComment<>"" Then
	
		Dim sIllegal()	As String
		If sSlash="/"	Then :sIllegal() = Array("/")											'linux
						Else :sIllegal() = Array("/", "\", ":", "*", "?", """", "<", ">", "|")	'Windows
		End If

		Dim sChar		As String
		For i=1 to Len(sComment)
			For Each sChar in sIllegal()
				if Mid(sComment, i, 1) = sChar Then sComment = Left(sComment,i-1) & "-" & Right(sComment,len(sComment)-i
			Next sChar
		Next i
		
	End If


	' --- do other backups --------------------------------------------------------------
	' For each file filter, let's see whether we should create a backup copy or not
	Dim sBackupName		As String
	
	Dim sDocType 		As String
	Dim sExt			As String	'file name extension
	Dim sSaveToURL		As String
	
	i = 1
	While sB(i) <> ""
		If 																				GetField(sB(i), "|", 1) = "BACKUP" Then	'Future: replace GetField -> split the line once into a array

			sDocType 			= 														GetField(sB(i), "|", 2)
			If  _
				(sDocType = "Base"		And oDoc.supportsService("com.sun.star.sdb.OfficeDatabaseDocument"			)) Or _
				(sDocType = "Calc"		And oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument"			)) Or _
				(sDocType = "Draw"		And oDoc.supportsService("com.sun.star.drawing.DrawingDocument"				)) Or _
				(sDocType = "Impress"	And oDoc.supportsService("com.sun.star.presentation.PresentationDocument"	)) Or _
				(sDocType = "Math"		And oDoc.supportsService("com.sun.star.formula.FormulaProperties"			)) Or _
				(sDocType = "Writer"	And oDoc.supportsService("com.sun.star.text.TextDocument"					)) 	  _
			Then																										'??Think this is right enough for formula.*

				sExt 			= 														GetField(sB(i), "|", 3)			'file name extension (used 2 places)
				
				' --- Check if the backup folder exists, if not we create it ------------------------
				On Error Resume Next
				MkDir 									sAbsPath														'Create directory (if not already found)
				On Error Goto 0

					'Backup name format: name.ext timestamp comment.ext		(Note: .ext twice)			'used once below
					sBackupName	= sDocName 				& "." & sExt _
								& sTimeStamp & sComment & "." & sExt		

					sSaveToURL	= ConvertToURL(sAbsPath & sBackupName)									'Name to save to (used once below)

					On Error Goto StoreToURLError	'Next line fails if original doc is *.xlsx.  To fix first save doc as *.ods.
					oDoc.storeToUrl(sSaveToURL, Array(MakePropertyValue( "FilterName", 	GetField(sB(i), "|", 5) ) ) )	'Now run the filter to write out the file

					On Error Goto 0
					RenameOlderBackups(					sAbsPath, sDocName, sExt, oDoc)
					PruneBackupsToMaxSize(iMaxCopies, 	sAbsPath, sDocName, sExt, iMsgBoxResult)	'And finally possibly remove older backups to limit number of them kept

			End If
		End If
		i = i + 1
	Wend
Exit Sub

'this needs further testing
StoreToURLError:
	MsgBox(Error$_
		&C2 &	"Perhaps your original file was not in an ODF (OpenOffice Document Format), e.g. it might have been in .xlsx"_
		,sbOkOnly+sbExclamation ,"SAVE FAILED")
	stop	

URL_error:
	MsgBox(Error$_
		&C2 &	"This happens during development or testing with the macro editor.  When another LO document is opened, "_
		&		"then closed, then ThisComponent for whatever reason no longer points to anything (even thought we still have another "_
		&		"document open (i.e. the original document we had open for backup testing)."_
		_
		&C2 &	"To fix it just set the focus on the document to backup, then back here in the basic editor, and re-run this."_
		,sbOkOnly+sbExclamation ,ucase("ThisComponent.URL is missing"))
	stop	

End Sub



'##################################################################################################
'### SUBS AND FUNCTIONS USED ABOVE ################################################################
'##################################################################################################

' === Create setttings module name ================================================================
Function SettingsName(sSettingsName As String, mVer As Variant) As String	'Helps deals with transition from old to new style names
	SettingsName = sSettingsName & iif(mVer=1, "", "V" & mVer)				'Version 1 name: 'foo', Version 2-n names: 'fooV2', 'fooV3', etc
End Function



'=== Close any Base forms (prompting user if necessary) ===========================================
Private Sub iBaseFormsClosed(oDoc As Object)
	'-End of line characters, (can't make these Const),  :-( tip: these don't get passed to subs
	Dim C2 As String	: C2 = chr(10)&chr(10)

	'--- Count design mode: Tables, Queries, Forms & Reports, (i.e. they are inside Frames of Frames in the Desktop)
	
	'Table	design:*	ToGet.odb : links.address counties, hi - LibreOffice Base: Table Design					 - LibreOffice Base: 
	'Table	view  :		links.address cities, hi - ToGet - LibreOffice Base: Table Data View					 - LibreOffice Base: 
	'
	'Query	design:* 	ToGet.odb : Items to get query - LibreOffice Base: Query Design							 - LibreOffice Base: 
	'Query	view  :		Items to get query - ToGet - LibreOffice Base: Table Data View							 - LibreOffice Base: 
	'
	'Form	design:*	ToGet.odb : links.items to get - LibreOffice Base: Database Form						 - LibreOffice Base: 
	'Form	view  :*	ToGet.odb : links.items to get - LibreOffice Base: Database Form						 - LibreOffice Base: (same as design!)
	'
	'Report design:*	ToGet.odb : Items to get ok by store report - LibreOffice Base: Oracle Report Builder	 - LibreOffice Base: 
	'Report view  :		(not under this parent)
	'
	'Truths:
	'	If "ToGet.odb : " prefix then we need to save this
	'	If " - LibreOffice Base: Table Design" 			suffix then it's a table 	in design
	'	if " - LibreOffice Base: Query Design" 			suffix then it's a query 	in design
	'	if " - LibreOffice Base: Database Form" 		suffix then it's a form 	in design or view
	'	if " - LibreOffice Base: Oracle Report Builder" suffix then it's a report 	in design


	Dim iOpenTables		As Integer	:iOpenTables	= 0		'Can't auto-close these (not yet at least)
	Dim iOpenQueries	As Integer	:iOpenQueries	= 0		'Can't auto-close these (not yet at least)
	Dim iOpenForms		As Integer	:iOpenForms		= 0		'Can auto-close these
	Dim iOpenReports	As Integer	:iOpenReports	= 0		'Can auto-close these
	
	Dim iFrame			As Integer	'Frame index
	Dim iFrameSub		As Integer	'Sub frame index

	Dim oFrames			As Object	: oFrames = StarDesktop.Frames	'All  		frames
	Dim oFrame			As Object									'Each		frame
	Dim oBaseFrames		As Object									'All  base	frames
	Dim oBaseFrame		As Object									'Each base	frame

	Dim sTxt			As String									: sTxt = "Me: " & oDoc.Title & chr(10)& chr(10)	& "Parents / Children:" & chr(10) 'place to compile output
	
	For iFrame=0 To oFrames.Count-1 Step 1

		'Looking for titles like: "Lookup5.odb - LibreOffice Base"	(First find our parent and ignore all other parents)
		oFrame = oFrames.getByIndex(iFrame)						: sTxt = sTxt & chr(10) & "--" & oFrame.Title & chr(10)
		
'		If instr(oFrame.Title, oDoc.Title & " - LibreOffice Base")<>0 Then
		If isSuffix(oFrame.Title, oDoc.Title & " - LibreOffice Base") Then

			oBaseFrames = oFrame.Frames
			For iFrameSub=0 To oBaseFrames.Count-1 Step 1
			
				oBaseFrame = oBaseFrames.getByIndex(iFrameSub)	: sTxt = sTxt & "      " & oBaseFrame.Title & chr(10)
				
				Select Case True
					Case True XOR NOT (isSuffix(oBaseFrame.Title," - LibreOffice Base: Table Design"			))
						iOpenTables		= 1+iOpenTables
					Case True XOR NOT (isSuffix(oBaseFrame.Title," - LibreOffice Base: Query Design"			))
						iOpenQueries	= 1+iOpenQueries
					Case True XOR NOT (isSuffix(oBaseFrame.Title," - LibreOffice Base: Database Form"			))
						iOpenForms		= 1+iOpenForms
					Case True XOR NOT (isSuffix(oBaseFrame.Title," - LibreOffice Base: Oracle Report Builder"	))
						iOpenReports	= 1+iOpenReports
					Case Else
						MsgBox("Title error: " & oBaseFrame.Title):stop
				End Select

			Next iFrameSub
		End if
	Next iFrame
	

'	MsgBox sTxt & chr(10) & "iOpenTables "		& iOpenTables _
'				& chr(10) & "iOpenQueries "		& iOpenQueries _
'				& chr(10) & "iOpenForms "		& iOpenForms _
'				& chr(10) & "iOpenReports "		& iOpenReports	_
'				': stop


	'--- Now if any forms or reports are open (with possibly unsaved edits!) then ask to close them, or abort the backup.  
	'		(Because I can't figure out how to save any current records changes before the backup).
	Dim i			As Integer	'documents index
	if iOpenForms+iOpenReports Then

		If 	MsgBox(	_
			 iOpenForms 	& " form" 	& iif(iOpenForms  =1,"","s") & " and "_
			&iOpenReports	& " report" & iif(iOpenReports=1,"","s")_
		 	& iif(iOpenForms+iOpenReports=1," is open and needs"," are open and need")_
			&		" to be closed before backup."_
			&C2	&	"Ok to close " _
			& iif(iOpenForms+iOpenReports=1,"it","them") & " now?" _
			, sbYesNo + sbQuestion + sbDefaultButton1 _
			,"Preparing to backup") = sbNo Then stop

		'-Close all my forms    (open or not).  This is harmless as some of them might already be closed, but I can't tell here which ones.
		Dim oForms 	As Object		:oForms		= oDoc.FormDocuments
		If oForms.count Then 
			For i=0 To oForms.count-1

				oForms.getByIndex(i).close
			Next i
		End If
		
		'-Close all my reports (open or not).  This is harmless as some of them might already be closed, but I can't tell here which ones.
		Dim oReports 	As Object	:oReports	= oDoc.ReportDocuments
		If oReports.count Then 
			For i=0 To oReports.count-1
				oReports.getByIndex(i).close
			Next i
		End If
	End If
	
	
	'I don't yet know how to close Table or Query windows (like w/ Forms above, so for now this warning will have to do)
	if iOpenTables+iOpenQueries Then
		If 	MsgBox(	_
			 iOpenTables & " Table"  & iif(iOpenTables	=1,""	,"s"	) & " and "_
			&iOpenQueries & " Quer"  & iif(iOpenQueries	=1,"y"	,"ies"	) & iif(iOpenTables+iOpenQueries=1," is"," are")_
			&		" open & may contain unsaved records."_
			&C2	&	"Ok to Ignore? Or Cancel to be safe, manually close, & retry?"						 	 _
			,sbOkCancel+sbQuestion+sbDefaultButton2 _
			,"CAUTION!") 			= sbNo Then stop
			
		'-Close all my Tables & Queries (open or not).  This is harmless as some of them might already be closed, but I can't tell here which ones.
		
	End If

End Sub


'--- Test string for a suffix
Private Sub isSuffix(s1 As String, s2 As String) As Boolean	'Look to see if s2 is suffix in s1
'	msgbox Right(s1,len(s2))
	isSuffix = (Right(s1,len(s2)) = s2)
End Sub
'Private Sub isSuffixTest()
'	MsgBox isSuffix("this is a test" , "is a test")	'true
'	MsgBox isSuffix("this is a test2", "is a test")	'false
'End Sub



'=== Return filename /wo path or ext ==============================================================
Private Sub GetFileName(byVal sNameWithFullPath as String, byval sSlash as String) as String

	Dim i as Long
	Dim j as Long

	GetFileName = ""

	'Search from the end of the full name.  "." will indicate the end of the file name and the beginning of the extension
	For i = Len(sNameWithFullPath) To 1 Step -1
		If Mid(sNameWithFullPath, i, 1) = "." Then

			' We have found a ".", so now we continue backwards and search for the path delimiter "\" or "/"
			For j = i - 1 to 1 Step -1
				If Mid(sNameWithFullPath, j, 1) = sSlash Then

					' We have found it, the file name is the
					' piece of string between the two
					GetFileName = Mid(sNameWithFullPath, j + 1, i - j - 1)
					j = 0
					i = 0
				End If
			Next j
		End If
	Next i
   
End Sub



'=== Return nth text field =========================================================================
' e.g. n=2 from "xxx|yyy|foo", gives "yyy".  (n=0 ok, but returns nothing)
Private Sub GetField(byVal sInput As String, byVal sDelimiter As String, ByVal n As Integer) 	As String

	sInput = sDelimiter & sInput & sDelimiter							'To simplify searching sandwitch the input in outer delimiters
	GetField = ""														'Default output if field or is empty found

	Dim iStart	As Long	:iStart = 1										'Char position after nth delimiter

	'A) Find the character position of the nth delimiter
	Dim i 		As Integer
	For i = 1 to n
		iStart = InStr(iStart, sInput, sDelimiter)+1					'Char position after ith delimiter
		If iStart = 1 			Then Exit Sub							'If search fails, i.e. ran out of delimiters too soon			, then silently return an empty string
	Next i
	If 	iStart = Len(sInput)+1 	Then Exit Sub							'If search fails, i.e. found nth delimiter, but its the last one, then silently return an empty string

	'B) Find the character position before the next delimiter
	Dim iEnd	As Long	:iEnd = InStr(iStart, sInput, sDelimiter)		'Char position at    found delimiter n+1

	'C) Return the portion of string between the two delimiters
	GetField = RTrim(Mid(sInput, iStart, iEnd - iStart))				'Input, Start, Length	(ignore extra delimiters we put on)

End Sub


''--Tests to make sure code above is working (uncomment and single step through)
'Sub test_GetField()	
'  dim s as string
'
'  'These should all return "", except as noted
'  s = getfield(""		, "|",-1)
'  s = getfield(""		, "|", 0)
'  s = getfield(""		, "|", 1)
'  s = getfield(""		, "|", 2)
'  s = getfield(""		, "|", 3)
'
'  s = getfield("x"		, "|",-1)
'  s = getfield("x"		, "|", 0)
'  s = getfield("x"		, "|", 1)			'should be "x"
'  s = getfield("x"		, "|", 2)
'  s = getfield("x"		, "|", 3)
'
'  s = getfield("x|y|z"	, "|",-1)
'  s = getfield("x|y|z"	, "|", 0)
'  s = getfield("x|y|z"	, "|", 1)			'should be "x"
'  s = getfield("x|y|z"	, "|", 2)			'should be "y"
'  s = getfield("x|y|z"	, "|", 3)			'should be "z"
'  s = getfield("x|y|z"	, "|", 4)
'  s = getfield("x|y|z"	, "|", 5)
'
'  s = getfield("x 	 "	, "|", 1)			'should be "x" (should remove mixed trailing tabs and spaces)
'  dim i as integer: i =len(s)				'should be 1
'end Sub



'=== Returns file filter type =====================================================================
' 	Credit: http://www.oooforum.org/forum/viewtopic.phtml?t=52047
Private Sub GetFilterType(byVal sFileName as String) as String

	'Get access to UNO methods ("services")
	Dim oSFA 			As Object	:oSFA		= createUNOService("com.sun.star.ucb.SimpleFileAccess"		)
	Dim oTD 			As Object	:oTD		= createUnoService("com.sun.star.document.TypeDetection"	)
	
	'Open given filename for reading
	Dim oInpStream 		As Object	:oInpStream	= oSFA.openFileRead(ConvertToUrl(sFileName))			'open given filenmae using ucb.SimpleFileAccess 

		'Get it's Type
		GetFilterType	= oTD.queryTypeByDescriptor(MakePropertyValue("InputStream",oInpStream), true)	'queryTypeByDescriptor

	oInpStream.closeInput()																				'close

End Sub


'=== Create and return a new com.sun.star.beans.PropertyValue =====================================
Private Sub MakePropertyValue(Optional sName As String, Optional sValue As Variant) As com.sun.star.beans.PropertyValue
    Dim oPropertyValue As New com.sun.star.beans.PropertyValue

    If Not IsMissing(sName	) Then oPropertyValue.Name 	= sName   
    If Not IsMissing(sValue	) Then oPropertyValue.Value = sValue

    MakePropertyValue() = oPropertyValue
End Sub



' === look for older style names and rename files if found ==============================
'older naming style:	ToGet--2017-05-07_22:20:16.odb	
'newer naming style:	ToGet.odb--2017-05-07_22:20:16.odb		(inserted extra .odb to make it easier to un-timestamp name)
Private Sub RenameOlderBackups(sAbsPath As String, sDocName As String, sExt As String, oDoc As Object)
	Dim mOlder() 		As String											'Array to store list of existing older backup path/file names
	Dim iOlder	 		As Integer 	:iOlder 		= 0						'Count of existing backup files	
	Dim stFileName 		As String
		
	'Get list of older style named backups	
	stFileName 	= Dir(sAbsPath, 0)		'Get FIRST normal file from pathname
	Do While (stFileName <> "")
	
		'Huristic to test for older style backup name: 	sDocName(no ext) & -- * sExt		where * is:  timestamp comment
		If _
			InStr(stFileName, sDocName & "--") And		_
			Right(stFileName,3				) = sExt	_
			Then	:ReDIM Preserve mOlder(iOlder)		:mOlder(iOlder) = stFileName	:iOlder = iOlder+1 	'get list of existing backups
		End if
									 stFileName 	= Dir()					'Get NEXT  normal file from pathname as initially used above
	Loop
	
	'Now rename files in the list to new style names
	Dim sNewName		As String
	For Each stFileName In mOlder()
		sNewName = Left(stFileName,len(sDocName)) & "." & sExt & right(stFileName,len(stFileName)-len(sDocName))

		Name sAbsPath & stFileName As sAbsPath & sNewName	'rename file:	Name OldName As NewName
	Next stFileName
End Sub



' === possibly remove older backups =====================================================
Private Sub PruneBackupsToMaxSize(iMaxCopies As Integer, sAbsPath As String, sDocName As String, sExt As String, iMsgBoxResult As Integer)
	if iMaxCopies = 0 then exit sub											'If iMaxCopies is = 0, there is no need to read, sort or delete any files.

		
	' --- First get list of existing backups --------------------------------------------
	Dim mArray() 		As String											'Array to store list of existing backup path/file names
	Dim iBackups 		As Integer 	:iBackups 		= 0						'Count of existing backup files	
	Dim stFileName 		As String	:stFileName 	= Dir(sAbsPath, 0)		'Get FIRST normal file from pathname
	Do While (stFileName <> "")
	
		'Huristic to test for deletable backups, finds: 	sDocName & -- * sExt		where * is:  timestamp comment
		If _
			Left(stFileName,Len(sDocName)+2	) = sDocName & "--" And _
		   Right(stFileName,3				) = sExt 				_
		   Then	:ReDIM Preserve mArray(iBackups)	:mArray(iBackups) = stFileName	:iBackups = iBackups+1 	'get list of existing backups
		End if
									 stFileName 	= Dir()					'Get NEXT  normal file from pathname as initially used above
	Loop


	'--- iMaxCopies < iMinCopies AND test mode: don't purge files, only report results --
	If iMsgBoxResult = sbYes Then msgbox("New backup saved, but didn't purge any older backups.  "_
			&C2 &	iBackups & " backups found.  iMaxCopies limit set to " & iMaxCopies & " backups."_
			,,"RESULTS") : stop 


	'--- Warn before deleting more than one backup---------------------------------------
	'Failsave check: This is incase iMaxCopies is reduced for testing, or other unforseen bug occurs.
	Dim iKill As Integer	:iKill = iBackups - iMaxCopies								'# of old backups to delete
	If iKill > 1 Then iMsgBoxResult = MsgBox(_
					"Only purge the oldest backup?  (No to Purge " & iKill & " older backups.)"_
			&C2 &	"After a backup in order to limit the total number of backups saved, normally "_
			&		"the oldest backup might be removed. But perhaps you recently decreased "_
			&		"iMaxCopies which could trigger this question." _
			,sbYesNo+sbExclamation+sbDefaultButton1 _
			,"UNEXPECTED FILE DELETION REQUEST")
	If iMsgBoxResult = sbYes Then iKill=1


	'--- Deleting oldest files ----------------------------------------------------------
	'Deletes oldest files exceeding the limit set in iMaxCopies
	iSort(mArray)																	'Sort list of existing backups (by timestamp, oldest first)
	Dim i As Integer  :For i = 0 to iKill -1: Kill(sAbsPath & mArray(i)): Next i	'now delete oldest ones as necessary

End Sub



'=== insertion sort (oldest first) ================================================================
Private Sub iSort(mArray)
	Dim Lb 	as integer	:Lb = lBound(mArray)	'lower array bound
	Dim Ub 	as integer	:Ub = uBound(mArray) 	'upper array bound
	
	Dim iT	As Long		'element under 	Test	, Array index	- What we are looking to possibly move and insert into lower already sorted stuff
	Dim sT 	as string 	'element under 	Test	, Element value	- Variable to hold what we are testing, so cell can get stomped on and not lost by stuff shifting up

	Dim iC	as Long		'element to		Compare	, Array index	- Index to search thru what is already sorted, to find what might be bigger than sT 

		
	for iT = Lb+1 to Ub											'Work forwards through array: from second element to last element
		sT = mArray(iT)												'Save element to test and possibly to move down (because will possibly get stomped on).

		For iC = iT-1 to Lb step -1									'Search backwards thru what's already sorted until we're less than what we are finding.
			If strComp(mArray(iC), sT, 0) < 1 Then  Exit For		'strComp returns -1 when mArray(iC) < t; Exit loop because we found insertion place
			mArray(iC+1) = mArray(iC)								'otherwise shift elements up 1 and step down and repeat the test
		Next iC

		mArray(iC+1) = sT											'Finally, insert moved element here (might even be the very first position)
	Next iT
End Sub


''--Test of iSort()
'Sub test_iSort()
'	Dim mArrayA(2) As String
'	mArrayA(0) = "x3"
'	mArrayA(1) = "x2"
'	mArrayA(2) = "x1"

'	iSort(mArrayA)

'	Dim x0 As String	:x0 = mArrayA(0)
'	Dim x1 As String	:x1 = mArrayA(1)
'	Dim x2 As String	:x2 = mArrayA(2)
'End Sub



' === Trim spaces and tabs from right end =========================================================
Private Sub RTrim(str As String) As String
	RTrim=str																	'simplify code; use returned value as working string
	Dim i as Long																'character counter (from end to start)
	For i = Len(RTrim) to 1 step -1
		If right(RTrim,1) <> chr(9) and right(RTrim,1) <> " " Then Exit Sub		'if trailing white space not found we're done
		RTrim = left(RTrim,len(RTrim)-1)										'otherwise remove trailing white space, step left, repeat
	Next
End Sub



' === Insert substring in string ==================================================================
'-Insert substring within string after delimiter
Function InsertAtDelimiter (sStr As String,  sSubStr As String, sDelimiter As String) As String
	InsertAtDelimiter = InsertSubString(sStr, sSubStr, InStr(1,sStr,sDelimiter) + len(sDelimiter))
End Function

'-Insert substring within string at position
Function InsertSubString (Str As String, sSubStr As String, iPosition As Long) As String
	InsertSubString = Left(Str,iPosition) & sSubStr & Right(Str,len(Str)-iPosition+1)
End Function



' === Math functions ==============================================================================
Function Max (x As Long, y as Long) As Long
	Max = IIf(x > y, x, y)
End Function



' === Copy code below and use it to create a new settings module of the given version =============
Sub CreateSettingsModule(oLib As Object, sProgramsName As String, sSettingsNameV As String, sProgramsVersion As String)
	'-End of line characters, (can't make these Const),  :-( tip: these don't get passed to subs
	Dim CR As String	: CR = chr(10)
	Dim C2 As String	: C2 = chr(10)&chr(10)

'	mri oLib:stop

	'Extract this file's code (Basic) in text
	Dim sDelimiter		As String	:sDelimiter			= "'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"&CR
	Dim sProgramsSource	As String	:sProgramsSource	= oLib.getByName(sProgramsName)					'text of this entire file
	Dim sSettingsSource	As String	:sSettingsSource	= _
		"Option Explicit	'BASIC	###### " & sSettingsNameV & " for caller V " & sProgramsVersion & " ######" &C2 &_
		Right(sProgramsSource,len(sProgramsSource)-instr(1,sProgramsSource,sDelimiter)-len(sDelimiter))	'text of just template below

	oLib.insertByName(sSettingsNameV, sSettingsSource)	'Copy template below into new settings module
End Sub



'##################################################################################################
'### SETTINGS MODULE TEMPLATE #####################################################################
'##################################################################################################
'
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'NORMALLY DON'T EDIT THE CODE BELOW
'  IT IS A TEMPLATE USED TO CREATE THE SETTINGS MODULES, E.G. AnnotatedBackupsSettings
'		You can edit this code below, but it will get overwritten if you update your software
'		The code below is not used to determine your settings.
'
'		To edit your settings edit the Module created when you first run the backup.  It's named 
'		AnnotatedBackupsSettings for version 1, AnnotatedBackupsSettingsV2, for version 2, etc.
'
'  THIS NEXT LINE IS A delimiter used above - DON'T EDIT IT
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'Editor=Wide load 4:  Set your wide load editor to 4 column tabs, fixed size font.  Suggest Kate (Linux) or Notepad++ (windows).


'################################ SETTINGS FOR AnnotatedBackups ###################################

'1 - Subfolder to hold backups (relative path).  Note: GETsPath="" can lead to data loss and will therefor produce a fatal warning.
Sub GETsPath()			As String
	GETsPath		= "AnnotatedBackups"
		'Example:
		'	GETsPath="foo"	Will put backups in .../documentdir/foo/, where documentdir hold your Base, Calc, Writer or other LO document.
End Sub


'2 - Max number of timestamped backup files to be retained (per file).  0=no limit.  
Sub GETiMaxCopies()	As Integer
	GETiMaxCopies	= 50						'	e.g. if have 10 and you set this to 8, then the 2 oldest backups will be auto deleted.
												'	All backup files need to be accessed to find the oldest, so huge values may be slow.
End Sub


'3 - Expert adjustments.  Enable lines below if you often need to save your documents in non-native formats 
Sub GETsB()			As Array
	Dim sB(200) As String				'Array of file types to possibly backup
	Dim i 		As Long		:i = 0		'Index into sB array of 


	'Example:  		Each time you backup you can easily save your Writer documents in these formats: .odt, .doc, .swx, etc.
	'
	'General:		'sB is an array that defines the components of your backkup job.  
	'				For any given module, it allows for multiple file formats to be saved at the same time (in seperate files).
	'
	'sB Structure:	Each line below contains five vertical bar, delimited columns with an optional comment, as follows:
	'					Backup?			Contains either `BACKUP` (all caps) to cause this filter to be run, or an empty string (of tabs or spaces).
	'					Module			Name of the associated Module name
	'					Ext				The file name extension for this file type
	'					Description		A brief description of the format
	'					Filter name		The name of the filter that will write the file in the proper format
	'					Comment			Varies, but an attempt to keep track of the working status of the filter.
	
	'Usage:			Enter BACKUP in the `Backup?` column to add an additional (supplemental) backup format to your backup job.  White space ignored.  
	'	
	' 					If you enable several file types with the same extension, only the backup file with the last filter will be saved.  
	' 					(In other words, earlier filter outputs will be overwritten).
	'
	'					Some lines are commented out (i.e. with a leading ', '* or '?).  See legend below.
	
	'Caution!		Some comments advise: "Strongly reccommend keeping this line unchanged".  
	'					If you turn off these ODF backups, which are for the native formats, you might loose 
	'					some of your work and not be able to recover!
	

	'		       Bakup?	|Module		|Ext	|Description		   								|Filter name								|Comment
	'===BASE=======---------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------
	i=i+1:sB(i) = "BACKUP	|Base		|odb	|ODF (Open Document Format) Database				|Base8"										'Strongly reccommend keeping this line unchanged
	'===BASE-END===---------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------



	'		       Bakup?	|Module		|Ext	|Description		   								|Filter name								|Comment
	'===CALC====== ---------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------
	i=i+1:sB(i) = "BACKUP	|Calc		|ods	|ODF (Open Document Format) Spreadsheet				|Calc8"										'Strongly reccommend keeping this line unchanged
	i=i+1:sB(i) = "BACKUP	|Calc		|ots	|ODF (Open Document Format) Spreadsheet Template	|calc8_template"							'in 5.2 save as menu

	
	'-CALC SAVE AS OTHER:
'*	i=i+1:sB(i) = "			|Calc		|fods	|Flat XML ODF Spreadsheet							*LO SUPPORTED but unknown filter name		'in 5.2 save as menu
'*	i=i+1:sB(i) = "			|Calc		|uos	|Unified Office Format Spreadsheet					*LO SUPPORTED but unknown filter name		'in 5.2 save as menu

'*	i=i+1:sB(i) = "			|Calc		|xlsx	|Microsoft Excel 2007-2013 XML						*LO SUPPORTED but unknown filter name		'in 5.2 save as menu (see also 2nd xlsx below)
'*	i=i+1:sB(i) = "			|Calc		|xlsm	|Microsoft Excel 2007-2016 (macro enabled)			*LO SUPPORTED but unknown filter name		'in 5.2 save as menu (out of order here)
	i=i+1:sB(i) = "			|Calc		|xml	|Microsoft Excel 2003 XML							|MS Excel 2003 XML"							'in 5.2 save as menu
	i=i+1:sB(i) = "			|Calc		|xls	|Microsoft Excel 97/2000/XP							|MS Excel 97"								'in 5.2 save as menu
	i=i+1:sB(i) = "			|Calc		|xlt	|Microsoft Excel 97/2000/XP Template				|MS Excel 97 Vorlage/Template"				'in 5.2 save as menu
	i=i+1:sB(i) = "			|Calc		|xls	|Microsoft Excel 95									|MS Excel 95"								'in 5.2 save as menu
	i=i+1:sB(i) = "			|Calc		|xlt	|Microsoft Excel 95 Template						|MS Excel 95 Vorlage/Template"				'in 5.2 save as menu
	i=i+1:sB(i) = "			|Calc		|xls	|Microsoft Excel 5.0								|MS Excel 5.0/95"							'in 5.2 save as menu
	i=i+1:sB(i) = "			|Calc		|xlt	|Microsoft Excel 5.0 Template						|MS Excel 5.0/95 Vorlage/Template"			'in 5.2 save as menu

	i=i+1:sB(i) = "			|Calc		|dif	|Data Interchange Format							|DIF"										'in 5.2 save as menu
	i=i+1:sB(i) = "			|Calc		|dbf	|dBase												|dBase"										'in 5.2 save as menu

	i=i+1:sB(i) = "			|Calc		|html	|HTML Document (OpenOffice.org Calc)				|HTML (StarCalc)"							'in 5.2 save as menu
	i=i+1:sB(i) = "			|Calc		|slk	|SYLK												|SYLK"										'in 5.2 save as menu

	i=i+1:sB(i) = "			|Calc		|csv	|Text CSV											|Text - txt - csv (StarCalc)"				'in 5.2 save as menu

'*	i=i+1:sB(i) = "			|Calc		|xlsx	|Open Office XML Spreadsheet						*LO SUPPORTED but unknown filter name		'in 5.2 save as menu


	'-CALC SAVE AS OLDER:
	i=i+1:sB(i) = "			|Calc		|sxc	|OpenOffice.org 1.0 Spreadsheet						|StarOffice XML (Calc)"						'older?
	i=i+1:sB(i) = "			|Calc		|stc	|OpenOffice.org 1.0 Spreadsheet Template			|calc_StarOffice_XML_Calc_Template"			'older?

	i=i+1:sB(i) = "			|Calc		|sdc	|StarCalc 5.0										|StarCalc 5.0"								'older?
	i=i+1:sB(i) = "			|Calc		|vor	|StarCalc 5.0 Template								|StarCalc 5.0 Vorlage/Template"				'older?
	i=i+1:sB(i) = "			|Calc		|sdc	|StarCalc 4.0										|StarCalc 4.0"								'older?
	i=i+1:sB(i) = "			|Calc		|vor	|StarCalc 4.0 Template								|StarCalc 4.0 Vorlage/Template"				'older?
	i=i+1:sB(i) = "			|Calc		|sdc	|StarCalc 3.0										|StarCalc 3.0"								'older?
	i=i+1:sB(i) = "			|Calc		|vor	|StarCalc 3.0 Template								|StarCalc 3.0 Vorlage/Template"				'older?


	'-CALC EXPORTS:
'	i=i+1:sB(i) = "			|Calc		|html	|XHTML												|XHTML Calc File"							'in 5.2 export menu (added)
'	i=i+1:sB(i) = "			|Calc		|xhtml	|XHTML												|XHTML Calc File"							'in 5.2 export menu (added)
	i=i+1:sB(i) = "			|Calc		|xml	|XHTML												|XHTML Calc File"							'in 5.2 export menu (?? extension is wrong)
	i=i+1:sB(i) = "			|Calc		|pdf	|PDF - Portable Document Format						|calc_pdf_Export"							'in 5.2 export menu
'*	i=i+1:sB(i) = "			|Calc		|png	|PNG - Portable Network Graphic						*LO SUPPORTED but unknown filter name		'in 5.2 export menu

	i=i+1:sB(i) = "			|Calc		|pxl	|Pocket Excel										|Pocket Excel"								'?
	'===CALC-END===---------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------


	
	'		       Bakup?	|Module		|Ext	|Description		   								|Filter name								|Comment
	'===DRAW=======---------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------
	i=i+1:sB(i) = "BACKUP	|Draw		|odg	|ODF (Open Document Format) Drawing					|Draw8"										'Strongly reccommend keeping this line unchanged
	i=i+1:sB(i) = "			|Draw		|otg	|ODF (Open Document Format) Drawing Template		|draw8_template"


	'-DRAW SAVE AS OLDER:
'?	i=i+1:sB(i) = "			|Draw		|odd	|ODF Drawing										|Draw8"										'Is this an obsolite extension, and we now use .odg?
'*	i=i+1:sB(i) = "			|Draw		|fodg	|ODF Flat XML										*LO SUPPORTED but unknown filter

	i=i+1:sB(i) = "			|Draw		|sxd	|OpenOffice.org 1.0 Drawing							|StarOffice XML (Draw)"						'older?
	i=i+1:sB(i) = "			|Draw		|std	|OpenOffice.org 1.0 Drawing Template				|draw_StarOffice_XML_Draw_Template"			'older?

	i=i+1:sB(i) = "			|Draw		|sxd	|StarDraw 5.0										|StarDraw 5.0"								'older?
	i=i+1:sB(i) = "			|Draw		|vor	|StarDraw 5.0 Template								|StarDraw 5.0 Vorlage"						'older?
	i=i+1:sB(i) = "			|Draw		|sxd	|StarDraw 3.0										|StarDraw 3.0"								'older?
	i=i+1:sB(i) = "			|Draw		|vor	|StarDraw 3.0 Template								|StarDraw 3.0 Vorlage"						'older?

	
	'-DRAW EXPORTS:
	i=i+1:sB(i) = "			|Draw		|html	|HTML - Document (OpenOffice.org Draw)				|draw_html_Export"							'in 5.2 export menu
	i=i+1:sB(i) = "			|Draw		|xml	|XHTML												|XHTML Draw File"							'in 5.2 export menu
	
	i=i+1:sB(i) = "			|Draw		|pdf	|PDF - Portable Document Format						|draw_pdf_Export"							'in 5.2 export menu
	i=i+1:sB(i) = "			|Draw		|swf	|SWF - Macromedia Flash								|draw_flash_Export"							'in 5.2 export menu
	i=i+1:sB(i) = "			|Draw		|bmp	|BMP - Windows Bitmap								|draw_bmp_Export"							'in 5.2 export menu
	i=i+1:sB(i) = "			|Draw		|emf	|EMF - Enhanced Metafile							|draw_emf_Export"							'in 5.2 export menu
	i=i+1:sB(i) = "			|Draw		|eps	|EPS - Encapsulated PostScript						|draw_eps_Export"							'in 5.2 export menu
	i=i+1:sB(i) = "			|Draw		|gif	|GIF - Graphics Interchange Format					|draw_gif_Export"							'in 5.2 export menu
	i=i+1:sB(i) = "			|Draw		|jpg	|JPEG - Joint Photographic Experts Group			|draw_jpg_Export"							'in 5.2 export menu	+.jpeg, .jfif, .jif, .jpe
	i=i+1:sB(i) = "			|Draw		|png	|PNG - Portable Network Graphic						|draw_png_Export"							'in 5.2 export menu
	i=i+1:sB(i) = "			|Draw		|svg	|SVG - Scalable Vector Graphics						|draw_svg_Export"							'in 5.2 export menu	+.jvgz
	i=i+1:sB(i) = "			|Draw		|tiff	|TIFF - Tagged Image File Format					|draw_tif_Export"							'in 5.2 export menu	+.tif
	i=i+1:sB(i) = "			|Draw		|wmf	|WMF - Windows Metafile								|draw_wmf_Export"							'in 5.2 export menu
	
	i=i+1:sB(i) = "			|Draw		|met	|MET - OS/2 Metafile								|draw_met_Export"							'?missing from export menu
	i=i+1:sB(i) = "			|Draw		|pbm	|PBM - Portable Bitmap								|draw_pbm_Export"							'?missing from export menu
	i=i+1:sB(i) = "			|Draw		|pgm	|PGM - Portable Graymap								|draw_pgm_Export"							'?missing from export menu
	i=i+1:sB(i) = "			|Draw		|ppm	|PPM - Portable Pixelmap							|draw_ppm_Export"							'?missing from export menu
	i=i+1:sB(i) = "			|Draw		|ras	|RAS - Sun Raster Image								|draw_ras_Export"							'?missing from export menu
	i=i+1:sB(i) = "			|Draw		|svm	|SVM - StarView Metafile							|draw_svm_Export"							'?missing from export menu
	i=i+1:sB(i) = "			|Draw		|xpm	|XPM - X PixMap										|draw_xpm_Export"							'?missing from export menu
	'===DRAW-END===---------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------


	
	'		       Bakup?	|Module		|Ext	|Description		   								|Filter name								|Comment
	'===IMPRESS====---------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------
	i=i+1:sB(i) = "BACKUP	|Impress	|odp	|ODF (Open Document Format) Presentation			|Impress8"									'Strongly reccommend keeping this line unchanged
'*	i=i+1:sB(i) = "			|Impress	|otp	|ODF (Open Document Format) Presentation Template	*LO SUPPORTED but unknown filter name		'In 5.2 save as menu

	'-IMPRESS SAVE AS OTHER FORMATS:
'*	i=i+1:sB(i) = "			|Impress	|odg	|ODF Drawing (Impress)								*LO SUPPORTED but unknown filter name		'In 5.2 save as menu
	i=i+1:sB(i) = "			|Impress	|odg	|OpenOffice.org 1.0 Drawing (OpenOffice.org Impress)|OpenOffice.org 1.0 Drawing (OpenOffice.org Impress)"	'In 5.2 save as menu 

'*	i=i+1:sB(i) = "			|Impress	|fodp	|Flat XML ODF Presentation							*LO SUPPORTED but unknown filter name		'In 5.2 save as menu
'*	i=i+1:sB(i) = "			|Impress	|uop	|Unified Office Format presentation					*LO SUPPORTED but unknown filter name		'In 5.2 save as menu

'*	i=i+1:sB(i) = "			|Impress	|pptx	|Microsoft PowerPoint 2007-2013						*LO SUPPORTED but unknown filter name		'In 5.2 save as menu (see also pptx below)
'*	i=i+1:sB(i) = "			|Impress	|ppsx	|Microsoft PowerPoint 2007-2013 Autoplay			*LO SUPPORTED but unknown filter name		'In 5.2 save as menu (see also ppsx below)
'*	i=i+1:sB(i) = "			|Impress	|potm	|Microsoft PowerPoint 2007-2013 Template			*LO SUPPORTED but unknown filter name		'In 5.2 save as menu (see also potm below)

	i=i+1:sB(i) = "			|Impress	|ppt	|Microsoft PowerPoint 97/2003/XP					|MS PowerPoint 97"							'In 5.2 save as menu 
'*	i=i+1:sB(i) = "			|Impress	|pps	|Microsoft PowerPoint 97-2003 Autoplay				*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 
	i=i+1:sB(i) = "			|Impress	|pot	|Microsoft PowerPoint 97/2003/XP Template			|MS PowerPoint 97 Vorlage"					'In 5.2 save as menu 

'*	i=i+1:sB(i) = "			|Impress	|pptx	|Office Open XML Presentation						*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 
'*	i=i+1:sB(i) = "			|Impress	|ppsx	|Office Open XML Presentation Autoplay				*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 
'*	i=i+1:sB(i) = "			|Impress	|potm	|Office Open XML Presentation Template				*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 


	'-IMPRESS SAVE AS OLDER:
	i=i+1:sB(i) = "			|Impress	|stp	|Open Document Presentation Template				|impress8_template"							'older?
	
	i=i+1:sB(i) = "			|Impress	|sdd	|StarImpress 5.0									|StarImpress 5.0"							'older?
	i=i+1:sB(i) = "			|Impress	|vor	|StarImpress 5.0 Template							|StarImpress 5.0 Vorlage"					'older?
	i=i+1:sB(i) = "			|Impress	|sdd	|StarImpress 4.0									|StarImpress 4.0"							'older?
	i=i+1:sB(i) = "			|Impress	|vor	|StarImpress 4.0 Template							|StarImpress 4.0 Vorlage"					'older?
	
	i=i+1:sB(i) = "			|Impress	|sda	|StarDraw 5.0 (OpenOffice.org Impress)				|StarDraw 5.0 (StarImpress)"				'older?
	i=i+1:sB(i) = "			|Impress	|vor	|StarDraw 5.0 Template (OpenOffice.org Impress)		|StarDraw 5.0 Vorlage (StarImpress)"		'older?
	i=i+1:sB(i) = "			|Impress	|sdd	|StarDraw 3.0 (OpenOffice.org Impress)				|StarDraw 3.0 (StarImpress)"				'older?
	i=i+1:sB(i) = "			|Impress	|vor	|StarDraw 3.0 Template (OpenOffice.org Impress)		|StarDraw 3.0 Vorlage (StarImpress)"		'older?

	i=i+1:sB(i) = "			|Impress	|sxi	|OpenOffice.org 1.0 Presentation					|StarOffice XML (Impress)"					'older?
	i=i+1:sB(i) = "			|Impress	|sti	|OpenOffice.org 1.0 Presentation Template			|impress_StarOffice_XML_Impress_Template"	'older?
	
	
	'-IMPRESS EXPORTS:
	i=i+1:sB(i) = "			|Impress	|html	|HTML - Document (OpenOffice.org Impress)			|impress_html_Export"						'in 5.2 export menu
	i=i+1:sB(i) = "			|Impress	|xml	|XHTML												|XHTML Impress File"						'in 5.2 export menu

	i=i+1:sB(i) = "			|Impress	|pdf	|PDF - Portable Document Format						|impress_pdf_Export"						'in 5.2 export menu
	i=i+1:sB(i) = "			|Impress	|swf	|SWF - Macromedia Flash								|impress_flash_Export"						'in 5.2 export menu
	
	i=i+1:sB(i) = "			|Impress	|bmp	|BMP - Windows Bitmap								|impress_bmp_Export"						'in 5.2 export menu
	i=i+1:sB(i) = "			|Impress	|emf	|EMF - Enhanced Metafile							|impress_emf_Export"						'in 5.2 export menu
	i=i+1:sB(i) = "			|Impress	|eps	|EPS - Encapsulated PostScript						|impress_eps_Export"						'in 5.2 export menu
	i=i+1:sB(i) = "			|Impress	|gif	|GIF - Graphics Interchange Format					|impress_gif_Export"						'in 5.2 export menu
	i=i+1:sB(i) = "			|Impress	|jpg	|JPEG - Joint Photographic Experts Group			|impress_jpg_Export"						'in 5.2 export menu
	i=i+1:sB(i) = "			|Impress	|png	|PNG - Portable Network Graphic						|impress_png_Export"						'in 5.2 export menu
	i=i+1:sB(i) = "			|Impress	|svg	|SVG - Scalable Vector Graphics						|impress_svg_Export"						'in 5.2 export menu
	i=i+1:sB(i) = "			|Impress	|tiff	|TIFF - Tagged Image File Format					|impress_tif_Export"						'in 5.2 export menu
	i=i+1:sB(i) = "			|Impress	|wmf	|WMF - Windows Metafile								|impress_wmf_Export"						'in 5.2 export menu
	i=i+1:sB(i) = "			|Impress	|pwp	|PWP - PlaceWare									|placeware_Export"							'in 5.2 export menu
	
	i=i+1:sB(i) = "			|Impress	|met	|MET - OS/2 Metafile								|impress_met_Export"						'?missing from export menu
	i=i+1:sB(i) = "			|Impress	|pbm	|PBM - Portable Bitmap								|impress_pbm_Export"						'?missing from export menu
	i=i+1:sB(i) = "			|Impress	|pct	|PCT - Mac Pict										|impress_pct_Export"						'?missing from export menu
	i=i+1:sB(i) = "			|Impress	|pgm	|PGM - Portable Graymap								|impress_pgm_Export"						'?missing from export menu
	i=i+1:sB(i) = "			|Impress	|ppm	|PPM - Portable Pixelmap							|impress_ppm_Export"						'?missing from export menu
	i=i+1:sB(i) = "			|Impress	|ras	|RAS - Sun Raster Image								|impress_ras_Export"						'?missing from export menu
	i=i+1:sB(i) = "			|Impress	|svm	|SVM - StarView Metafile							|impress_svm_Export"						'?missing from export menu
	i=i+1:sB(i) = "			|Impress	|xpm	|XPM - X PixMap										|impress_xpm_Export"						'?missing from export menu
	'===IMPRESS-END---------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------


	'		       Bakup?	|Module		|Ext	|Description		   								|Filter name								|Comment
	'===MATH====== ---------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------
	i=i+1:sB(i) = "BACKUP	|Math		|odf	|ODF (Open Document Format) Formula					|Math8"										'Strongly reccommend keeping this line unchanged

	
	'-MATH SAVE AS OTHER:
'	i=i+1:sB(i) = "			|Math		|mml	|MathML 2.0											*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 

	'-MATH EXPORTS:
	i=i+1:sB(i) = "			|Math		|pdf	|PDF - Portable Document Format						|math_pdf_Export"							'in 5.2 export menu -is this filter ok???
	'===MATH-END------------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------
	
	
	
	'		       Bakup?	|Module		|Ext	|Description		   								|Filter name								|Comment
	'===WRITER=====---------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------
	i=i+1:sB(i) = "BACKUP	|Writer		|odt	|ODF (Open Document Format) Text (Writer)			|Writer8"									'Strongly reccommend keeping this line unchanged
	i=i+1:sB(i) = "			|Writer		|ott	|ODF (Open Document Format) Text (Writer) Template	|writer8_template"							'In 5.2 save as menu 
	
	'-WRITER SAVE AS OTHER FORMATS:
'*	i=i+1:sB(i) = "			|Writer		|fodt	|Flat XML ODF Text Document							*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 
'*	i=i+1:sB(i) = "			|Writer		|uof	|Unified Office Format Text							*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 

'	i=i+1:sB(i) = "			|Writer		|docx	|Microsoft Word 2007-2013 XML						*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 
	i=i+1:sB(i) = "			|Writer		|xml	|Microsoft Word 2003 XML							|MS Word 2003 XML"
'	i=i+1:sB(i) = "			|Writer		|doc	|Microsoft Word 97-2003 							*LO SUPPORTED but unknown filter name		'In 5.2 save as menu (see doc listings below):
	i=i+1:sB(i) = "			|Writer		|doc	|Microsoft Word 97/2000/XP							|MS Word 97"								'(older description)
	i=i+1:sB(i) = "			|Writer		|doc	|Microsoft Word 95									|MS Word 95"								'older?
	i=i+1:sB(i) = "			|Writer		|doc	|Microsoft Word 6.0									|MS WinWord 6.0"							'older?
	
'	i=i+1:sB(i) = "			|Writer		|dot	|Microsoft Word 97-2003 Template					*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 

	i=i+1:sB(i) = "			|Writer		|xml	|DocBook											|DocBook File"								'In 5.2 save as menu 


	'-WRITER SAVE AS OLDER:
	i=i+1:sB(i) = "			|Writer		|sxw	|OpenOffice.org 1.0 Text Document					|StarOffice XML (Writer)"					'older?
	i=i+1:sB(i) = "			|Writer		|stw	|OpenOffice.org 1.0 Text Document Template			|writer_StarOffice_XML_Writer_Template"		'older?

	i=i+1:sB(i) = "			|Writer		|sdw	|StarWriter 5.0										|StarWriter 5.0"							'older?
	i=i+1:sB(i) = "			|Writer		|vor	|StarWriter 5.0 Template							|StarWriter 5.0 Vorlage/Template"			'older?
	i=i+1:sB(i) = "			|Writer		|sdw	|StarWriter 4.0										|StarWriter 4.0"							'older?
	i=i+1:sB(i) = "			|Writer		|vor	|StarWriter 4.0 Template							|StarWriter 4.0 Vorlage/Template"			'older?
	i=i+1:sB(i) = "			|Writer		|sdw	|StarWriter 3.0										|StarWriter 3.0"							'older?
	i=i+1:sB(i) = "			|Writer		|vor	|StarWriter 3.0 Template							|StarWriter 3.0 Vorlage/Template"			'older?

	i=i+1:sB(i) = "			|Writer		|bib	|BibTeX												|BibTeX_Writer"								'older?

	
	'-WRITER EXPORTS:
	i=i+1:sB(i) = "			|Writer		|html	|XHTML												|XHTML Writer File"							'in 5.2 export menu
	i=i+1:sB(i) = "			|Writer		|html	|HTML Document (OpenOffice.org Writer)				|HTML (StarWriter)"							'older?
	i=i+1:sB(i) = "			|Writer		|pdf	|PDF - Portable Document Format						|writer_pdf_Export"							'in 5.2 export menu
'	i=i+1:sB(i) = "			|Writer		|txt	|MediaWiki Text										*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 
	i=i+1:sB(i) = "			|Writer		|txt	|Text												|Text"										'(older description)
	i=i+1:sB(i) = "			|Writer		|txt	|Text Encoded										|Text (encoded)"							'(older description)
'	i=i+1:sB(i) = "			|Writer		|jpg	|JPEG - Joint Photographic Experts Group			*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 
'	i=i+1:sB(i) = "			|Writer		|xml	|Writer Layout										*LO SUPPORTED but unknown filter name		'In 5.2 save as menu 
'	i=i+1:sB(i) = "			|Writer		|png	|PNG - Portable Network Graphic						*LO SUPPORTED but unknown filter name		'in 5.2 export menu

	i=i+1:sB(i) = "			|Writer		|ltx	|LaTeX 2e											|LaTeX_Writer"								'?missing from export menu
	i=i+1:sB(i) = "			|Writer		|pdb	|AportisDoc (Palm)									|AportisDoc Palm DB"						'?missing from export menu
	i=i+1:sB(i) = "			|Writer		|psw	|Pocket Word										|PocketWord File"							'?missing from export menu
	i=i+1:sB(i) = "			|Writer		|rtf	|Rich Text Format									|Rich Text Format"							'?missing from export menu
	'===WRITER-END=---------|-----------|-------|---------------------------------------------------|-------------------------------------------|--------------------------------------

	
	'---------------------------------------------------------------------------------------------------------------------
	'Legend:
	' '?	= Unknown backup type (possibly older).
	' '* 	= Was missing in list above.  From current 5.2.3.3 Save-As menu.
	' '		= Other reason for disabling, but keeping in this list.
	
	GETsB	= Sb
End Sub


' %%%%%% NOTES %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	InStr(		 str, str)
'	InStr(start, str, str)
'	InStr(start, str, str, mode)
'		Attempt to find string 2 in string 1. Returns 0 if not found and starting location if it is found.
'		The optional start argument indicates where to start looking. The default value for mode is 1
'		(case-insensitive comparison). Setting mode to 0 performs a case-sensitive comparison.
'
'	InStrRev(str, find, start, mode)
'		Return the position of the first occurrence of one string within another, starting from the right
'		side of the string. Only available with “Option VBASupport 1”. Start and mode are optional.

