Option Explicit	'BASIC	###### AnnotatedBackups v 1.5.12 ######

'Editor=Wide load 4:  Set your wide load editor to 4 column tabs, fixed size font.  Suggest Kate (Linux) or Notepad++ (windows).


' --- NAME AND PURPOSE -----------------------------------------------------------------------------------------------------
'
'	AnnotatedBackups  - On-demand Save and Backup for LibreOffice. 


' --- DOCUMENTATION --------------------------------------------------------------------------------------------------------
'
'	https://github.com/TopView/AnnotatedBackups


' --- CREDITS --------------------------------------------------------------------------------------------------------------
'
'	Rewritten from earlier work named 'AutomaticBackup' by squenson, & extended by Ratslinger.  Ref & credits:
'
'	squenson: https://forum.openoffice.org/en/forum/memberlist.php?mode=viewprofile&u=2781&sid=78e2eae7c08fba145326798ec04077b8
'	[Basic] Save a document and create a timestamped copy: https://forum.openoffice.org/en/forum/viewtopic.php?f=21&t=23531
'
'	ratslinger: https://ask.libreoffice.org/en/question/88856/suggeston-for-location-of-backup-files/?answer=89030#post-id-89030
'	see also: https://ask.libreoffice.org/en/question/75460/libo-515-on-debian-85-writer-close-without-asking-to-save-in-need-of-an-automatic-incremental-saving-function/


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


'=== Do one or more backups of the current or given file, possibly removing older backups =============
Sub AnnotatedBackups()			'was: Sub AnnotatedBackups(Optional oDoc As Object)

	'--- Get settings from AnnotatedBackupsSettings module (so we can upgrade w/o loosing settings) ------
	Dim sPath 		As String	:sPath 		= GETsPath()		'Relative path from documents to backups.
	Dim iMaxCopies 	As Integer	:iMaxCopies = GETiMaxCopies()	'Max number of timestamped backup files to be retained (per file).  See GetiMaxCopies.
	Dim sB(200) 	As String	:sB()		= GETsB()			'Array of file types to possibly backup


	'--- Get optional comment and honor abort request -----------------------------------
	'(do this early, so as not to save or close anything if canceled)															'
	Dim sComment	As String
	sComment	 	= InputBox(	"(Reminders: iMaxCopies=" & iMaxCopies & ", relative backup path = ./" & sPath & ")" & chr(10) & chr(10) &_
								"Optional backup description, (this gets appended to backup's filename):",_
								"Filename Annotation","none")		'"none" is needed because an empty string and a cancel button are the same thing.
	If sComment = "" Then Exit Sub
	
	sComment = iif(sComment = "none", "", " " & sComment)


	'--- Get document -------------------------------------------------------------------															'
	Dim oDoc 		As Object	:oDoc		= ThisComponent		'Was: "If IsMissing(oDoc) Then Dim oDoc As Object :oDoc = ThisComponent" But, not sure what the If was for as it doesn't work.
	DIM sUrl_From	AS STRING	:sUrl_From	= oDoc.URL			'Default From URL


	'--- Adjust oDoc if called from a Base Form, and also close any open Base Forms, but without chaning oDoc used by other LO modules
'	msgbox "Title: " & oDoc.Title & chr(10) & "URL: " & oDoc.URL &chr(10) & "db: " & oDoc.supportsService("com.sun.star.sdb.OfficeDatabaseDocument") & "    parent: " & iif(isnull(oDoc.parent),"null","not"): stop
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
		if		not	iBaseFormsClosed(oDoc) Then Exit Sub					'so make sure that all forms are closed

	Else 																	' other: A Base Form, or another LO module, i.e. Calc, Draw, Impress, Math, or Writer
		If not isnull(oDoc.parent) then 									'If a Base form?
			'Unravel - allow this to be run from within a Base Form		-Note!  A new, non-base docuemnt's URL is also empty.
			DO WHILE sUrl_From = ""		: oDoc = oDoc.Parent	:sUrl_From	= oDoc.URL	:LOOP
			if 	not	iBaseFormsClosed(oDoc) Then Exit Sub					'again, make sure that all forms are closed
		End If
	End If



	'--- Make sure document is saved before proceeding ----------------------------------															'
	' If a new document is open (i.e. unsaved), then first save it so we can get its path and filter type
	If Not(oDoc.hasLocation) Then
		Dim oDocNew		As Object	:oDocNew 	 = oDoc.CurrentController.Frame
		Dim oDispatcher As Object	:oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
		oDispatcher.executeDispatch(oDocNew, ".uno:SaveAs", "", 0, array())		'Create a new file
	End If

	If Not(oDoc.hasLocation) Then
		Msgbox 	"ERROR - Your document was not been saved, so backup failed. " & _
				"(Perhaps you tried saving to an unwritable folder or to a file already in use.)"
		Exit Sub
	End If


	' --- Set up pathnames (document to backup and place to back it up) -----------------

	' - Get document to backup's name w/ full path
	' Retrieve the document name and path from the URL
	Dim sDocURL					As String	:sDocURL 				= oDoc.getLocation()		'used once below
	Dim sDocNameWithFullPath 	As String	:sDocNameWithFullPath	= ConvertFromURL(sDocURL)	'Source path/filename


	'Detect from path if we have "/" (Linux) or "\" (Windows)
	Dim sSlash 		As String
	Dim sOtherSlash As String
	If Instr(1, sDocNameWithFullPath, "/") > 0 _
		Then :sSlash = "/"	:sOtherSlash = "\"		'Linux
		Else :sSlash = "\"	:sOtherSlash = "/"		'Windows
	End If


	' --- extract filename --------------------------------------------------------------
	Dim sDocName As String	:sDocName = GetFileName(sDocNameWithFullPath, sSlash)				'Source filename


	' If sPath contains the wrong delimiter for our OS, replace it
	Dim s As String	:s = ""						'move sPath into this char at a time, but replacing / with \ (or vise versa)
	Dim i As Integer							'Index used three places
	For i = 1 to Len(sPath)
		If Mid(sPath, i, 1) = sOtherSlash _
			Then :s = s & sSlash
			Else :s = s & Mid(sPath, i, 1)
		End If
	Next i
	sPath = s


	' --- default folder ----------------------------------------------------------------
	' sPath is relative.
	'	""		relative - put backups where document is stored		.../basedir/.
	'	"foo	relative - put backups where document is stored		.../basedir/foo/.
	'	"/"		relative - put backups where document is stored		.../basedir/.			Note: Leading slash is ok too
	'	"/foo	relative - put backups where document is stored		.../basedir/foo/.		Note: Leading slash is ok too
	i 										=  Len(sDocNameWithFullPath)
	While 									   Mid(sDocNameWithFullPath, i,  1) <> sSlash	:i=i-1	:Wend		'strip off doc filename to get abs ppath
	Dim sAbsPath	As String	:sAbsPath	= Left(sDocNameWithFullPath, i - 1) & "/" & sPath					'DocumentPath/sPath


	' --- Check if the backup folder exists, if not we create it ------------------------
	On Error Resume Next
	MkDir sAbsPath																	'Create directory (if not already found)
	On Error Goto 0

	
	' --- Add a slash at the end of the path, if not already present --------------------
	If Right(sAbsPath, 1) <> sSlash Then sAbsPath = sAbsPath & sSlash

	
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
	Dim sBackupName		As String	:sBackupName		= sDocName & sTimeStamp & sComment 		'used twice below
	
	Dim sDocType 		As String
	Dim sExt			As String	'file name extension
	Dim sSaveToURL		As String
	i = 1
	While sB(i) <> ""
		If 																					GetField(sB(i), "|", 1) = "BACKUP" Then	'Future: replace GetField -> split the line once into a array

			sDocType 			= 															GetField(sB(i), "|", 2)
			If  _
				(sDocType = "Base"		And oDoc.supportsService("com.sun.star.sdb.OfficeDatabaseDocument"			)) Or _
				(sDocType = "Calc"		And oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument"			)) Or _
				(sDocType = "Draw"		And oDoc.supportsService("com.sun.star.drawing.DrawingDocument"				)) Or _
				(sDocType = "Impress"	And oDoc.supportsService("com.sun.star.presentation.PresentationDocument"	)) Or _
				(sDocType = "Math"		And oDoc.supportsService("com.sun.star.formula.FormulaProperties"			)) Or _
				(sDocType = "Writer"	And oDoc.supportsService("com.sun.star.text.TextDocument"					)) 	  _
			Then																											'??Think this is right enough for formula.*

				sExt 			= 															GetField(sB(i), "|", 3)			'file name extension (used 2 places)

					sSaveToURL	= ConvertToURL(  		sAbsPath &   sBackupName & "." & sExt)								'Name to save to
					oDoc.storeToUrl(sSaveToURL, Array(MakePropertyValue( "FilterName", 		GetField(sB(i), "|", 5) ) ) )	'Now run the filter to write out the file

					PruneBackupsToMaxSize(iMaxCopies,	sAbsPath,Len(sBackupName & "." & sExt),sDocName,sExt)				'And finally possibly remove older backups to limit number of them kept

			End If
		End If
		i = i + 1
	Wend
		
End Sub




'##################################################################################################
'############################################ FUNCTIONS ###########################################
'##################################################################################################

'=== Close any Base forms (prompting user if necessary) ===========================================
Function iBaseFormsClosed(oDoc As Object) As Integer
	'--- Count open Tables, Queries, & Forms (i.e. they are inside Frames of Frames in the Desktop)
	Dim iOpenForms	As Integer	:iOpenForms	= 0		'Forms							-Can auto-close these
	Dim iOpenTQs	As Integer	:iOpenTQs	= 0		'TQ		 = Tables or Queries	-Can't auto-close these (not yet at least)
	
	Dim iFrame		As integer
	Dim iForm		As integer


'=======================================================================================
'0 Lookup.odb - LibreOffice Base
'0 b Companies - Lookup - LibreOffice Base: Table Data View
'0 b Query1 - Lookup - LibreOffice Base: Table Data View
'0 b Lookup.odb : Sample search and edit form - LibreOffice Base: Database Form



	For iFrame=0 To StarDesktop.Frames.Count-1 Step 1

		'Looking for titles like: "Lookup5.odb - LibreOffice Base"
		If instr(StarDesktop.Frames.getByIndex(iFrame).Title, oDoc.Title & " - LibreOffice Base")<>0 Then
'			msgbox(iFrame & " " & StarDesktop.Frames.getByIndex(iFrame).Title

			For iForm=0 To StarDesktop.Frames.getByIndex(iFrame).Frames.Count-1 Step 1

				'Looking for Forms which have titles like: "Lookup5.odb : <form name>"
				If instr(StarDesktop.Frames.getByIndex(iFrame).Frames.getByIndex(iForm).Title,oDoc.Title & " : ")<>0 _
				Then :iOpenForms	= iOpenForms+1	':msgbox "found a child window that's also a form"
				Else :iOpenTQs		= iOpenTQs  +1	':msgbox "found a child window"
				End if
			Next iForm
		End if
	Next iFrame


	'--- Now if any forms are open (with possibly unsaved edits!) then ask to close them, or abort the backup.  
	'		(Because I can't figure out how to save any current records changes before the backup).
	if iOpenForms Then

		If 	msgBox(iOpenForms & " form" & iif(iOpenForms=1," is open and needs","s are open and need") 	&_
					" to be closed before backup." & chr(10) & chr(10) 									&_
					"Ok to close " & iif(iOpenForms=1,"it","them") & " now?",							 _
					4+32+128,_
					"Preparing to backup") 	= 7 Then iBaseFormsClosed = False: Exit Function		'4=Yes/No + 32="?" + 128=first button (Yes) is default

		'-Close all my forms (open or not).  This is harmless as some of them might already be closed, but I can't tell here which ones.
		Dim oForms 	As Object	:oForms	= oDoc.FormDocuments
		If oForms.count Then 
			For iForm=0 To oForms.count-1
				oForms.getByIndex(iForm).close
			Next iForm
		End If
	End If
	
	'I don't yet know how to close Table and Query windows (like w/ Forms above, so for now this warning will have to do)
	if iOpenTQs Then
		If 	msgBox(iOpenTQs & " Table or Query window" & iif(iOpenTQs=1," is","s are") 					&_
					" open & may contain unsaved records." & chr(10) & chr(10) 							&_
					"Ok to Ignore? Or Cancel to be safe, manually close, & retry?",						 _
					1+32+256,_
					"CAUTION!") 				= 7 Then iBaseFormsClosed = False: Exit Function	'1=Ok/Cancel + 32="?" + 256=2nd button (Yes) is default
	End If

	iBaseFormsClosed = True		'Successful
End Function



'=== Return filename /wo path or ext ==============================================================
Function GetFileName(byVal sNameWithFullPath as String, byval sSlash as String) as String

	Dim i as Long
	Dim j as Long

	GetFileName = ""

	' Let's search from the end of the full name
	' a "." that will indicate the end of the
	' file name and the beginning of the extension
	For i = Len(sNameWithFullPath) To 1 Step -1
		If Mid(sNameWithFullPath, i, 1) = "." Then

			' We have found a ".", so now we continue
			' backwards and search for the path delimiter "\" or "/"
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
   
End Function



'=== Return nth text field =========================================================================
' e.g. n=2 from "xxx|yyy|foo", gives "yyy".  (n=0 ok, but returns nothing)
Function GetField(byVal sInput As String, byVal sDelimiter As String, ByVal n As Integer) 	As String

	sInput = sDelimiter & sInput & sDelimiter							'To simplify searching sandwitch the input in outer delimiters
	GetField = ""														'Default output if field or is empty found

	Dim iStart	As Long	:iStart = 1										'Char position after nth delimiter

	'A) Find the character position of the nth delimiter
	Dim i 		As Integer
	For i = 1 to n
		iStart = InStr(iStart, sInput, sDelimiter)+1					'Char position after ith delimiter
		If iStart = 1 			Then Exit Function						'If search fails, i.e. ran out of delimiters too soon			, then silently return an empty string
	Next i
	If 	iStart = Len(sInput)+1 	Then Exit Function						'If search fails, i.e. found nth delimiter, but its the last one, then silently return an empty string

	'B) Find the character position before the next delimiter
	Dim iEnd	As Long	:iEnd = InStr(iStart, sInput, sDelimiter)		'Char position at    found delimiter n+1

	'C) Return the portion of string between the two delimiters
	GetField = RTrim(Mid(sInput, iStart, iEnd - iStart))				'Input, Start, Length	(ignore extra delimiters we put on)

End Function


''--Tests to make sure code above is working (uncomment and single step through)
'function test_GetField()	
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
'end function



'=== Returns file filter type =====================================================================
' 	Credit: http://www.oooforum.org/forum/viewtopic.phtml?t=52047
Function GetFilterType(byVal sFileName as String) as String

	'Get access to UNO methods ("services")
	Dim oSFA 			As Object	:oSFA		= createUNOService("com.sun.star.ucb.SimpleFileAccess"		)
	Dim oTD 			As Object	:oTD		= createUnoService("com.sun.star.document.TypeDetection"	)
	
	'Open given filename for reading
'	Dim sURL 			As String	:sUrl		= ConvertToUrl(sFileName)
	Dim oInpStream 		As Object	:oInpStream	= oSFA.openFileRead(ConvertToUrl(sFileName))			'open given filenmae using ucb.SimpleFileAccess 

		'Get it's Type
'		Dim aProps(0) 	As new  com.sun.star.beans.PropertyValue
'			aProps(0).Name	= "InputStream"
'			aProps(0).Value	=   oInpStream

'		GetFilterType	= oTD.queryTypeByDescriptor(aProps(), true)										'queryTypeByDescriptor
		GetFilterType	= oTD.queryTypeByDescriptor(MakePropertyValue("InputStream",oInpStream), true)	'queryTypeByDescriptor

	oInpStream.closeInput()																				'close

End Function


'=== Create and return a new com.sun.star.beans.PropertyValue =====================================
Function MakePropertyValue(Optional sName As String, Optional sValue) As com.sun.star.beans.PropertyValue

    Dim oPropertyValue As New com.sun.star.beans.PropertyValue

    If Not IsMissing(sName	) Then oPropertyValue.Name 	= sName   
    If Not IsMissing(sValue	) Then oPropertyValue.Value = sValue

    MakePropertyValue() = oPropertyValue

End Function


' === possibly remove older backups =====================================================
Sub PruneBackupsToMaxSize(iMaxCopies As Integer, sAbsPath As String, iLenBackupName As Integer, sDocName As String, sExt As String)
	if iMaxCopies = 0 then exit sub											'If iMaxCopies is = 0, there is no need to read, sort or delete any files.

		
	' --- First get list of existing backups --------------------------------------------
	Dim mArray() 		As String											'Array to store list of existing backup path/file names
	Dim iBackups 		As Integer 	:iBackups 		= 0						'Count of existing backup files	
	
	Dim stFileName 		As String	:stFileName 	= Dir(sAbsPath, 0)		'Get FIRST normal file from pathname
	Do While (stFileName <> "")
	
		'Huristic to test for deletable backups
'		If 	 Len(stFileName					) = iLenBackupName	And	_		'patch this in to only remove un-commented names and not purge names given a suffix comment
		If _
			Left(stFileName,Len(sDocName)	) = sDocName		And _
		   Right(stFileName,3				) = sExt 				_
		   Then	:ReDIM Preserve mArray(iBackups)	:mArray(iBackups) = stFileName	:iBackups = iBackups+1 	'get list of existing backups
		End if
									 stFileName 	= Dir()					'Get NEXT  normal file from pathname as initially used above
	Loop


	'--- Sort list of existing backups (by timestamp, oldest first) ---------------------
	iSort(mArray)


	'--- Deleting oldest files ----------------------------------------------------------
	'Deletes oldest files exceeding the limit set in iMaxCopies
	Dim iKill			As Integer	:iKill = iBackups - iMaxCopies								'# of old backups to delete
	Dim x 				As Integer	:For x = 0 to iKill -1: Kill(sAbsPath & mArray(x)): Next x	'now delete them

End Sub



'=== insertion sort (oldest first) ================================================================
Function iSort(mArray)
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
End Function


''--Test of iSort() function
'Function test_iSort()
'	Dim mArrayA(2) As String
'	mArrayA(0) = "x3"
'	mArrayA(1) = "x2"
'	mArrayA(2) = "x1"

'	iSort(mArrayA)

'	Dim x0 As String	:x0 = mArrayA(0)
'	Dim x1 As String	:x1 = mArrayA(1)
'	Dim x2 As String	:x2 = mArrayA(2)
'End Function



' === Trim spaces and tabs from right end =========================================================
Function RTrim(str As String) As String
	RTrim=str																		'simplify code; use returned value as working string
	Dim i as Long																	'character counter (from end to start)
	For i = Len(RTrim) to 1 step -1
		If right(RTrim,1) <> chr(9) and right(RTrim,1) <> " " Then Exit Function	'if trailing white space not found we're done
		RTrim = left(RTrim,len(RTrim)-1)											'otherwise remove trailing white space, step left, repeat
	Next
End function
