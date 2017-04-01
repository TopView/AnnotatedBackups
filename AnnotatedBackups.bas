Option Explicit	'BASIC	###### ANNOTATEDBACKUPS ######

'Editor=Wide load 4:  Set your wide load editor to 4 column tabs, fixed size font.  Suggest Kate (Linux) or Notepad++ (windows).



' --- NAME & CREDITS -------------------------------------------------------------------------------------------------------

'	 AnnotatedBackups  - rewritten from earlier work named 'AutomaticBackup' by squenson, & extended by Ratslinger.  Ref & credits:

'		squenson: https://forum.openoffice.org/en/forum/memberlist.php?mode=viewprofile&u=2781&sid=78e2eae7c08fba145326798ec04077b8
'		[Basic] Save a document and create a timestamped copy: https://forum.openoffice.org/en/forum/viewtopic.php?f=21&t=23531
'
'		ratslinger: https://ask.libreoffice.org/en/question/88856/suggeston-for-location-of-backup-files/?answer=89030#post-id-89030
'		see also: https://ask.libreoffice.org/en/question/75460/libo-515-on-debian-85-writer-close-without-asking-to-save-in-need-of-an-automatic-incremental-saving-function/


' --- Creative Commons - Attribution-ShareAlike / CC BY-SA ----------------------------------------------------------------
'	The previous work that this is based on was not oficially licensed, but still freely offered for use and further development.  Because I was asked
'	to select a license when uploading this to LibreOffice Wiki, I chose this which seems in keeping with the intent of the the other authors.
'
'	This license lets others remix, tweak, and build upon your work even for commercial purposes, as long as they credit you and license 
'	their new creations under the identical terms. This license is often compared to “copyleft” free and open source software licenses. 
'	All new works based on yours will carry the same license, so any derivatives will also allow commercial use. This is the license used by 
'	Wikipedia, and is recommended for materials that would benefit from incorporating content from Wikipedia and similarly licensed projects. 



' --- FEATURES -------------------------------------------------------------------------------------------------------------

'		* Creates one or more copies of the current document with a timestamp and optional comment suffix added to the filename, 
'		  into an absolute or relative backup folder.

'			For example, if the current document is MyTest.odt, each time you run the macro a copy will be saved in  
'				C:\BackupDocs (or any folder you like) and named "MyTest--2009-10-10_17:25:36 MyComment.odt".

'		* Works with:  Base, Calc, Draw, Impress and Writer.  (But does not backup up non-embedded databases.)

'		* Older backups can be automatically pruned by setting the iMaxCopies limit counter to non-zero.

'		Terms:  "Format" and "Filter" is used almost interchangably.  
'			A "filter" is a program to take a document and produce an output file from it.  Sometimes it is just saving
'			the file.  Other times it is converting it into some other format.  oDoc.storeToUrl is called to do the work.
'			See also: https://help.libreoffice.org/Common/About_Import_and_Export_Filters .


' --- INSTALLATION AND USE -------------------------------------------------------------------------------------------------

'		!WARNING: 	If you were using the older AutomaticBackup (see warning below)

'		1) Create a library for the code open LibreOffice:

'			Menu> Tools | Macros | Organize Macros | LibreOffice Basic...				(opens LibreOffice Basic Macros dialog)

'				Select: Organizer...													(opens LibreOffice Basic Macros Organizer dialog)

'					Select: My Macros > Standard										(opens New Module dialog)
'					Click the 'New...' button
'					Enter for Name: "AnnotatedBackups" and click OK

'				Click Close


'		2) Put this Basic code into the newly created libary:

'			Select: AnnotatedBackups and click Edit										(opens My Macros & Dialogs.Standard - LibreOffice Basic dialog)
'				Paste this entire text into the large code window.

'		3) Setyp:

'				Set 'sPath'			- to either an absolute or relative backup path below.
'					!WARNING: 		- If you were using the older AutomaticBackup you might want to give sPath a different path, because the 
'										naming of backup files has changed, and otherwise AnnotatedBackups might wipe out your old backups.
'
'				set 'iMaxCopies'	- to the maximum number of backups retained before purging of older ones begins, or 0 for no purging.


'		4a) Add a menu icon to each of the LibreOffice components (Base, Calc, etc.).  For Writer:

'			Menu> File | New | Text Document											(opens LibreOffice Writer dialog with 'Untitled' document)

'			Menu> Tools | Customize...													(opens Customize dialog)

'				Select the 'Toolbars' tab (if it's not already selected)

'				Set 'Toolbar' to "Standard" 

'				Under 'Toolbar Content' click the 'Add...' button						(opens Add Commands dialog)

'					Under Category select:
'						'LibreOffice Macros' > My Macros > Standard > AnnotatedBackups

'					Under Commands select:
'				
'						"AnnotatedBackups" and and click the 'Add' button				(adds this to the menu in the already open Customize dialog)

'					Click Close															('AnnotatedBackups' will have been added to your Toolbar Content)

'				Use the down arrow button to move the icon below the 'Save As...' menu icon


'		4b) Install a menu icon for AnnotatedBackups:

'			Download AnnotatedBackups.gif to your desktop from (to be added later)		[Possibly someone can create a nicer icon]

'			Click the 'Modify' pull down button and select 'Change Icon...'

'				Click the 'Import...' button.
'
'					In the file type pull down select the GIF file filter
'					Then locate and open 'AnnotatedBackups.gif'

'				Select the new icon, then click Ok.

'			Click Ok

'		5) Test it:

'			When you first click he AnnotatedBackups icon, you are likely to get a Save As file selection window.  
'			This is because the new document you created called 'Untitled' has not yet been saved.  Possiblty rename it, 
'			and then save it someplace.

'			Next you will get a "Backup as: '<file type>' ?  iMaxCopies =nnn)" dialog.

'			Extend the default filename with a few words to say what this backup represents and then click Ok.  You can also just click Ok.

'			By default the backup folder is set to 'Relative' below and named 'Backup', so you can look in the Backup/. subdirectory of where 
'			you saved your document for your backups.

'			After you have created a few backups, you will notice that any more than the iMaxLimit are purged, 
'			so you always retain the iMaxCopies most recent backups.


'		6) "For advanced users, this macro can make several copies of the document, in different formats. 
'				For example, you can save a Writer document in the format .odt, .doc, .swx, etc." 	- squenson Oct 2009


'	USE AT YOUR OWN RISK.  No claims or warranties implied or otherwise as to its’ performance or accuracy.
'	Please send updates, corrections, or suggestions to: EasyTrieve <CustomerService@OutWestClassifieds.org>



' --- HISTORY --------------------------------------------------------------------------------------------------------------
' v 1.6.07		2017-03-28	Fixed unravel of oDoc so only unravels Base Forms.  Added support for Math.
' v 1.5.06		2017-03-31	Counts only my open forms which might need to be closed.
' v 1.5.05		2017-03-28	Default iMaxCopies: 50
' v 1.5.04		2017-03-28	Default backup path: /AnnotatedBackups (relative)
'
' v 1.5.03		2017-03-28	Corrected Annotated name spelling.  lol, really!
'							Added warning for those using AutomaticBackup to use a new backup path
'
' v 1.5.02		2017-03-18	Reformatted for wide screens (additional styling comments are foun at end of code below),
'							Simplified code,
'							Removed unused variables,
'							Improved some procedure and variable names, 
'							Added many comments: improves readability & testability,
'							Organized and updated filter list: added comments, and allow white space with tabs in columns,
'							Simplified sorting,
'							Slightly reformatted timestamp,
'							Added annotation capiability,
'							Allow it to backup from Base even if inside of a Form,
'							Fixed bug in older backup purging, and
'							Added new installation and usage notes.	- EasyTrieve
'
' v 1.4			2017-03-01	Added procedure to auto-remove older backups. - Ratslinger
'
' v 1.3			2009-10-10	Now works with Windows XP and Linux (Ubuntu) file path format containing "/" or "\".
'
' v 1.2			2007-04-16	Add several new filters and make the selection of which filters to use for backup relatively easy for the User.
'							Allow save of the document in case it is not yet saved.
'
' v 1.1.1		2007-04-12	Improve the handling of the path, with possibility of relative path (from current file path) or even empty path 
'							(same folder as current file).
'
' v 1.1.0		2007-04-10	Work with the four main document types of OOo (Writer, Calc, Impress and Draw).
'---------------------------------------------------------------------------------------------------------------------------



'=== Do one or more backups of the current or given file, possibly removing older backups =========
Sub AnnotatedBackups()			'was: Sub AnnotatedBackups(Optional oDoc As Object)
	Dim sB(200) As String				'Array of file types to possibly backup
	Dim i 		As Long		:i = 0		'Index into sB array of 



	'##################################################################################################
	'####################################### SETTABLE OPTIONS #########################################
	'##################################################################################################

	'--- Step 1) Set the following two variables --------------------------------------------------

	Dim iMaxCopies 	As Integer	:iMaxCopies	=50	'Max number of timestamped backup files to be retained (per file).  0=no limit.  
												'	e.g. if have 10 and you set this to 8, then the 2 oldest backups will be auto deleted.
												'	All backup files need to be accessed to find the oldest, so huge values may be slow.
												'	Previously called iMaxFiles.
												
	Dim sPath 		As String	:sPath 		="/AnnotatedBackups"	'Path to backups. Relative if empty or slash prefix, otherwise absolute (not recomended).
	'Examples:
	'	""								= Relative (	recomended ).  Put backups in folder where document is stored			.../documentdir/.
	'	"/foo							= Relative (	recomended ).  Put backups in folder where document is stored+sPath		.../documentdir/foo/.
	
	'	C:\My Documents\BackupFolder	= Absolute (not recomended*).  Put backups in root										C:\My Documents\BackupFolder\.	
	'	"foo"							= Absolute (not recomended*).  Put backups in root										/foo/. 	(note: likely fails in linux root)
	'
	' !!! Absolute sPath is NOT recomended because backups can be overwritten if there are same named documents in multiple different paths.


	' --- Step 2) Enable lines below if you often need to save your documents in non-native formats 
	'
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
	i=i+1:sB(i) = "			|Calc		|ots	|ODF (Open Document Format) Spreadsheet Template	|calc8_template"							'in 5.2 save as menu

	
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

	


	'##################################################################################################
	'########################################  MAIN CODE  #############################################
	'##################################################################################################


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
'	mri oDoc: stop
		If not isnull(oDoc.parent) then 									'If a Base form?
			'Unravel - allow this to be run from within a Base Form		-Note!  A new, non-base docuemnt's URL is also empty.
			DO WHILE sUrl_From = ""		: oDoc = oDoc.Parent	:sUrl_From	= oDoc.URL	:LOOP
			if 	not	iBaseFormsClosed(oDoc) Then Exit Sub					'again, make sure that all forms are closed
'		Else																'its an other app
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
	For i = 1 to Len(sPath)
		If Mid(sPath, i, 1) = sOtherSlash _
			Then :s = s & sSlash
			Else :s = s & Mid(sPath, i, 1)
		End If
	Next i
	sPath = s


	' --- default folder ----------------------------------------------------------------
	' sPath is relative if empty or prefixed with a slash (this is confusing), otherwise it's absolute
	'	"foo"	absolute - put backups in root						/foo/.
	'	""		relative - put backups where document is stored		.../basedir/.
	'	"/foo	relative - put backups where document is stored		.../basedir/foo/.
	If sPath = "" or Left(sPath, 1) = sSlash Then
		i 		=  Len(sDocNameWithFullPath)
		While 	   Mid(sDocNameWithFullPath, i,  1) <> sSlash	:i=i-1	:Wend		'strip off doc filename - search backwards for first slash or backslash
		sPath	= Left(sDocNameWithFullPath, i - 1) & sPath							'DocumentPath/sPath		- prefix SPath with Source path
	End If

	
	' --- Check if the backup folder exists, if not we create it ------------------------
	On Error Resume Next
	MkDir sPath																		'Create directory (if not already found)
	On Error Goto 0

	
	' --- Add a slash at the end of the path, if not already present --------------------
	If Right(sPath, 1) <> sSlash Then sPath = sPath & sSlash

	
	' --- Save current document changes -------------------------------------------------
	' Save the current document only if it has changed, is not new (has been saved before) and is not read only
	If oDoc.isModified and oDoc.hasLocation and Not(oDoc.isReadOnly) Then oDoc.store()


	' --- get timestamp -----------------------------------------------------------------
	'  the timestamp, so it will be identical for all the backup copies
	Dim sTimeStamp 		As String	:sTimeStamp	= "--" & 	Format(Year(	Now), "0000-"	) & _
															Format(Month(	Now), "00-"		) & _
															Format(Day(		Now), "00\_"	) & _
															Format(Hour(	Now), "00:"		) & _
															Format(Minute(	Now), "00:"		) & _
															Format(Second(	Now), "00"		)
										
	' --- do other backups --------------------------------------------------------------
	' For each file filter, let's see whether we should create a backup copy or not
	Dim sBackupName		As String	:sBackupName		= sDocName & sTimeStamp 	'used several places
	
	'-- Prompt for confirmation and possible comment to append
	'		   InputBox(text, title, default text)
	sBackupName 	= InputBox(	"Target directory and filename:   (Tip: You can add a comment to the filename.)" &chr(10)&chr(10)&"    "& sPath ,_
								"Backup?   (iMaxCopies=" & iMaxCopies & ")", sBackupName)

	If len(sBackupName) Then
	
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

						sSaveToURL	= ConvertToURL(  		sPath &   sBackupName & "." & sExt)									'Name to save to
						oDoc.storeToUrl(sSaveToURL, Array(MakePropertyValue( "FilterName", 		GetField(sB(i), "|", 5) ) ) )	'Now run the filter to write out the file

						PruneBackupsToMaxSize(iMaxCopies,	sPath,Len(sBackupName & "." & sExt),sDocName,sExt)					'And finally possibly remove older backups to limit number of them kept

				End If
			End If
			i = i + 1
		Wend
		
	Else
		MsgBox("Backup Canceled",0," ")
	End If
End Sub




'##################################################################################################
'############################################ FUNCTIONS ###########################################
'##################################################################################################

'=== Close any Base forms (prompting user if necessary) ===========================================
Function iBaseFormsClosed(oDoc As Object) As Integer
	'--- Find how many of my forms are open (i.e. they are inside Frames of Frames in the Desktop)
	Dim iOpenForms	As Integer	:iOpenForms = 0
	Dim iFrame		As integer
	Dim iForm		As integer

	For iFrame=0 To StarDesktop.Frames.Count-1 Step 1

		'Looking for titles like: "Lookup5.odb - LibreOffice Base"
		If instr(StarDesktop.Frames.getByIndex(iFrame).Title,".odb - ")<>0 Then
'			msgbox(iFrame & " " & StarDesktop.Frames.getByIndex(iFrame).Title

			For iForm=0 To StarDesktop.Frames.getByIndex(iFrame).Frames.Count-1 Step 1

				'Looking for titles like: "Lookup5.odb : <form name>"
				If instr(StarDesktop.Frames.getByIndex(iFrame).Frames.getByIndex(iForm).Title,oDoc.Title & " : ")<>0 Then
'					msgbox(iFrame & " " & StarDesktop.Frames.getByIndex(iFrame).Frames.getByIndex(iForm).Title
					iOpenForms = iOpenForms+1		':msbbox "found one"
				End If
			Next iForm
		End If
	Next iFrame


	'--- Now if any forms are open (with possibly unsaved edits!) then ask to close them, or abort the backup.  
	'		(Because I can't figure out how to save any current records changes before the backup).
	if iOpenForms Then

		If msgBox(iOpenForms & " form" & iif(iOpenForms=1," is open and needs","s are open and need") 	&_
						" to be closed before backup." & chr(10) & chr(10) 								&_
						"Ok to close " & iif(iOpenForms=1,"","them") & " now?",							 _
					4+32+128,_
					"Preparing to backup") = 7 Then iBaseFormsClosed = False: Exit Function		'4=Yes/No= + 32="?" + 128=first button (Yes) is default

		'-Close all my forms (open or not).  This is harmless as some of them might already be closed, but I can't tell here which ones.
		Dim oForms 	As Object	:oForms	= oDoc.FormDocuments
		If oForms.count Then 
			For iForm=0 To oForms.count-1
				oForms.getByIndex(iForm).close
			Next iForm
		End If
		
	End If
	iBaseFormsClosed = True
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
	Dim oSFA 			As Object	:oSFA		= createUNOService ("com.sun.star.ucb.SimpleFileAccess"		)
	Dim oTD 			As Object	:oTD		= createUnoService(	"com.sun.star.document.TypeDetection"	)
	
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
Sub PruneBackupsToMaxSize(iMaxCopies As Integer, sPath As String, iLenBackupName As Integer, sDocName As String, sExt As String)
	if iMaxCopies = 0 then exit sub											'If iMaxCopies is = 0, there is no need to read, sort or delete any files.

		
	' --- First get list of existing backups --------------------------------------------
	Dim mArray() 		As String											'Array to store list of existing backup path/file names
	Dim iBackups 		As Integer 	:iBackups 		= 0						'Count of existing backup files	
	
	Dim stFileName 		As String	:stFileName 	= Dir(sPath, 0)			'Get FIRST normal file from pathname
	Do While (stFileName <> "")
	
		'Huristic to test for deletable backups
'		If 	 Len(stFileName					) = iLenBackupName	And	_		'patch this is to only remove un-commented names and not purge names given a suffix comment
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
	Dim iKill			As Integer	:iKill = iBackups - iMaxCopies							'# of old backups to delete
	Dim x 				As Integer	:For x = 0 to iKill -1: Kill(sPath & mArray(x)): Next x	'now delete them

End Sub



'=== insertion sort (oldest first) ================================================================
Function iSort(mArray)
	Dim Lb 	as integer	:Lb = lBound(mArray)	'lower array bound
	Dim Ub 	as integer	:Ub = uBound(mArray) 	'upper array bound
	
	Dim iT	As Long		'element under 	Test	, Array index	- What we are looking to possibly move and insert into lower already sorted stuff
	Dim sT 	as string 	'element under 	Test	, Element value	- Variable to hold what we are testing, so cell can get stomped on and not lost by stuff shifting up
	
	Dim iC	as Long		'element to		Compare	, Array index	- Index to search thru what is already sorted, to find what might be bigger than sT 

		
	for iT = Lb+1 to Ub											'Work forwards through array: from second element to last element
		sT = mArray(iT)												'Save element to test and possibly to move down (because will possibly will get stomped on).

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





SUB BaseBackup (oDoc As Object, sPath As string)

	' 1) --- Get sUrl_From (the place to copy from) --------------------------------------------------
	DIM sTitle		AS STRING	:sTitle			= oDoc.Title
	DIM sUrl_From 	AS STRING	:sUrl_From		= oDoc.URL

	'If the macro is run when you launch the ODB file the sUrl_From will be correct. However, if
	'the macro is carried out by a form, it must first determine whether a URL is available. If the URL is
	'empty, a higher level (oDoc.Parent) for a value is looked up.
	DO WHILE sUrl_From = ""
		oDoc		= oDoc.Parent
		sTitle		= oDoc.Title
		sUrl_From	= oDoc.URL
	LOOP

	
	' 2) --- Get sUrl_To (the place to copy to) ------------------------------------------------------
	DIM oPath		AS OBJECT	:oPath			= createUnoService("com.sun.star.util.PathSettings")

	DIM sUrl_To		AS STRING	:sUrl_To		= oPath.Backup & "/" & 99 &"_" & sTitle


	' 3) --- Make the backup copy - Copy Start to End (source to target) -----------------------------
	FileCopy(sUrl_From, sUrl_To)
	
END SUB



'Backup utility from Base manual.  Kept here for reference.  Bug from manual is fixed below.
'	SUB Databasebackup (inMax AS INTEGER)
'
'
'		' 1) --- Get sUrl_From (the place to copy from) --------------------------------------------------
'		DIM oDoc		AS OBJECT	:oDoc			= ThisComponent
'
'		DIM sTitle		AS STRING	:sTitle			= oDoc.Title
'		DIM sUrl_From	AS STRING	:sUrl_From		= oDoc.URL
'
'		DO WHILE sUrl_From = ""
'			'If the macro is run when you launch the ODB file, sTitle and sUrl_From will be correct. However, if
'			'the macro is carried out by a form, it must first determine whether a URL is available. If the URL is
'			'empty, a higher level (oDoc.Parent) for a value is looked up.
'			oDoc		= oDoc.Parent
'			sTitle		= oDoc.Title
'			sUrl_From	= oDoc.URL
'		LOOP
'
'		
'		' 2) --- Get sUrl_To (the place to copy to) ------------------------------------------------------
'		DIM oPath		AS OBJECT	:oPath			= createUnoService("com.sun.star.util.PathSettings")
'
'		'HOW THIS WORKS:
'		'	inMax specifies a pool of possible numbered backups slots.
'		'	As each new backup copy is created, the backup number has to increased by one to the next unfound slot.
'		'	When all backup slots are filled (i.e. i reaches inMax+1) the next backup will overwrite what's in slot 1.
'		'	Then the next backup after that is placed in slot 2.
'		'		A search is done thru existing backups to look for the transition between old and new to find this slot.
'
'		DIM i			AS INTEGER												'Sequence thru pool of existing backup serial numbers to find an empty slot
'		DIM k			AS INTEGER												'If pool is full, then used this to search for the roll around point between old and new.
'
'
'		FOR i = 1 TO inMax + 1													'search upwards for unused backup serial# to use
'
'			'look for a backup-i that does not yet exist
'			IF NOT FileExists(oPath.Backup & "/" & i & "_" & sTitle) THEN		'if file name: path/#_title
'
'				'only do this if backup-i is missing
'				IF i > inMax THEN
'
'					'only do this if i=inMax+1
'	'				FOR k = 1 TO	inMax - 1 TO 1	STEP -1						'Syntax error.  Was this way in manual.  FOR with two "TO"s???  Something was wrong.
'					FOR				inMax - 1 TO 1  STEP -1						'corrected to this
'
'						'Search backwards until dates out of sequence is found, or the roll around point.  Return i where we can overwrite the oldest backup.
'						'	if backup-k is newer than backup-k+1 then get i and exit loops
'						'	if backup-k is older than backup-k+1 then loop to lower value of k (keep searching)
'						IF FileDateTime(oPath.Backup & "/" & k & "_" & sTitle) <= FileDateTime(oPath.Backup & "/" & k+1 & "_" & sTitle) _
'							THEN :IF k = 1 THEN :i = k		:EXIT FOR			'i=  1	;if none newer found (then all in date order) so roll around and overwrite the first slot
'							ELSE :				 i = k + 1	:EXIT FOR			'i=k+1	;else found transition between old and new, e.g. old, older, ^ newer, new; so use ^+1 slot
'						END IF
'
'					NEXT
'
'				END IF
'				EXIT FOR														'Now exit outer loop and use i to create copy below
'
'			END IF
'
'		NEXT	'if backup-i is found, then loop upwards to next i
'
'
'		'Set backup target as set by i above.
'		DIM sUrl_To	AS STRING	:sUrl_To = oPath.Backup & "/" & i &"_" & sTitle
'
'
'		' 3) --- Make the backup copy - Copy Start to End (source to target) -----------------------------
'		FileCopy(sUrl_From, sUrl_To)
'
'	END SUB



'##################################################################################################
'###################################  CODE STYLING COMMENTS  ######################################
'##################################################################################################
'
'Set editor to 4 column tabs, fixed with font, wide screen.
'
'	* Goals: Good code is many things, but for me it's easy to read, and easy to see if it's working properly.
'
'	* Test procedures included: I have deliberately left in some 'commented out test procedures'.  
'		Sometimes when things aren't working or understood it's helpful to quickly test the parts.
'		Other times it's useful to have a simple, runnable example of how a function works and what it can do.
'		Rather than delete and remove from code after development, I choose to leave test code in and comment out.
'
'	* If Then..: You might notice that I wrap "If" structures differently than you're used to.  
'		I use what makes the most sense to me, in terms of simplicty and readability.
'
'	* DIMs: I find that in Basic keeping DIMs close to where they are used makes the most sense. 
'		That way I can quckly and easily (without a lot of scrolling) usuially see who uses what variables.
'		So I often use the colon (:) to put the DIM on the same line as where the var is initialized.  
'		It allows vars to be commented once. :-)
'		In rewriting this, I also removed a number of unused DIMs from the original code that weren't used.  
'
'	* Many renames: I shoot for good, short, & useful names.  (e.g. Rather than j, iStart).
'
'
'	* Wide load: I rewrote this to satisfy my own taste.  I find coding using greater screen width more productive.
'
'	* Blank lines: I use extra blank lines to see blocks.  I like blocks.
'
'	* Horizontal lines: I like horizontal lines to seperate things.  They come in different strengths, - = #.
'
'	* Vertical alignment: I like to line things up vertically.
'		It helps me see many things more clearly which are often related from line to line.  Parallelism.
'
'	* Indenting: I like rigorous indenting with tabs.  Tabs are much quicker to edit than spaces.
'		Tip: Kate and Notepad++ allow easy and quick block indent or unindented using tabs to reglarize indents.
'		Kate and Notepad++ also have block mode, which allows cutting/pasteing vertical blocks.  
'		(Notepad++ has the edge in that space and tab also works in vertical blocks to move a bunch of stuff together.)
'
'	* Comments: You might notice that I use the wide screen to put the code on the left, and comments on the right.  
'		This allows me to read the code without the comments, but keep the comments close by, 
'		and also very importantly, keep the number of lines to the minimum, .. 
'		so I can get the most possible code on the screen at the same time.
'
'
'	* Credits: 
'		-I am endebted to the other authors before me for developing this basic code.  Thank you.  I learned a lot from you!
'		-I revised and simplified the insertion sort and removed the credits.  
'			Heck, the first insertion sort was written before I was born, and I'm an old guy.  I think this history is interesting:
'
'				"The quicksort algorithm was developed in 1959 by Tony Hoare while in the Soviet Union, as a visiting student at Moscow 
'				State University. At that time, Hoare worked in a project on machine translation for the National Physical Laboratory. 
'				As a part of the translation process, he needed to sort the words of Russian sentences prior to looking them up in a 
'				Russian-English dictionary that was already sorted in alphabetic order on magnetic tape.[4] After recognizing that his 
'				first idea, ***insertion sort***, would be slow, he quickly came up with a new idea that was Quicksort."
'
'					ref, see "History" at https://en.wikipedia.org/wiki/Quicksort
'
'		-I am also grateful to a book on the elements of coding style I was forced to read 30 years ago.
