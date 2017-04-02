Option Explicit	'BASIC	###### AnnotatedBackupsSettings for v 1.5.10 ######

'Editor=Wide load 4:  Set your wide load editor to 4 column tabs, fixed size font.  Suggest Kate (Linux) or Notepad++ (windows).


'################################ SETTINGS FOR AnnotatedBackups ###################################

'1 - Relative path from documents to backups.
Function GETsPath()			As String
	GETsPath		= "AnnotatedBackups"
		'Examples:
		'	""		= Put backups in folder where document is stored			.../documentdir/.
		'	"foo	= Put backups in folder where document is stored+sPath		.../documentdir/foo/.
End Function


'2 - Max number of timestamped backup files to be retained (per file).  0=no limit.  
Function GETiMaxCopies()	As Integer
	GETiMaxCopies	= 50							'	e.g. if have 10 and you set this to 8, then the 2 oldest backups will be auto deleted.
												'	All backup files need to be accessed to find the oldest, so huge values may be slow.
End Function


'3 - Expert adjustments.  Enable lines below if you often need to save your documents in non-native formats 
Function GETsB()			As Array
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
	
	GETsB	= Sb
End Function
