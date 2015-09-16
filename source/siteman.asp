<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
</SCRIPT>
<%
'  ** Copyright 1999, 2000 by John Martin d/b/a www.ANYPORTAL.com **
'  ** All Rights Reserved.                                        **
'  **                                                             **
'  ** This software is freeware and is not in the public domain.  **
'  ** You are hereby granted the right to freely distribute this  **
'  ** software as long as this copyright notice remains in place. **
'  **                                                             **
'  ** Comments or suggestions?   email: andmore@alief.com         **
'  **                                                             **
'  ** Date       Remarks                                          **
'  ** ---------  -----------------------------------------------  **
'  ** 25 MAY 99  original                                         **
'  ** 26 MAY 99  allow the script to run from a subdirectory      **
'  ** 27 MAY 99  increase security use of cookie                  **
'  ** 03 JUN 99  fix UNIX html file record endings                **
'  ** 07 JUN 99  fix spaces in file name problem                  **
'  ** 10 JUL 99  fix subdirectory problem with createimagetag     **
'  ** 10 JUL 99  add create document/folder logic                 **
'  ** 11 JUL 99  fix spaces in file name, again                   **
'  ** 11 JUL 99  .cfm & .php3 now edit like .asp/.html, etc.      **
'  ** 25 JUL 99  add interface to SA-FILEUP to upload files       **
'  ** 25 AUG 99  recode authorization routine, allow no password  **     
'  ** 31 AUG 99  some cosmetic; integrate with email community    **
'  ** 01 SEP 99  add link on detail page                          **
'  ** 05 SEP 99  add missing EndHTML on detail page               **
'  ** 24 OCT 00  plug /../ hole                                   **
'  ** 14 NOV 00  add Windows login security method                **
'  ** 14 NOV 00  convert in-line HTML to response.write           **
'  ** 14 NOV 00  improve shortcut parsing, clean-up link styles   **


	Option Explicit

	'universal variables (these undo the option explicit)

	Dim action
	Dim a,b,c,i,item,j
	Dim f,fso
	Dim arr,tstr

	'security

	Dim gblPassword
	gblPassword = "writeln"	'your password here
	              '^------ Null forces mandatory Windows login.

	'configuration

	Dim gblSiteName,gblSiteCode
	gblSiteName = Request.ServerVariables("SERVER_NAME")
	gblSiteCode = ""

	Dim gblNow 'server may not be local time
	gblNow = Now

	Dim gblFace,gblColor	'needs three quotes
	gblFace = """Arial, Helvetica, sans-serif"""
	gblColor = """#000066"""

	Dim gblRed
	gblRed = """#FF0000"""

	Dim gblReverse
	gblReverse = """#E0E0E0"""

	'global variables

	Dim gblTitle,gblPageText
	gblTitle = " * * * TITLE NOT SET * * * "
	gblPageText = "&nbsp;"

	'global constants

	Dim gblScriptName
	gblScriptName = Request.ServerVariables("Script_Name")
	gblScriptName = Mid(gblScriptName,InstrRev(gblScriptName,"/") + 1)

	Dim gblRoot
	gblRoot = Replace(Request.ServerVariables("Script_Name"),"/" & gblScriptName,"")

'-----------
'subprograms
'-----------

'--
'StartHTML
Sub StartHTML
	response.write "<HTML><HEAD><TITLE>" & gblSiteName & " " & gblTitle & "</TITLE>" & VBCRLF
	response.write "<META NAME=""description"" CONTENT=""AnyPortal"" " & gblTitle & ". " & gblSiteName & ">" & VBCRLF
	response.write "<META NAME=""keywords"" CONTENT=""anyportal, " & Lcase(gblTitle) & ", anyportal " & Lcase(gblTitle) & ", one file footprint, www.anyportal.com, andmore, the ANDMORE Companies, Houston, Texas, active server pages, ASP, asp, 100% ASP, 100% asp"">" & VBCRLF
	response.write "</HEAD>" & VBCRLF
	response.write "<BODY BGCOLOR=""#FFFFFF""><TABLE WIDTH=""100%"">" & VBCRLF
	response.write "<TR><TD ALIGN=""RIGHT"" VALIGN=""BOTTOM""><FONT COLOR=" & gblColor & " SIZE=3 FACE=" & gblFace & ">" & gblSiteName
	If Request.ServerVariables("LOGON_USER")="" Then
	Else
		response.write " (<FONT SIZE=1>USER:</FONT> " & Request.ServerVariables("LOGON_USER") & ")"
	End If
	response.write "</FONT></TD></TR>" & VBCRLF
	response.write "<TR><TD ALIGN=""LEFT"" VALIGN=""BOTTOM"" BGCOLOR=" & gblColor & "><FONT FACE=" & gblFace & " SIZE=4 COLOR=""#FFFFFF""><B>&nbsp;" & gblTitle & "</B></FONT></TD></TR>" & VBCRLF
	response.write "<TR><TD ALIGN=""LEFT"" VALIGN=""TOP""><FONT FACE=" & gblFace & " SIZE=2>" & gblPageText & "</FONT></TD></TR>" & VBCRLF
	response.write "</TABLE>" & VBCRLF
	response.write "<" & "!" & "-- begin " & gblScriptName & " --" & ">" & VBCRLF
	response.write "<" & "!" & "-- ---------------------------------------------------------- --" & ">" & VBCRLF
End Sub 'StartHTML

'--
'EndHTML
Sub EndHTML
	response.write "<" & "!" & "-- ---------------------------------------------------------- --" & ">" & VBCRLF
	response.write "<" & "!" & "-- end " & gblScriptName & " --" & ">" & VBCRLF
	response.write "<HR><FONT SIZE=1 FACE=" & gblFace & "><FONT COLOR=" & gblColor & " SIZE=3 FACE=" & gblFace & ">" & gblSiteName
	If Request.ServerVariables("LOGON_USER")="" Then
	Else
		response.write " (<FONT SIZE=1>USER:</FONT> " & Request.ServerVariables("LOGON_USER") & ")"
	End If
	response.write "</FONT><BR>" &  FormatDateTime(gblNow,1)  & " &nbsp; " &  FormatDateTime(gblNow,3)  & "" & VBCRLF
	response.write "<BR>AnyPortal " & gblTitle & " &copy; Copyright " & Year(gblNow) & " by <A TITLE=""www.anyportal.com is a project of the ANDMORE Companies -- Houston, Texas"" HREF=""http://www.anyportal.com"">www.AnyPortal.com</A><BR></FONT>" & VBCRLF
	response.write "</BODY></HTML>" & VBCRLF
	response.write VBCRLF
End Sub 'EndHTML

'--
' Authorize
Function Authorize
Dim a,i,pw
	If _
	(gblPassword="") OR _
	(Request.Cookies(gblSiteCode & gblScriptName)=Condensation(SStr(gblPassword))) OR _
	Request.ServerVariables("LOGON_USER")<>"" _
	Then
		Authorize = TRUE
	Else
		If Request.QueryString("w")="y" AND Request.ServerVariables("LOGON_USER")="" Then
			Response.Status = "401 Access Denied"
			StartHTML
			response.write "<BLOCKQUOTE><FONT FACE=" & gblFace & " SIZE=5>"
			response.write "<FONT COLOR=""##FF0000""><B>Access denied.</B></FONT><FONT SIZE=2>"
			response.write "<BR>Sorry, but the username/password you supplied<BR> was not recognized by the <A HREF=""http://" & gblSiteName & """>" & gblSiteName & "</A> web site " & VBCRLF
			response.write "<P>Contact your web site administrator for more information." & VBCRLF
			response.write "</FONT></FONT></BLOCKQUOTE>" & VBCRLF
			EndHTML
			Response.End
		End If
		Authorize = FALSE
		pw = Request.Form("password")
		a = Condensation(pw)
		If pw<>"" OR Request.Form("OK")<>"" Then
			If pw = gblPassword Then
				'cookie expires when browser is closed...
				Response.Cookies(gblSiteCode & gblScriptName) = a
				'set a permanent one to never see this page again
				If Request.Form("SAVE") = "on" Then Response.Cookies(gblSiteCode & gblScriptName).Expires = gblNow+30
				Response.Redirect gblScriptName & "?d="
			Else
				gblPageText = gblPageText & "<FONT TITLE=""Sorry. That's not the password. Try again."" COLOR=" & gblRed & "><B>Invalid password.</B></FONT>"
			End If
		End If
		If Request.ServerVariables("SERVER_SOFTWARE")>="Microsoft-IIS/4.0" Then
			StartHTML
			response.write "<FORM METHOD=""POST"" ACTION=""" & gblScriptName & """><BLOCKQUOTE><TABLE CELLPADDING=5>" & VBCRLF
			response.write "<TR>" & VBCRLF
			response.write "<TD><FONT TITLE=""The password method uses cookies to secure this site. For the correct password, contact the web site administrator."" FACE=" & gblFace & " SIZE=1>PASSWORD:</FONT>" & VBCRLF
			response.write "<INPUT TYPE=""PASSWORD"" SIZE=17 NAME=""Password""></TD>" & VBCRLF
			response.write "<TD BGCOLOR=" & gblReverse & "><FONT FACE=" & gblFace & " SIZE=1 TITLE=""Check this box to save a cookie in the browser of this machine. You won't have to log-in again for the next 30 days.""> &nbsp; SAVE COOKIE?</FONT>" & VBCRLF
			response.write "<INPUT TYPE=""CHECKBOX"" NAME=""SAVE""></TD>" & VBCRLF
			response.write "<TD><INPUT TYPE=""SUBMIT"" NAME=""OK"" VALUE=""ENTER""></TD>" & VBCRLF
			response.write "</TR>" & VBCRLF
			response.write "<TR><TD COLSPAN=3>" 
			response.write "<FONT FACE=""Wingdings"" SIZE=6 COLOR=""#000000"">" & chr(255) & "</FONT><FONT TITLE=""The login method uses your Windows username and password to secure this site."" FACE=" & gblFace & " SIZE=3> Use Windows <A HREF=""" & gblScriptName & "?w=y"">login</A>.</FONT></TR>" & VBCRLF
			response.write "</TABLE></BLOCKQUOTE></FORM>" & VBCRLF
			response.write VBCRLF
		Else
			gblPageText = "Your web server identified itself as """ & Request.ServerVariables("SERVER_SOFTWARE") & """."
			StartHTML
			response.write "<BLOCKQUOTE><FONT FACE=" & gblFace & " SIZE=5><B>Sorry.</B><P>" & VBCRLF
			response.write "AnyPortal " & gblTitle & " requires Microsoft NT/2000 Internet Information Server (IIS) 4.0 or greater." & VBCRLF
			response.write "</FONT></BLOCKQUOTE>" & VBCRLF
		End If
		EndHTML
	End If
End Function 'Authorize

'--
' Condensation
Function Condensation(s)
	a = 0
	For i = 1 to len(s)
		a = (ASC(mid(s,i,1)) + a*2) Mod 77411
	Next 'i
	Condensation = Right("00000" & Cstr(a),5) & Right("00000" & Cstr((len(s)*23)+25433),5)
End Function 'Condensation(s)

'--
' CreateImageTag
Function CreateImageTag(fn,altstr,align,border)
Dim f,fso,pn
Dim tstr,alignstr,borderstr
Dim chars,hw,width,height

	If border="" Then
		borderstr = " BORDER=0"
	Else
		borderstr = " BORDER=" & Cstr(border)
	End If
	If align="" Then
		alignstr = ""
	Else
		alignstr = " ALIGN=""" 
		Select Case UCase(left(align,1))
		Case "L"
			tstr = "LEFT"
		Case "R"
			tstr = "RIGHT"
		Case "C"
			tstr = "CENTER"
		Case Else
		End Select
		alignstr = " ALIGN=""" & tstr & """"
	End If		

	Set fso = CreateObject("Scripting.FileSystemObject")
	pn = Server.MapPath(fn)
	tstr = ""
	Set f = fso.OpenTextFile(pn)

	Select Case UCase(Right(fn,4))
	Case ".GIF",".JPG"
		If NOT f.AtEndOfStream Then
			If UCase(Right(fn,4))=".GIF" Then 'always works
				chars		= f.read(10)
				width		= asc(mid(chars,8,1))*256 + asc(mid(chars,7,1))
				height	= asc(mid(chars,10,1))*256 + asc(mid(chars,9,1))
				hw = " WIDTH=" & width & " HEIGHT=" & height
			Else 'usually works
				chars		= f.read(200)
				height	= asc(mid(chars,164,1))*256 + asc(mid(chars,165,1))
				width		= asc(mid(chars,166,1))*256 + asc(mid(chars,167,1))
				If (height>600) OR (height<3) OR (WIDTH<3) OR (WIDTH>600) Then
					'could be wrong height, width... forget 'em
				Else
					hw = " WIDTH=" & width & " HEIGHT=" & height
				End If
			End If
		End If
		tstr = "<IMG SRC=""" & Replace(Replace(fn,"\","/")," ","%20") & """" & hw & borderstr & alignstr & " ALT=""" & altstr & """>"
	End Select
	f.Close
	Set f = Nothing
	Set fso = Nothing
	CreateImageTag = tstr
End Function 'CreateImageTag

'--
' DetailPage
Sub DetailPage
Dim chars,fstr,hw,height,width
Dim IsTextFile,pathname
Dim fsize,fdatecreated,fdatelastmodified

	pathname = fsDir & fn
	If right(pathname,1) = "\" Then pathname = Left(pathname,len(pathname)-1)
	
	' create if you gotta
	If fso.FileExists(pathname) Then
	Else
		Select Case UCase(Request.QueryString("T"))
		Case "D" 'create document
			Set f = fso.CreateTextFile(pathname)
			f.Close
			Set f= Nothing
		Case "F" 'create folder
			Set f = fso.CreateFolder(pathname)
			pathname = pathname & "\"
			response.redirect gblScriptName & "?d=" & URLSpace(pathname)
		End Select
	End If
	
	StartHTML
	response.write "<P><FONT FACE=""Andale Mono, Monotype.com, Courier New, Courier, sans-serif"" SIZE=4><B>" & pathname & "</B><BR>" & VBCRLF
	response.write "<A HREF=""" & webbase & fn & """>" & webbase & fn & "</A><BR></FONT>" & VBCRLF
	
	If fso.FileExists(pathname) Then 
		' fetch Window's file information 
		Set f = fso.GetFile(pathname) 
		fsize = f.size 
		fdatecreated = f.datecreated 
		fdatelastmodified = f.datelastmodified
		response.write "<PRE>" & VBCRLF
		response.write "    file size:  " & FormatNumber(fsize,0) & " characters" & VBCRLF
		response.write " file created: &nbsp;<B>" & FormatDateTime(fdatecreated,1) & " </B>&nbsp;" & FormatDateTime(fdatecreated,3) & VBCRLF
		response.write "last modified: &nbsp;<B>" & FormatDateTime(fdatelastmodified,1) & " </B>&nbsp;" & FormatDateTime(fdatelastmodified,3) & VBCRLF
		response.write "</PRE>" & VBCRLF
		Set f = Nothing
	End If
	
	response.write "<FORM ACTION=""" & gblScriptName & """ METHOD=""POST"">" & VBCRLF
	response.write "<INPUT TYPE=""HIDDEN"" NAME=""fsDIR"" VALUE=""" & fsDir & """>" & VBCRLF
	
	IsTextFile = FALSE
	Select Case UCase(Right(fn,4))
	Case ".GIF",".JPG"
		tstr = CreateImageTag(basedir & fn,fn & " (" & FormatNumber(Int(fsize/1024*10+.05)/10,1) & " Kb)","",0)
		response.write "<FONT FACE=""Andale Mono, Monotype.com, Courier New, Courier, sans-serif"" SIZE=2>"
		response.write Server.HTMLEncode(tstr) & "</FONT><BR><BR>" & tstr & "<P>" & VBCRLF
	Case ".URL"
		Set f = fso.OpenTextFile(pathname)
		If NOT f.AtEndOfStream Then tstr = f.readall
		f.Close
		Set f = Nothing
		response.write "<FONT COLOR=""#3333FF"" FACE=""Andale Mono, Monotype.com, Courier New, Courier, sans-serif"" SIZE=2>" & VBCRLF
		response.write Replace(Server.HTMLEncode(tstr),VBCRLF,VBCRLF & "<BR>")
		response.write "</FONT>" & VBCRLF
	Case ".TXT",".ASA",".ASP",".HTM","HTML",".CFM","PHP3"
		'read the file
		Set f = fso.OpenTextFile(pathname)
		If NOT f.AtEndOfStream Then fstr = f.readall
		f.Close
		Set f = Nothing
		Set fso = Nothing
		IsTextFile = TRUE
		response.write "<TABLE BGCOLOR=" & gblReverse & "><TR><TD>" & VBCRLF
		response.write "<FONT TITLE=""Use this text area to view or change the contents of this document. Click [SAVE] to store the updated contents to the web server."" FACE=" & gblFace & "SIZE=1><B>DOCUMENT CONTENTS</B></FONT><BR>" & VBCRLF
		response.write "<TEXTAREA NAME=""FILEDATA"" ROWS=18 COLS=70 WRAP=""OFF"">" & Server.HTMLEncode(fstr) & "</TEXTAREA>" & VBCRLF
		response.write "</TD></TR></TABLE>" & VBCRLF
	End Select
	response.write VBCRLF & "<BR><BR>" & VBCRLF
	If IsTextFile Then
		response.write "<INPUT TYPE=""TEXT"" SIZE=48 MAXLENGTH=255 NAME=""PATHNAME"" VALUE=""" & pathname & """>" & VBCRLF
		response.write "<INPUT TYPE=""RESET"" VALUE=""RESET""> <INPUT TYPE=""SUBMIT"" NAME=""POSTACTION"" VALUE=""SAVE"">" & VBCRLF
		response.write "<INPUT TYPE=""SUBMIT"" NAME=""POSTACTION"" VALUE=""CANCEL""><BR>" & VBCRLF
	Else
		response.write "<INPUT TYPE=""HIDDEN"" NAME=""PATHNAME"" VALUE=""" & pathname & """>" & VBCRLF
		response.write "<INPUT TYPE=""SUBMIT"" NAME=""POSTACTION"" VALUE=""BACK""><BR>" & VBCRLF
	End If
	response.write "<HR><FONT TITLE=""Check OK and click [DELETE] to delete this document from the web server. (Cannot be undone.)"" FACE=" & gblFace & "SIZE=1><B>OK TO DELETE """ & UCase(fn) & """? </B></FONT>" & VBCRLF
	response.write "<INPUT TYPE=""CHECKBOX"" NAME=""DELETEOK"">" & VBCRLF
	response.write "<INPUT TYPE=""SUBMIT"" NAME=""POSTACTION"" VALUE=""DELETE"">" & VBCRLF
	response.write "</FORM>" & VBCRLF
	EndHTML
End Sub 'DetailPage

'--
' DisplayCode
Sub DisplayCode
Dim fn,fso,f
Dim code,tstr
Dim a,arr,i

	fn = Request.QueryString("c")
	response.write "<HTML><HEAD><TITLE>" & fn & "</TITLE></HEAD><BODY>" & VBCRLF
	response.write "<STYLE>" & VBCRLF
	response.write "<!" & "--" & VBCRLF
	response.write "SPAN{color:Navy;background-color:Yellow}" & VBCRLF
	response.write "--" & ">" & VBCRLF
	response.write "</STYLE>" & VBCRLF

	If Instr(fn,fsroot)=1 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(fn, 1, 0, 0)
		If f.AtEndOfStream Then
			code = ""
		Else
			code = f.ReadAll					'totally unconverted
		End If
		'quickly format code for readability...
		' could be smarter, but it sure is simple!
		tstr = Server.HTMLEncode(code)
		tstr = Replace(tstr,chr(9),"   ")
		tstr = Replace(tstr,"  ","&nbsp;&nbsp;")
		tstr = Replace(tstr,"&lt;%","<SPAN>&lt;" & "%</SPAN><FONT COLOR=""#000000"">")
		tstr = Replace(tstr,"%&gt;","<SPAN>%" & "</FONT>&gt;</SPAN>")
		tstr = Replace(tstr,"&lt;!--","<I><FONT COLOR=""#CC0033"">&lt;!--")
		tstr = Replace(tstr,"--&gt;","--&gt;</I></FONT>")

		response.write "<TABLE WIDTH=""100%"" BGCOLOR=" & gblColor & "><TR><TD><FONT COLOR=""#FFFFFF"" FACE=""Andale Mono, Monotype.com, Courier New, Courier, sans-serif"" SIZE=5><B>" & VBCRLF
		response.write "&nbsp;" & fn & "</B></FONT></TD></TR></TABLE>" & VBCRLF

		response.write "<FONT COLOR=""#0000FF"" FACE=""Andale Mono, Monotype.com, Courier New, Courier, sans-serif"" SIZE=2>" & VBCRLF
		response.write "<!" & "-- code listing --" & ">" & VBCRLF & VBCRLF
		arr = Split(Replace(tstr,chr(13),""),chr(10)) 'handle unix/linux files, too
		For i = 0 to UBound(arr)
			'add line numbers and output
			response.write "<BR><FONT COLOR=""#008000"">" & Right("000" & i+1,4) & ":</FONT> "
			tstr = arr(i)
			If left(Replace(Replace(tstr,"&nbsp;","")," " ,""),1) = "'" Then
				response.write "<FONT COLOR=""#CC0033""><I>" & tstr & "</I></FONT>" & VBCRLF
			Else
				response.write tstr & VBCRLF
			End If
		Next 'i
		response.write VBCRLF & "<!" & "-- end of code listing --" & ">" & VBCRLF
		response.write "</FONT>" & VBCRLF 
	Else
		response.write "<P><FONT COLOR=""#CC0033"" SIZE=3>Cannot access " & fn & "</FONT>" & VBCRLF
	End If
	response.write "<HR></BODY></HTML>"
End Sub 'DisplayCode

'--
' DisplayFileName
Sub DisplayFileName(dirfile,fhandle)
Dim newgif,linktarget
Dim fsize

	response.write "<TR>" & VBCRLF
	If dirFile="DIR" Then
		linktarget = "<A HREF=""" & gblScriptName & "?d=" & URLSpace(fhandle) & "\"" TITLE=""Click here to move down a level and list the documents in this folder."">"
		tstr = "<FONT FACE=" & gblFace & " SIZE=2>" & linktarget & LCase(fhandle.name) & "</A></FONT>"
		response.write "<TD VALIGN=""TOP"" ALIGN=""RIGHT"">" & MockIcon("fldr") & "</TD>" & VBCRLF
		response.write "<TD COLSPAN=3 VALIGN=""TOP"" BGCOLOR=" & gblReverse & ">" & Tstr & "</TD>" & VBCRLF
	Else
		newgif = ""
		If fhandle.datelastmodified+14>gblNow Then newgif = MockIcon("newicon")
		b = ""
		If len(fhandle.name)>4 Then b = Ucase(Right(fhandle.name,4))
		If Left(b,1) = "." Then b = Right(b,3)
		Select Case b
		Case "ASP","HTM","HTML","ASA","TXT","CFM","PHP3"
			newgif = newgif & " <A TARGET=""_blank"" HREF=""" & gblScriptName & "?c=" & URLSpace(fsDir & fhandle.name) &  """ TITLE=""Click here to list the contents of this document."" STYLE=""{text-decoration:none}"">" & MockIcon("view") & "</A>"
			tstr = webbase & replace(fhandle.name," ","%20")
		Case "URL"
			tstr = ShortCutURL
		Case Else
			tstr = webbase & replace(fhandle.name," ","%20")
		End Select
		If fhandle.size<10240 Then
			If fhandle.size=0 Then
				fsize = "0"
			Else
				fsize = FormatNumber(fhandle.size,0,0,-2)
			End If
		Else
			fsize = FormatNumber((fhandle.size+1023)/1024,0,0,-2) & "K"
		End If
		tstr = "<FONT FACE=" & gblFace & " SIZE=2><A HREF=""" & tstr & """ TITLE=""Click here to link to this document."">" & LCase(fhandle.name) & "</A></FONT>" & newgif
		response.write "<TD VALIGN=""TOP"" ALIGN=""RIGHT""><A HREF=""" & gblScriptName & "?f=" & URLSpace(fhandle.name) & "&d=" & URLSpace(fsDir) & """ TITLE=""Click here to view more details about this document."" STYLE=""{text-decoration:none}"">" & MockIcon(b) & "</A></TD>" & VBCRLF
		response.write "<TD VALIGN=""TOP"" BGCOLOR=" & gblReverse & ">" & Tstr & "</TD>" & VBCRLF
		response.write "<TD VALIGN=""TOP"" BGCOLOR=" & gblReverse & "><FONT FACE=" & gblFace & " SIZE=1>" & FormatDateTime(fhandle.datelastmodified,0) & "</FONT></TD>" & VBCRLF
		response.write "<TD VALIGN=""TOP"" BGCOLOR=" & gblReverse & "><FONT FACE=" & gblFace & " SIZE=1>" & fsize & " bytes</FONT></TD>" & VBCRLF
	End If 
	response.write "</TR>" & VBCRLF
End Sub 'DisplayFileName

'--
' MockIcon (icon emulator)
Function MockIcon(txt)
Dim tstr,d

	'Sorry, mac/linux users.
	tstr = "<FONT FACE=""WingDings"" SIZE=4 COLOR=" & gblRed & ">"
	Select Case Lcase(txt)
	Case "bmp","gif","jpg","tif","jpeg","tiff"
		d = 176
	Case "doc"
		d = 50
	Case "exe","bat","bas","c","src"
		d = 255
	Case "file"
		d = 51
	Case "fldr"
		d = 48
	Case "htm","html","asa","asp","cfm","php3"
		d = 182
	Case "pdf"
		d = 38
	Case "txt","ini"
		d = 52
	Case "xls"
		d = 252
	Case "zip","arc","sit"
		d = 59
	Case "newicon"
		tstr = "<FONT TITLE=""This document has been modified sometime during the last 14 days."" FACE=""WingDings"" SIZE=4 COLOR=" & gblRed & ">"
		d = 171
	Case "view"
		d = 52
	Case Else
		d = 51
	End Select
	tstr = tstr & Chr(d) & "</FONT>"
	MockIcon = tstr
End Function 'mockicon

'--
' Navigate
Sub Navigate
Dim emptyDir

	emptyDir = TRUE
	response.write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=3 WIDTH=""100%"">"

	' get the directory of file names
	If toplevel Then
		parent = ""
	Else
		parent = fso.GetParentFolderName(fsDir) & "\"
		response.write "<TR><TD VALIGN=""TOP"" ALIGN=""RIGHT""><FONT FACE=""WingDings"" SIZE=4 COLOR=" & gblRed & ">" & chr(199) & "</FONT></TD>" & VBCRLF
		response.write "<TD COLSPAN=3><FONT FACE=" & gblFace & " SIZE=1><B><A TITLE=""Click here to move up a level to the parent folder."" HREF=""" & gblScriptName & "?d=" & URLSpace(parent) & """>" & UCASE(fso.GetParentfolderName(fsDir) & "\") & "</A></B></FONT></TD></TR>" & VBCRLF
	End If
	Set f = fso.GetFolder(fsDir)
	Set FileList = f.subFolders
	a = 0
	For Each fn in FileList
		emptyDir = FALSE
		If a = 0 Then
			a = 1
			response.write "<TR><TD VALIGN=""TOP"">&nbsp;</TD>" & VBCRLF
			response.write "<TD COLSPAN=3><HR><FONT FACE=" & gblFace & " SIZE=4><B>Additional Folders</B></FONT></TD>" & VBCRLF
			response.write "</TR>" & VBCRLF
			response.write "<TR><TD VALIGN=""TOP"">&nbsp;</TD>" & VBCRLF
			response.write "<TD COLSPAN=3 VALIGN=""BOTTOM""><FONT FACE=" & gblFace & " COLOR=" & gblRed & " SIZE=1><B>FOLDER NAME</B></FONT></TD>" & VBCRLF
			response.write "</TR>" & VBCRLF
		End If
		DisplayFileName "DIR",fn
	Next 'fn

	response.write "<TR><TD VALIGN=""TOP"">&nbsp;</TD>" & VBCRLF
	response.write "<TD COLSPAN=3><HR><FONT FACE=" & gblFace & " SIZE=4><B>" & fsDir & "</B></FONT></TD>" & VBCRLF
	response.write "</TR>" & VBCRLF
	response.write "<TR><TD VALIGN=""TOP"">&nbsp;</TD>" & VBCRLF
	response.write "<TD VALIGN=""BOTTOM""><FONT FACE=" & gblFace & " COLOR=" & gblRed & " SIZE=1><B>DOCUMENT NAME</B></FONT></TD>" & VBCRLF
	response.write "<TD VALIGN=""BOTTOM""><FONT FACE=" & gblFace & " COLOR=" & gblRed & " SIZE=1><B>LAST UPDATE</B></FONT></TD>" & VBCRLF
	response.write "<TD VALIGN=""BOTTOM""><FONT FACE=" & gblFace & " COLOR=" & gblRed & " SIZE=1><B>FILE SIZE</B></FONT></TD>" & VBCRLF
	response.write "</TR>" & VBCRLF
	response.write "" & VBCRLF

	Set filelist = f.Files
	For Each fn in filelist
		emptyDir = FALSE
		DisplayFileName "FILE",fn
	Next 'fn

	If emptyDir Then
		response.write "  <FORM METHOD=""POST"" ACTION=""" & gblScriptName & """>" & VBCRLF
		response.write "  <TR><TD></TD><TD COLSPAN=3 VALIGN=""BOTTOM"" BGCOLOR=" & gblReverse & ">" & VBCRLF
		response.write "  <INPUT TYPE=""HIDDEN"" NAME=""PARENT"" VALUE=""" & parent & """>" & VBCRLF
		response.write "  <INPUT TYPE=""HIDDEN"" NAME=""PATHNAME"" VALUE=""" & fsDir & """>" & VBCRLF
		response.write "  <FONT FACE=" & gblFace & " SIZE=1> &nbsp; OK TO DELETE THIS EMPTY FOLDER? </FONT>" & VBCRLF
		response.write "  <INPUT TYPE=""CHECKBOX"" NAME=""OK""> &nbsp;" & VBCRLF
		response.write "  <INPUT TYPE=""SUBMIT"" NAME=""POSTACTION"" VALUE=""DELETE"">" & VBCRLF
		response.write "  </TD></TR></FORM>" & VBCRLF
	End If
	response.write "<TR><TD></TD><TD COLSPAN=3><HR></TD></TR>" & VBCRLF
	response.write "  <FORM METHOD=""GET"" ACTION=""" & gblScriptName & """>" & VBCRLF
	response.write "  <TR><TD></TD><TD COLSPAN=3 VALIGN=""BOTTOM"" BGCOLOR=" & gblReverse & ">" & VBCRLF
	response.write "  <FONT FACE=" & gblFace & " SIZE=1> &nbsp; CREATE NEW </FONT>" & VBCRLF
	response.write "  <INPUT TYPE=""RADIO"" NAME=""T"" VALUE=""D"" CHECKED><FONT FACE=" & gblFace & " SIZE=1>DOCUMENT</FONT>" & VBCRLF
	response.write "  <FONT FACE=" & gblFace & " SIZE=1> -OR- </FONT>" & VBCRLF
	response.write "  <INPUT TYPE=""RADIO"" NAME=""T"" VALUE=""F""><FONT FACE=" & gblFace & " SIZE=1>FOLDER:</FONT> &nbsp;" & VBCRLF
	response.write "  <FONT FACE=" & gblFace & " SIZE=1> &nbsp; NAME </FONT> &nbsp;" & VBCRLF
	response.write "  <INPUT TYPE=""TEXT"" NAME=""F"" SIZE=14> &nbsp;" & VBCRLF
	response.write "  <INPUT TYPE=""HIDDEN"" NAME=""D"" VALUE=""" & fsDir & """>" & VBCRLF
	response.write "  <INPUT TYPE=""SUBMIT"" VALUE=""CREATE"">" & VBCRLF
	response.write "  <NOBR><FONT FACE=" & gblFace & " SIZE=1> &nbsp; OR <A HREF=""" & gblScriptName & "?u=Y&d=" & URLSpace(fsDir) & """>UPLOAD</A> USING SA-FILEUP</FONT></NOBR>" & VBCRLF
	response.write "  </TD></TR></FORM>" & VBCRLF
	response.write "</TABLE>" & VBCRLF
End Sub 'Navigate

'--
' ShortCutURL
Function ShortCutURL
Dim f,fstr,tstr
	tstr = ""
	Set f = fso.OpenTextFile(fn)
	Do While NOT f.AtEndOfStream 
		tstr = f.readline
		If len(tstr)<7 Then
		Else
			If left(lcase(tstr),4)="url=" Then
				fstr = tstr
			End If
		End If
	Loop
	f.Close
	Set f= Nothing
	If fstr = "" Then
		ShortCutURL = fn
	Else
		ShortCutURL = Replace(mid(fstr,5,255)," ","%20")
	End If
End Function 'ShortCutURL

'--
' SStr (force null to "")
Function SStr(v)
Dim rt
	If IsNull(v) Then 
		rt = ""
	Else
		rt = Trim(Cstr(v))
	End If
	SStr = rt
End Function 'sstr


'--
' UploadPage
Sub UploadPage
	StartHTML
	response.write "<P><TABLE BORDER=0 CELLPADDING=5><TR><TD WIDTH=5></TD><TD BGCOLOR=" & gblReverse & " VALIGN=""""TOP"""">" & VBCRLF
	response.write "<FORM ENCTYPE=""multipart/form-data"" METHOD=""POST"" ACTION=""" & gblScriptName & "?u=D&d=" & URLSpace(fsDir) & """>" & VBCRLF
	response.write "<FONT SIZE=1 FACE=" & gblFace & ">NAME OF DESTINATION FOLDER ON WEB SITE</FONT><BR>" & VBCRLF
	response.write "<FONT SIZE=4 FACE=" & gblFace & "><B>" & fsDir & "</B></FONT><P>" & VBCRLF
	response.write "<FONT SIZE=1 FACE=" & gblFace & ">PATHNAME OF LOCAL DOCUMENT<BR>(SEND THIS FILE TO THE WEB SERVER)</FONT><BR><INPUT SIZE=30 TYPE=""FILE"" NAME=""F1""><P>" & VBCRLF
	response.write "<INPUT TYPE=""SUBMIT"" VALUE=""UPLOAD"">" & VBCRLF
	response.write "<P><FONT SIZE=2 FACE=" & gblFace & ">If the <B>[BROWSE...]</B> button is not displayed," & VBCRLF
	response.write "<BR>you must upgrade your <A HREF=""http://www.netscape.com"">Netscape</A>" & VBCRLF
	response.write "or <A HREF=""http://www.microsoft.com"">Microsoft</A> browser." & VBCRLF
	response.write "</FORM></TD>" & VBCRLF
	response.write "<TD VALIGN=""TOP""><FONT SIZE=2 FACE=" & gblFace & ">" & VBCRLF
	response.write "<P>Your browser:<BR>HTTP_USER_AGENT: " & Request.ServerVariables("HTTP_USER_AGENT") & "" & VBCRLF
	response.write "<P>Upload also requires that <A TARGET=""_blank"" HREF=""http://www.softartisans.com"">the SA-FileUp object</A> is registered on your web server." & VBCRLF
	response.write "<BR>(Some object is <B>always</B> required for uploads<I>!!!</I>)" & VBCRLF
	response.write "</FONT>" & VBCRLF
	response.write "<FORM METHOD=""POST"" ACTION=""" & gblScriptName & """>" & VBCRLF
	response.write "<INPUT TYPE=""HIDDEN"" NAME=""fsDir"" VALUE=""" & fsDir & """><BR>" & VBCRLF
	response.write "<FONT SIZE=2 FACE=" & gblFace & ">DON'T USE SA-FILEUP?<BR>SORRY! CLICK HERE...</FONT><BR>" & VBCRLF
	response.write "<INPUT TYPE=""SUBMIT"" NAME=""POSTACTION"" VALUE=""CANCEL"">" & VBCRLF
	response.write "</FORM>" & VBCRLF
	response.write "</TD></TR></TABLE><P>" & VBCRLF
	EndHTML
End Sub 'UploadPage

'--
' URLspace
Function URLSpace(s)
	URLSpace = replace(replace(s,"+","%2B")," ","+")
End Function 'URLSpace

'----
'MAIN
'----
Dim filelist,fn,upl
Dim TextObject,fhandle,lsplit

Dim fsDir,baseDir,webbase
Dim fsRoot,webRoot
Dim pathname,parent,toplevel

	gblTitle = "Site Manager"

	If NOT Authorize Then
		'function will output HTML for password
	Else
		'initialization
	
		Set fso = CreateObject("Scripting.FileSystemObject")
	
		'dynamically find out where the documents and web pages are located
	
		fsDir = replace(LCase(Request.QueryString("d")),"/../","/")
		If fsDir = "" Then fsDir = Request.Form("fsDir")
		fsRoot = LCase(Replace(Server.MapPath(gblScriptName),"\" & gblScriptName,"") & "\")
		If Instr(fsdir,fsroot) <> 1 Then fsDir = fsRoot
		If Lcase(fsDir) = Lcase(fsRoot) Then toplevel = TRUE
		basedir = Replace(Mid(fsDir,len(fsRoot),250),"\","/")
		webRoot = "http://" & Request.ServerVariables("SERVER_NAME") & Replace(Request.ServerVariables("SCRIPT_NAME"),"/" & gblScriptName,"")
		webbase = replace(webroot & basedir," ","%20")

		'process a GET/POST request
	
		If Request.QueryString("u")="D" Then
			Action = "UPLOAD"
		Else
			Action = Request.Form("POSTACTION")
			pathname = Request.Form("PATHNAME")
		End If
		Select Case UCase(Action)
		Case "UPLOAD"
			Set upl = Server.CreateObject("SoftArtisans.FileUp")
			tstr = Mid(upl.UserFilename, InstrRev(upl.UserFilename, "\") + 1)
			If tstr = "" Then
			Else
				upl.SaveAs fsdir & tstr
			End If
		Case "SAVE"
			Select Case UCase(Right(pathname,4))
			Case ".TXT",".ASA",".ASP",".HTM","HTML",".CFM","PHP3"
				If Instr(pathname,fsroot) = 1 Then
					Set f = fso.CreateTextFile(pathname)
					f.write Request.Form("FILEDATA")
					f.close 
				End If
			End Select
		Case "DELETE" 'either document or folder
			If Request.Form("OK") = "on" Then
				parent = Request.Form("Parent")
				If Instr(pathname,fsroot) = 1 Then
					fso.DeleteFolder Left(pathname,Len(pathname)-1),TRUE
					response.redirect gblScriptName & "?d=" & URLSpace(parent)
				End If
			End If
			If Request.Form("DELETEOK") = "on" Then
				If Instr(pathname,fsroot) = 1 Then
			      If fso.FileExists(Request.Form("PathName")) Then
			         Set f = fso.GetFile(Request.Form("PathName"))
		   	      f.delete
		      	End If
				End If
			End If
		End Select
		If Action <> "" Then
			tstr = gblScriptName & "?d="
			If NOT toplevel Then	tstr = tstr & URLSpace(fsDir)
			response.redirect tstr
		End If
	
		'check for mode... navigate, code display, upload, or detail?
	
		fn = LCase(Request.QueryString("f"))
		If fn="" Then
			If Request.QueryString("u") = "Y" Then
				gblTitle = gblTitle & " (Upload Page)"
				gblPageText = "Use this page to upload a single document to this web site."
				UploadPage
			Else
				If Request.QueryString("c") = "" Then
					gblPageText	= "Use this page to add, delete or revise documents on this web site."
					StartHTML
					Navigate
					EndHTML
				Else
					DisplayCode
				End If
			End If
		Else
			gblTitle = gblTitle & " (Detail Page)"
			gblPageText = "Use this page to view, modify or delete a single document on this web site."
			DetailPage
		End If
	End If
%>