<%@ Language=VBScript%>
<%
Option Explicit
On Error Resume Next

Dim x, y, z
Dim intID, intTD
Dim	FullPath, IconFolder, AppPath
Dim strDetail, strList
Dim blnExpand, blnDownload, blnIcon, blnHidden, blnIconFolder
Dim strAttributes, strCompressed, strHidden, strInitDirectory, lIsTableLinesInAlternateColors
Dim bShowRowHeader, bShowNameColumn, bShowSizeColumn, bShowAttributesColumn, bShowLastModifiedColumn
Dim sRootNodeCaption, bShowLastSpacingColumn, lMaxFileNameLengthForDisplay

'Tabed to show relation.
Dim	FileSystem
Dim		Drives
Dim			Drive
Dim		Folders
Dim			Folder
Dim			SubFolders
Dim				SubFolder
Dim			Files
Dim				File

'------------------------------------------
' Author: BenWhite@columbus.rr.com
'
' Preferences:
'
' Initial Directory (Example: "C:\Temp")
	strInitDirectory = "..\links\browse"
'
' Enable/Disable Downloads (may be slow For large files)
	blnDownload = True
'
' Endable/Disable View of Hidden files/folders
 	blnHidden = False
'
' Endable/Disable Dynamic Icon Loading (this sometimes fails due to server connection limitations)
 	blnIcon = False
'Show table lines in alternate color
	lIsTableLinesInAlternateColors = 2	'1=Yes, 2=No
'Should too long file names to be limited
	lMaxFileNameLengthForDisplay = 95
'Which columns to show?
	bShowRowHeader = False
	bShowNameColumn = True
	bShowSizeColumn = False
	bShowAttributesColumn = False
	bShowLastModifiedColumn = False
	bShowLastSpacingColumn = False
'Captions of root nodes
	sRootNodeCaption = "Shital's Handpicked Links"

'
' Note:
'  If you want to avoid possible download abuse.
'  You should delete the download.asp file
'------------------------------------------

'Main folder generation routine
Private Sub GetDetail()
	On Error Resume Next
	Dim OpenLink, OriginalID, FileSize, SubdirCount, DirSize, FreeSpace, ImageFile, VirtualPath

	strInitDirectory = Left(strInitDirectory, instrrev(strInitDirectory,"\",Len(strInitDirectory)-1,1))
	FullPath = strInitDirectory & FullPath

	FreeSpace = FileSystem.GetDrive(FileSystem.GetDriveName(FullPath)).FreeSpace
	Set Folder = FileSystem.GetFolder(FullPath)
	Set SubFolders = Folder.SubFolders
	Set Files = Folder.Files

	x=0
	z=0
	OriginalID = intID
	blnIconFolder = LCase(IconFolder) = LCase(Folder.Path & "\")

	If LCase(strInitDirectory) <> LCase(Left(FullPath & "\",Len(strInitDirectory))) Then
		response.write "alert('Error: Permission denied');" & vbcrlf
		Set Files = Nothing
		Set SubFolders = Nothing
		Set Folder = Nothing
		Exit Sub
	End If

	strList = "<table cellpadding=0 cellspacing=0 border=0 class=SmallText>"
	strDetail = "<table bgcolor=#ffffff cellpadding=0 cellspacing=0 border=0 height=100% width=100% ><tr><td valign=top>" & _
		"<table cellpadding=0 cellspacing=0 border=0 width=100% style='padding:0 7;' class=SmallText>"
	
	If bShowRowHeader = True Then
		strDetail = strDetail &	"<tr>"
		If bShowNameColumn = True Then	strDetail = strDetail &		"<td class=ColumnTitle colspan=2>Name</td>" 
		If bShowSizeColumn = True Then	strDetail = strDetail &		"<td class=ColumnTitle align=right>Size</td>"
		If bShowLastModifiedColumn = True Then	strDetail = strDetail &		"<td class=ColumnTitle >Modified</td>" 
		If bShowAttributesColumn = True Then	strDetail = strDetail &		"<td class=ColumnTitle >Attributes</td>" 
		If bShowLastSpacingColumn = True Then	strDetail = strDetail &		"<td class=ColumnTitle width=100% >&nbsp;</td>"
		strDetail = strDetail &		"</tr>"
	End If
	
	For Each SubFolder In SubFolders
		GetAttributes SubFolder.Attributes
		If blnHidden or (strHidden = "") Then
			z = Len(strList)
			OpenLink = "class=hand onclick='OpenFolder("&intID&", true)'"
			If not blnExpand Then
				strDetail = strDetail & "<tr class=TableLine"&(x*lIsTableLinesInAlternateColors) mod 2&" style='cursor:default;padding-top:1;'>"
					If bShowNameColumn = True Then	strDetail = strDetail &		"<td style='padding-left:1;padding-right:3;' "&OpenLink&" "&strHidden&">"&Icon("_folder")&"</td>"
					If bShowNameColumn = True Then	strDetail = strDetail &		"<td style='padding-left:0;' nowrap><span "&OpenLink&" "&strCompressed&">"&SubFolder.Name&"</span></td>"
					If bShowSizeColumn = True Then	strDetail = strDetail &		"<td nowrap>&nbsp;</td>"
					If bShowLastModifiedColumn = True Then	strDetail = strDetail &		"<td nowrap>"&SubFolder.DateLastModified&"</td>"
					If bShowAttributesColumn = True Then	strDetail = strDetail &		"<td nowrap>"&strAttributes&"&nbsp;</td>"
					If bShowLastSpacingColumn = True Then	strDetail = strDetail &		"<td>&nbsp;</td>"
				strDetail = strDetail &			"</tr>"
			End If
			SubdirCount = SubFolder.SubFolders.Count
			If err Then SubdirCount = 0 : err = 0
			If SubdirCount > 0 Then
				strList = strList & "<tr><td>"&_
					"<img id=img0_"&intID&" src=images/_plus.gif width=16 height=16 class=hand onclick='OpenFolder("&intID&", false)'>"&_
					"<img id=img1_"&intID&" src=images/_minus.gif style=display:none; width=16 height=16 class=hand onclick='CloseFolder("&intID&")'></td>"
			Else
				strList = strList & "<tr><td>"&_
					"<img id=img0_"&intID&" src=images/_blank.gif width=16 height=16>"&_
					"<img id=img1_"&intID&" src=images/_blank.gif width=16 height=16 style='display:none;'></td>"
			End If
			VirtualPath = Server.URLEncode(Mid(SubFolder.Path,Len(strInitDirectory) + 1) & "\")
			strList = strList & "<td style='padding-right:5;' "&OpenLink&">"&_
				"<img id=img2_"&intID&" src=images/_folder.gif width=16 height=16 border=0>"&_
				"<img id=img3_"&intID&" src=images/_folderopen.gif style=display:none; width=16 height=16 border=0></td>"&_
				"<td width=100% nowrap><span "&OpenLink&" "&strCompressed&">"& SubFolder.Name & "</span></td></tr>"&_
				"<tr id=tr"&intID&" style='display:none;'>"&_
				"<td style='width:16;background-image:url(images/_vert.gif);'>&nbsp;</td>"&_
				"<td colspan=2 id=td"&intID&" link="""& VirtualPath & """ pid="&intTD&"></td></tr>"
			intID = intID + 1
			x=x+1
		End If
	Next
	'Fix the images For the last item in the list (End images)
	If (z > 0) Then strList = Left(strList,z) & Replace(Replace(Replace(strList,"s.gif","send.gif", z+1),"k.gif","kend.gif"),"t.gif","tend.gif")

	If not blnExpand Then
		For Each File In Files
			GetAttributes File.Attributes
			If blnHidden or (strHidden = "") Then
				FileSize = File.Size
				DirSize = DirSize + FileSize
				FileSize = FileSize / 1024
				If (FileSize > 0) And (FileSize < 1) Then FileSize = 1

				If blnIconFolder Then
					ImageFile = Mid(File.Path,Len(IconFolder), instrrev(File.Path,".") - Len(IconFolder))
				Else
					ImageFile = File.Path
				End If
				
				'For .url file provide the direct link to the targets
				Dim sShortCutLink
				sShortCutLink = ""
				If blnDownload Then
					If LCase(Right(File.Path,4)) = ".url" Then
						OpenLink = ""
						sShortCutLink = FindLink(File.Path)	
					Else
						OpenLink = "class=hand onclick='GetFile("""&Server.URLEncode(File.Path)&""")'" 
					End If
				Else 
					OpenLink = ""
				End If
				strDetail = strDetail & "<tr class=TableLine"& (x*lIsTableLinesInAlternateColors) mod 2&" style='cursor:default;padding-top:1;'>"
				
				If bShowNameColumn = True Then
					strDetail = strDetail & 	"<td style='padding-left:1;padding-right:3;' "&OpenLink&" "&strHidden&">"&Icon(ImageFile)&"</td>"
					strDetail = strDetail & 	"<td style='padding-left:0;' "&strCompressed&" nowrap>"
					
					If sShortCutLink = "" Then
						strDetail = strDetail & "<a title=" & """" & FileName(File.Name) & """" & "><span  "&OpenLink&">"&FileNameLimited(File.Name)&"</span></a>"
					Else
						strDetail = strDetail & "<a " & " title=" & """" & FileName(File.Name) & """" & " class=" & """" & "ActiveLink" & """" &  " target=" & """" & "new" & """" & " href=" & """" & sShortCutLink & """" & "><span>"&FileNameLimited(File.Name)&"</span>" & "</a>"
					End If
					
					strDetail = strDetail & "</td>"
				End If
				
				If bShowSizeColumn = True Then	strDetail = strDetail & "<td align=right nowrap>"&FormatNumber(Int(FileSize),0,-1,-1)&" KB</td>"
				If bShowLastModifiedColumn = True Then strDetail = strDetail & "<td nowrap>"&File.DateLastModified&"</td>"
				If bShowAttributesColumn = True Then strDetail = strDetail & "<td nowrap>"&strAttributes&"&nbsp;</td>"
				If bShowLastSpacingColumn = True Then strDetail = strDetail & "<td>&nbsp;</td>"
				strDetail = strDetail & "</tr>"
				x=x+1
			End If
		Next
	End If

	strList = strList & "</table>"
	strDetail = strDetail & "</table></td></tr></table>"

	If err <> 0 Then
		response.write "alert('Error: " & FixJS(err.description) & "');" & vbcrlf
		response.write "parent.divDetail.innerHTML = '<span style=visibility:hidden;>blank</span>';" & vbcrlf
	Else
		If OriginalID <> intID Then
			If blnExpand Then response.write "parent.tr"&intTD&".style.display='inline';" & vbcrlf
			response.write "parent.td"&intTD&".innerHTML = '" & FixJS(strList) & "';" & vbcrlf
		End If
		If not blnExpand Then
			response.write "parent.divDetail.innerHTML = '" & FixJS(strDetail) & "';" & vbcrlf
			response.write "parent.window.status = '" & x & " object(s) (Disk free space:"&FormatSize(FreeSpace)&")      (Directory size:"&FormatSize(DirSize)&")';" & vbcrlf
		End If
	End If

	Set Files = Nothing
	Set SubFolders = Nothing
	Set Folder = Nothing
End Sub

'Initial drive generation routine
Private Sub InitDrives()
	On Error Resume Next
	Dim PathName, OpenLink
	Dim strRootNode
	
	If sRootNodeCaption <> "" Then
		strRootNode = sRootNodeCaption
	Else
		strRootNode = Request.ServerVariables("HTTP_HOST")
	End If
	
	OpenLink = " class=hand onclick='Root();'"
	strList = "<table cellpadding=0 cellspacing=0 border=0 style='padding:0 0;'>"&_
		"<tr><td style='padding-left:2;' "&OpenLink&"><img src='images/_computer.gif' width=16 height=16 border=0></td>"&_
		"<td colspan=2><span "&OpenLink&">"& strRootNode &"</span></td></tr>"

	If Len(strInitDirectory) > 1 Then
		If Len(strInitDirectory) > 3 Then
			PathName = Mid(strInitDirectory, instrrev(strInitDirectory,"\",Len(strInitDirectory)-1,1)+1)
			AddObject "_folder", PathName, "", Server.URLEncode(PathName&"\"), "End"
		Else
			AddDrive FileSystem.GetDrive(strInitDirectory), "End"
		End If
	Else
		Set Drives = FileSystem.Drives
		For Each Drive In Drives
			AddDrive Drive, ""
		Next
		Set Drives = Nothing
		AddObject "_icons", "Icons", "", Server.URLEncode(IconFolder), "End"
	End If

	strList = strList & "</table>"

	If err <> 0 Then response.write "alert('Error: " & FixJS(err.description) & "');" & vbcrlf
	response.write "parent.divList.innerHTML = '" & FixJS(strList) & "';" & vbcrlf
	response.write "parent.divDetail.innerHTML = '<span style=visibility:hidden;>blank</span>';" & vbcrlf

End Sub

Private Sub AddDrive(Drive, ImageEnd)
	Dim DriveName, DriveImage, DriveVolume
	select case Drive.DriveType
		Case 0: DriveImage = "_unknown": DriveName = "???"
		Case 1:
			select case LCase(Drive.DriveLetter)
				case "a":	DriveImage = "_floppya":  DriveName = "3 <span style='font-size:10pt'>&#189;</span> Floppy"
				case "b":	DriveImage = "_floppyb":  DriveName = "5 <span style='font-size:10pt'>&#188;</span> Floppy"
				case Else:	DriveImage = "_removable":  DriveName = "Removable Disk"
			End select
		Case 2: DriveImage = "_drive":   DriveName = "Local Disk"
		Case 3: DriveImage = "_network": DriveName = "Network Drive"
		Case 4: DriveImage = "_cdrom":   DriveName = "Compact Disk"
		Case 5: DriveImage = "_ram":     DriveName = "RAM Drive"
	End select
	DriveVolume = ""
	If Drive.DriveType = 3 Then
		If Drive.IsReady Then DriveVolume = Drive.ShareName
	Else
		If Drive.DriveType <> 1 Then
			If Drive.IsReady Then DriveVolume = Drive.VolumeName
		End If
	End If
	If DriveVolume <> "" Then DriveName = DriveVolume
	AddObject DriveImage, DriveName, Drive.DriveLetter, Drive.DriveLetter&":%5C",ImageEnd
End Sub

Private Sub AddObject(DriveImage, DriveName, DriveLetter, OpenLocation, ImageEnd)
	Dim OpenLink
	OpenLink = "class=hand onclick='OpenFolder("&intID&", true)'"
	strList = strList & "<tr><td style='padding-left:5;'>"
	If DriveLetter = "" Then
		'strList = strList & "<img id=img0_"&intID&" src=images/_blank"&ImageEnd&".gif width=16 height=16>"&_
		'"<img id=img1_"&intID&" src=images/_blank"&ImageEnd&".gif style=display:none;></td>"
	Else
		Drivename = DriveName & " ("&DriveLetter&":)"
	End If
	strList = strList & "<img id=img0_"&intID&" src=images/_plus"&ImageEnd&".gif width=16 height=16  class=hand onclick='OpenFolder("&intID&", false)'>"&_
	"<img id=img1_"&intID&" src=images/_minus"&ImageEnd&".gif style=display:none; width=16 height=16 class=hand onclick='CloseFolder("&intID&")'></td>"
	strList = strList & "<td style='padding-right:5;' "&OpenLink&">"&_
		"<img id=img2_"&intID&" src='images/"&DriveImage&".gif' width=16 height=16 border=0>"&_
		"<img id=img3_"&intID&" src='images/"&DriveImage&".gif' style=display:none; width=16 height=16 border=0></td>"&_
		"<td width=100% nowrap><span "&OpenLink&">"&DriveName&"</span></td>"&_
		"</tr><tr id=tr"&intID&" style='display:none;'>"&_
		"<td style='width:22;background-position:5 0;background-image:url(images/_vert"&ImageEnd&".gif);'>&nbsp;</td>"&_
		"<td colspan=2 id=td"&intID&" link="""&OpenLocation&"""></td></tr>"
	intID = intID + 1
End Sub

'Return a javascript frendly string
Private Function FixJS(strInput)
	FixJS = replace(replace(strInput,"\","\\"),"'","\'")
End Function

'Correctly format data size to 3 characters (### XX)
Private Function FormatSize(Size)
	Size = Size * 1000 / 1024
	If Size < 10 ^ 3 Then
		FormatSize = Size & " Bytes"
	ElseIf Size < 10 ^ 4 Then
		FormatSize = FormatNumber(Size/10^3, 2) & " KB"
	ElseIf Size < 10 ^ 5 Then
		FormatSize = FormatNumber(Size/10^3, 1) & " KB"
	ElseIf Size < 10 ^ 6 Then
		FormatSize = FormatNumber(Size/10^3, 0) & " KB"
	ElseIf Size < 10 ^ 7 Then
		FormatSize = FormatNumber(Size/10^6, 2) & " MB"
	ElseIf Size < 10 ^ 8 Then
		FormatSize = FormatNumber(Size/10^6, 1) & " MB"
	ElseIf Size < 10 ^ 9 Then
		FormatSize = FormatNumber(Size/10^6, 0) & " MB"
	ElseIf Size < 10 ^ 10 Then
		FormatSize = FormatNumber(Size/10^9, 2) & " GB"
	ElseIf Size < 10 ^ 11 Then
		FormatSize = FormatNumber(Size/10^9, 1) & " GB"
	Else
		FormatSize = FormatNumber(Size/10^9, 0) & " GB"
	End If
End Function

'Search .lnk or .url files for shortcut
Private Function FindLink(strFileName)
	dim WshShell, oLink
    Set WshShell = Server.CreateObject("WScript.Shell")
    Set oLink = WshShell.CreateShortcut(strFileName)
    FindLink = oLink.TargetPath
    Set oLink = Nothing
    Set WshShell = Nothing
End Function

'Return the path portion of the file path
Private Function ParseFolder(PathString)
	Dim x
	x = instrrev(PathString, "\")
	If x > 0 Then ParseFolder = Left(PathString, x)
End Function

'Returns the cooresponding Icon to display
Private Function Icon(file)
	Dim x, ext, LinkImage, alt
	If LCase(Right(file,4)) = ".lnk" Then
		file = FindLink(file)
		alt = file
		If Len(file) = 3 Then
			select case FileSystem.GetDrive(file).DriveType
				Case 0: file = "_unknown"
				Case 1:
					select case LCase(FileSystem.GetDrive(file).DriveLetter)
						case "a":	file = "_floppya"
						case "b":	file = "_floppyb"
						case Else:	file = "_removable"
					End select
				Case 2: file = "_drive"
				Case 3: file = "_network"
				Case 4: file = "_cdrom"
				Case 5: file = "_ram"
	   		End select
	   	Else
			If (instrrev(file,"\") > instrrev(file,".")) Then
				if FileSystem.FolderExists(file) then file = "_folder"
			End If
	   	End If
		LinkImage = "<img src='images/_link.gif' alt=""Location: "&alt&""" align=absmiddle width=16 height=16 border=0 style='position:absolute;left:0;'>"
	ElseIf LCase(Right(file,4)) = ".url" Then
		alt = FindLink(file)
		LinkImage = "<img src='images/_link.gif' alt=""Location: "&alt&""" align=absmiddle width=16 height=16 border=0 style='position:absolute;left:0;'>"
	End If
	x = instrrev(file, ".")
	If x > 0 Then
		ext = Mid(file,x+1)
		If LCase(ext) = "ico" and blnIcon Then
			ext = "download.asp?type=image/x-icon&file="&file
		Else
			If Not FileSystem.FileExists(IconFolder & "ext_" & ext & ".gif") Then ext = ""
			ext = "images/ext_" & ext & ".gif"
		End If
	Else
		If InStr(1, file, "_") > 0 Then ext = "images/" & file & ".gif" Else ext = "images/ext_.gif"
	End If
	Icon = "<img src="""&ext&""" width=16 height=16 border=0>" & LinkImage
End Function

'Returns the Filename to display - limi upto max chars
Private Function FileNameLimited(ByVal value)
	Dim sExt, aExt
	If blnIconFolder Then
		If Left(value,1) <> "_" Then value = UCase(Mid(value,5))
		value = Replace(value,".GIF"," ",1,-1,1)
		If value = " " Then value = "Default"
	Else
		If InStr(1,value,".") > 0 then
			aExt = Array(".cnf.",".desklink.",".lnk.",".mapimail.",".mydocs.",".pif.",".shb.",".shs.",".url.",".xnk.")
			sExt = LCase(Mid(value, instrrev(value,"."))) & "."
			if UBound(Filter(aExt, sExt)) <> -1 then value = Left(value, instrrev(value,".")-1)
		End If
	End If
	If Len(value) > lMaxFileNameLengthForDisplay Then
		value = Left (value, lMaxFileNameLengthForDisplay) & "..."
	End If
	FileNameLimited = value
End Function

'Returns the Filename to display
Private Function FileName(ByVal value)
	Dim sExt, aExt
	If blnIconFolder Then
		If Left(value,1) <> "_" Then value = UCase(Mid(value,5))
		value = Replace(value,".GIF"," ",1,-1,1)
		If value = " " Then value = "Default"
	Else
		If InStr(1,value,".") > 0 then
			aExt = Array(".cnf.",".desklink.",".lnk.",".mapimail.",".mydocs.",".pif.",".shb.",".shs.",".url.",".xnk.")
			sExt = LCase(Mid(value, instrrev(value,"."))) & "."
			if UBound(Filter(aExt, sExt)) <> -1 then value = Left(value, instrrev(value,".")-1)
		End If
	End If
	FileName = value
End Function


'Return the corresponding attributes
Private Sub GetAttributes(value)
	strAttributes = ""
	strCompressed = ""
	strHidden = ""
	If (value And 1) > 0	Then strAttributes = "R"
	If (value And 2) > 0	Then strAttributes = strAttributes & "H": strHidden = "style='filter:alpha(opacity=50);'"
	If (value And 4) > 0	Then strAttributes = strAttributes & "S"
	If (value And 32) > 0	Then strAttributes = strAttributes & "A"
	If (value And 2048) > 0	Then strAttributes = strAttributes & "C": strCompressed = "style='color:#0000ff'"
End Sub

%>

<html>
<body>
<SCRIPT>
<!--
<%
	'Get Query Info
	blnExpand = (Request.QueryString("x") = "1")
	FullPath =  ParseFolder(Request.QueryString("path"))
	AppPath = ParseFolder(Replace(Request.ServerVariables("PATH_TRANSLATED"),"/","\"))
	IconFolder = AppPath  & "images\"
	intTD = request.querystring("td")
	intID = request.querystring("id")
	If intID = "" Then intID = 0

	Set FileSystem = Server.CreateObject("Scripting.FileSystemObject")
	
	If Left(strInitDirectory,1) = "." Then
		strInitDirectory =FileSystem.GetFolder(AppPath & strInitDirectory).Path
	End If

	If FullPath <> "" Then
		GetDetail
	Else
		InitDrives
	End If

	Set FileSystem = Nothing

	response.write "parent.id = " & intID & ";" & vbcrlf
%>
//-->
</SCRIPT>
</body>
</html>
