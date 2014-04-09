<%
' Programmed by webmaster@pitic.net
if Session("Inter") = "" then Response.Redirect "index.asp"
if Request.form("Modif") = "True" then
Set Fso = CreateObject("Scripting.FileSystemObject")
Set Directorio = Fso.GetFolder(Session("Inter")).Files
select case Request.Form.Item(Request.Form.Count-1)
case "Rename"
%>
<!-- #include file = "Util/intro.int" -->
<html>
<head>
<META HTTP-EQUIV="REFRESH" CONTENT="600; URL=index.asp?timeout=Si">
<title>Web Site</title>
</head>
<!-- #include file = "Util/style.int" -->
<body background="Util/notebook.gif">
<p align=center class=grande>Web Site Administration</p>
<P align=center class=normalb>User: <%= Session("Nombre")%></P>
<P align=center class=normal>Root Folder: <%= Session("Direc")%></P><center>
<FORM name="Files" method="post" action=edit.asp>
<P class=normalb>Rename Files.</P>
<table border="1" bordercolor="#000000" cellspacing="0" cellpadding="0" width=420 bgcolor="#FFCC00">	
<%Num = 0
for each DD in Directorio
if Request.Form(dd.name) = "ON" then
%>
<tr>
<td width="20"  align="center" class=normal>&nbsp;</td>
<td width="80"  align="center" class=normal>Rename:</td>
<td width="150" align="Left"   class=normal>&nbsp;<%= dd.Name%></td>
<td width="30"  align="center" class=normal>A:</td>
<td width="250" align="Left"   class=normal><input name=<%= dd.name%> type="text" id=<%= dd.name%>>&nbsp;.<%= fso.GetExtensionName(dd.name)%></td>
</tr>
<%
Num = Num  + 1
Session(Num) = dd.name
end if
next
Session("NumFiles") = Num
%>
</table><br>
<input name="DoRen" type="submit" value="Modify">
<input name="Modif" type="hidden" value="True">
</FORM>
<P class=normal>&nbsp;</P>
<P align=center class=normal>
<A href="upload.exe?Ref=<%= Session("Ref")%>">Upload Files</A><br>
<A href="index.asp?LogOff=1&Nombre=<%= Session("Nombre")%>">Close Session</A><br>
</P>
</body>
</html>
<!-- #include file = "Util/fin.int" -->
<%
Case "Modify"
	Dim Cuantos(),Ident1,Ident3
	Ident1 = "No":Ident3 = "No"
	for a = 1 to cint(Request.Form.Count) - 2
		redim preserve Cuantos(a)
		Cuantos(a) = Request.Form(a).Item
		if Cuantos(a) = "" then Ident1 = "Si"
		if not ident1 = "Si" then
			for b = 1 to len(Cuantos(a))
				select case mid(Cuantos(a),b,1)
					case "\","/",":","?","<",">","|",chr(34)
						Ident3 = "Si"
				end select
			next
		end if
	next
	if Ident1 = "Si" then
		Response.Write "<html>"
		Response.Write "<META HTTP-EQUIV=""REFRESH"" CONTENT=""600; URL=index.asp?timeout=Si"">"
		Response.Write "<head>"
		Response.Write "<title>Web Site</title>"
		Response.Write "<body background=""Util/notebook.gif"">"
		Response.Write "<p align=center><font size=5><b>Web Site Administration</b></font></p>"
		Response.Write "<p align=center>Not Valid File Name. Try with another name.</p>"
		Response.Write "<p align=center></p>"		
		Response.Write "<p align=center><INPUT type=button Value=Back onclick='window.history.back(-1);'></p>"
		Response.Write "</body>"
		Response.Write "</html>"
	elseif Ident3 = "Si" then
		Response.Write "<html>"
		Response.Write "<META HTTP-EQUIV=""REFRESH"" CONTENT=""600; URL=index.asp?timeout=Si"">"
		Response.Write "<head>"
		Response.Write "<title>Web Site</title>"
		Response.Write "<body background=""Util/notebook.gif"">"
		Response.Write "<p align=center><font size=5><b>Web Site Administration</b></font></p>"
		Response.Write "<p align=center>Not valid Character in File Name, You can not use: \/:*?" & chr(34) & "<>| . Try another with a diferent name.</p>"
		Response.Write "<p align=center></p>"		
		Response.Write "<p align=center><INPUT type=button Value=Back onclick='window.history.back(-1);' id=button1 name=button1></p>"
		Response.Write "</body>"
		Response.Write "</html>"
	else	
%>
<!-- #include file = "Util/intro.int" -->
<html>
<head>
<META HTTP-EQUIV="REFRESH" CONTENT="600; URL=index.asp?timeout=Si">
<title>Sitio Web</title>
</head>
<!-- #include file = "Util/style.int" -->
<body background="Util/notebook.gif">
<p align=center class=grande>Web Site Administration</p>
<P align=center class=normalb>User: <%= Session("Nombre")%></P>
<P align=center class=normal>Root Folder: <%= Session("Direc")%></P><center>
<p align="center" class=normalb>Renamed Files</p>
<div align="center">
<center>
<table border="1" bgcolor="#FFCC00" cellspacing="0" cellpadding="0" width="500">
<%for a = 1 to ubound(Cuantos)
Vie = session("inter") & session(a)
Nue = session("inter") & Cuantos(a) & "." & fso.GetExtensionName(session(a))
fso.CopyFile vie,nue,true%>
<tr>
<td>
<p align="center">
<font class=normalb>Rename File:<br>Before: </font><font class=normal><%= Vie%><br>
<font class=normalb>After: </font><font class=normal><%= Nue%></font>
</p>
<%If Not Vie = Nue Then fso.DeleteFile(Vie) 
next%>
</tr>
</table>
</center>
</div>
<p align="center"><%= formatdatetime(date,0)%></p>
<p align=center><A href="manager.asp">Back to Web Site Administration</A></p>
</body>
</html>
<!-- #include file = "Util/fin.int" -->
<%
	end if
Case " Delete "
%>
<!-- #include file = "Util/intro.int" -->
<html>
<head>
<META HTTP-EQUIV="REFRESH" CONTENT="600; URL=index.asp?timeout=Si">
<title>Web Site</title>
</head>
<!-- #include file = "Util/style.int" -->
<body background="Util/notebook.gif">
<p align=center class=grande>Web Site Administration</p>
<P align=center class=normalb>User: <%= Session("Nombre")%></P>
<P align=center class=normal>Root Folder: <%= Session("Direc")%></P><center>
<FORM name="Files" method="post" action"edit.asp">
<P class=normalb>Delete Files.</P>
<table border="1" bordercolor="#000000" cellspacing="0" cellpadding="0" width=170 bgcolor="#FFCC00">	
<%Num = 0
for each DD in Directorio
if Request.Form(dd.name) = "ON" then%>
<tr>
<td width="20"  align="center" class=normal>&nbsp;</td>
<td width="150" align="Left"   class=normal>&nbsp;<%= dd.Name%></td>
</tr>
<%
Num = Num  + 1
Session(Num) = dd.name
end if
next
Session("NumFiles") = Num
%>
</table><br>
<input name="DoRen" type="submit" value="Delete">
<input name="Modif" type="hidden" value="True">
</FORM>
<P class=normal>&nbsp;</P>
<P align=center class=normal>
<A href="upload.exe?Ref=<%= Session("Ref")%>">Upload Files</A><br>
<A href="index.asp?LogOff=1&Nombre=<%= Session("Nombre")%>">Close Session</A><br>
</P>
</body>
</html>
<!-- #include file = "Util/fin.int" -->
<%
Case "Delete"
%>
<!-- #include file = "Util/intro.int" -->
<html>
<head>
<META HTTP-EQUIV="REFRESH" CONTENT="600; URL=index.asp?timeout=Si">
<title>Web Site</title>
</head>
<!-- #include file = "Util/style.int" -->
<body background="Util/notebook.gif">
<p align=center class=grande>Web Site Administration</p>
<P align=center class=normalb>User: <%= Session("Nombre")%></P>
<P align=center class=normal>Root Folder: <%= Session("Direc")%></P><center>
<p align="center" class=normalb>Deleted Files</p>
<div align="center">
<center>
<table border="1" bgcolor="#FFCC00" cellspacing="0" cellpadding="0" width="500">
<%for a = 1 to cint(session("NumFiles"))
fso.DeleteFile(session("inter") & session(a))%>
<tr>
<td>
<p align="center">
<font class=normalb>Deleted File:<br>Name: </font><font class=normal><%= session(a)%>
</p> 
<%next%>
</tr>
</table>
</center>
</div>
<p align="center"><%= formatdatetime(date,0)%></p>
<p align=center><A href="manager.asp">Back to Web Site Administration</A></p>
</body>
</html>
<!-- #include file = "Util/fin.int" -->
<%
Case "Edit HTML"
for each DD in Directorio
if Request.Form(dd.name) = "ON" then
	Archivo = dd.name
	exit for
end if
next
If not Archivo = "" then
select case lcase(fso.GetExtensionName(archivo))
case "htm","html","txt"
Set FileX = Fso.OpenTextFile(Session("Inter") & Archivo)
Texto = FileX.readall
set FileX = nothing
Session("Editar") = Archivo
%>
<!-- #include file = "Util/intro.int" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Web Site</TITLE>
</HEAD>
<!-- #include file = "Util/style.int" -->
<BODY background="Util/notebook.gif">
<P align=center class=grande>Web Site Administration</P>
<P align=center class=normalb>User: <%= Session("Nombre")%></P>
<FORM name="Editar" method="post" action="edit.asp">
<table align=center border="1" bordercolor="#000000" cellspacing="0" cellpadding="0" height=460 width=620 bgcolor="#FFCC00">
<tr>
<td align="center">
<P align=center class=normal>File: <%= Archivo%>
<TEXTAREA name=txt style="HEIGHT: 410px; WIDTH: 570px"><%= Texto%></TEXTAREA>
<input type=submit name=gua value="Save">
<input type=submit name=Cer  value="Save and Continue">
<input type=submit name=Vol  value="Back to Web Site Admin with no saveing">
<input type=hidden name=Modif  value="True">
</P>
</td>
</table>
</form>
</BODY>
</HTML>
<!-- #include file = "Util/fin.int" -->
<%
Case else
%>
<!-- #include file = "Util/intro.int" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Web Site</TITLE>
</HEAD>
<!-- #include file = "Util/style.int" -->
<BODY background="Util/notebook.gif">
<P align=center class=grande>Web Site Administration</P>
<P align=center class=normalb>User: <%= Nombre%></P>
<P align=center class=normal>Not valid File</P>
</BODY>
</HTML>
<!-- #include file = "Util/fin.int" -->
<%
end select
else
%>
<!-- #include file = "Util/intro.int" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Web Site</TITLE>
</HEAD>
<!-- #include file = "Util/style.int" -->
<BODY background="Util/notebook.gif">
<P align=center class=grande>Web Site Administration</P>
<P align=center class=normalb>User: <%= Nombre%></P>
<P align=center class=normal>Not Valid File</P>
</BODY>
</HTML>
<!-- #include file = "Util/fin.int" -->
<%
end if
Case "Save"
	Set FFF = fso.OpenTextFile(Session("Inter") & Session("Editar"),2,false)
	FFF.write Request.Form("txt")
	FFF.Close
	Set FFF = nothing
	Response.Redirect "manager.asp"
Case "Back to Web Site Admin with no saveing"
	Response.Redirect "manager.asp"	
Case "Save and Continue"
	Set FFF = fso.OpenTextFile(Session("Inter") & Session("Editar"),2,false)
	FFF.write Request.Form("txt")
	Texto = Request.Form("txt")
	FFF.Close
	Set FFF = nothing
%>
<!-- #include file = "Util/intro.int" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Web Site</TITLE>
</HEAD>
<!-- #include file = "Util/style.int" -->
<BODY background="Util/notebook.gif">
<P align=center class=grande>Web Site Administration</P>
<P align=center class=normalb>User: <%= Session("Nombre")%></P>
<FORM name="Editar" method="post" action="edit.asp">
<table align=center border="1" bordercolor="#000000" cellspacing="0" cellpadding="0" height=460 width=620 bgcolor="#FFCC00">
<tr>
<td align="center">
<P align=center class=normal>File: <%= Session("Editar")%>
<TEXTAREA name=txt style="HEIGHT: 410px; WIDTH: 570px"><%= Texto%></TEXTAREA>
<input type=submit name=gua value="Save">
<input type=submit name=Cer  value="Save and Continue">
<input type=submit name=Vol  value="Back to Web Site Admin with no saveing">
<input type=hidden name=Modif  value="True">
</P>
</td>
</table>
</form>
</BODY>
</HTML>
<!-- #include file = "Util/fin.int" -->
<%
Case "  Copy  "
%>
<!-- #include file = "Util/intro.int" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Web Site</TITLE>
</HEAD>
<!-- #include file = "Util/style.int" -->
<BODY background="Util/notebook.gif">
<P align=center class=grande>Web Site Administration</P>
<P align=center class=normalb>User: <%= Session("Nombre")%></P>
<FORM name="Copiar" action="manager.asp">
<table align=center border="1" bordercolor="#000000" cellspacing="0" cellpadding="0" width=320 bgcolor="#FFCC00">
<tr>
<td align="center">
<P align=center class=normal>Function under Construction: Copy</P>
<P align=center class=normal></P>
<input type=submit name=Vol  value="Back to Web Site Administration">
<P align=center class=normal></P>
</td>
</table>
</form>
</BODY>
</HTML>
<!-- #include file = "Util/fin.int" -->
<%
end select
else
	Response.Redirect "manager.asp"
end if%>