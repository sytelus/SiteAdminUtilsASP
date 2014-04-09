<%
' Programmed by webmaster@pitic.net
'The program Stores Information in the Session Object of the IIS.
'The Variables used are:
'
'Session("Nombre") = Name of the user.
'Session("Inter") = Path of the admin scripts (in a local web server is something like c:\inetpub\wwwroot\admin)
'Session("Direc") = The root of the User.
'Session("DirDB") = The Path of the DataBase.
'Session("Permiso") = The permission asigned for the user.
'	The "Permiso" varable can store "Si" = yes or "No" = no.

if Session("Nombre") = "" then
Nombre = Request.Form("Usuario")
Session("Nombre") = Nombre
else
Nombre = Session("Nombre")
end if
'~才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才
Inter = Request.ServerVariables("APPL_PHYSICAL_PATH")
'Some Web Servers change the Path of the script
'Inter = Inter & "Admin\"
'~才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才
DirDB = Inter & "Util\Util.mdb"
Inter = Left(Left(Inter, InStrRev(Inter, "\") - 1), InStrRev(Left(Inter, InStrRev(Inter, "\") - 1), "\"))
Inter = Inter & Nombre & "\"
Direc = Right(Left(Inter, Len(Inter) - 1), Len(Left(Inter, Len(Inter) - 1)) - InStrRev(Left(Inter, Len(Inter) - 1), "\"))
Session("Inter") = Inter
Session("Direc") = Direc	
Session("DirDB") = DirDB
Session("Nombre") = Nombre
Permiso = Session("Permiso")
if not permiso = "Si" then
	if nombre = "" then Response.Redirect "index.asp?Txt=1"
	Contra = Request.Form("Contra")
	Set Conn = Createobject("ADODB.Connection")
	Set Rst = Createobject("ADODB.Recordset")
	Conn.Provider = "Microsoft.JET.OLEDB.4.0"
	Conn.Open DirDB
	Rst.CursorType = adOpenKeyset
	Rst.ActiveConnection = conn
	Rst.Open "SELECT * FROM Usuarios ;"
	Pass = 0
	Admin = false
	Do While Not rst.EOF
		if lcase(nombre) = lcase(rst.Fields(0).Value) then
			if contra = rst.Fields(1).Value then
				pass = 1
				If lcase(nombre) = "admin" then 
					Admin = True
					Session("admgr") = hex(date)
				end if
			else
				session.Abandon
				Response.Redirect "index.asp?Txt=1"
			end if
		end if
		rst.MoveNext
	loop
	if pass = 0  then Response.Redirect "index.asp?Txt=1&Nombre=" & nombre
	If Admin = true then Response.Redirect "admin.asp"
	conn.Close
	set rst = nothing
	set conn = nothing
	Session.Timeout = 20
	Session("Permiso") = "Si"
else
	Session("Permiso") = "Si"
end if
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
<div align="center"><center><table border="1" bordercolor="#808080" cellspacing="0" cellpadding="0" width=670 bgcolor="#FFFFFF" ><tr><td>
<P align=center class=chica>~才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才</P>
<FORM name="Files" method="post" action="edit.asp">
<%Set Fso = CreateObject("Scripting.FileSystemObject")
  Set Directorio = Fso.GetFolder(Session("Inter")).Files%>
<div align="center"><table border="1" bordercolor="#000000" cellspacing="0" cellpadding="0" width=620 bgcolor="#FFCC00">
<tr>
<td width="20"  align="middle" bgcolor="#c0c0c0" class=normalb>X</td>
<td width="150" align="middle" bgcolor="#c0c0c0" class=normalb>&nbsp;Name</td>
<td width="100" align="middle" bgcolor="#c0c0c0" class=normalb>&nbsp;View</td>
<td width="100" align="middle" bgcolor="#c0c0c0" class=normalb>&nbsp;Size</td>
<td width="250" align="middle" bgcolor="#c0c0c0" class=normalb>&nbsp;Type</td>
</tr>
<% FN = 0:SN = 0
for each DD in Directorio 
FN = FN + 1:SN = SN + cdbl(dd.size)%>
<tr>
<td width="20"  align="middle" class=normal><input type="checkbox" name="<%= dd.name%>" value="ON"></td>
<td width="150" align="left"   class=normal>&nbsp;<%= dd.name%></td>
<td width="100" align="left"   class=normal>&nbsp;<A href="../<%= Nombre%>/<%= dd.name%>">Preview</a></td>
<%Num = formatnumber((cdbl(dd.size)/1024),0):If Num < 1 then Num = 1%>
<td width="100" align="right"  class=normal>&nbsp;<%= Num%> KB</td>
<td width="250" align="left"   class=chica>&nbsp;<%= dd.type%></td>
</tr>
<% next %>
<tr>
<td width="20" align="middle" bgcolor="#c0c0c0" class=normalb>X</td>
<td width="150" align="left" bgcolor="#c0c0c0" class=normalb>&nbsp;<%= FN%> Files</td>
<td width="100" align="left" bgcolor="#c0c0c0" class=normal>&nbsp;</A></td>
<%Num = formatnumber(SN/1024,0):If num < 1 then Num = 1%>
<td width="100" align="right" bgcolor="#c0c0c0" class=normalb>&nbsp;<%= Num%> KB</td>
<td width="250" align="left" bgcolor="#c0c0c0" class=chica>&nbsp;</td>
</tr>
</table></div><br>
<center>
<input name="Ren" type="submit" value="Rename">
<input name="Del" type="submit" value=" Delete ">
<input name="Mov" type="submit" value="Edit HTML">
<input name="Cop" type="submit" value="  Copy  ">
<input name="Modif" type="hidden" value="True">
</center></FORM>
<center>
<FORM align="center" method="post" name="LogOff" action="index.asp">
<input type=submit name=Subir  value="Upload Files">
<input type=submit name=Cerrar value="Close Session" >
<input name="LogOff" type="hidden" value="1">
<input name="Usuario" type="hidden" value="<%= Nombre%>">
</FORM>
<P align=center class=chica>~才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才才</P>
</td></tr></table></center></div>
<br>
<DIV align=center>
<TABLE align=center bgColor=FFCC00 border=1 borderColor=black cellPadding=0 cellSpacing=0 width=300>
<TR>
<TD width=300 bgColor=#c0c0c0 colspan="2"><P align=center class=normalb>Space in Disc</P></TD>
</TR>

<TR>
<TD width=150><P align=center class=normal>Used:</P></TD>
<TD width=150><P align=center class=normal><%= formatnumber(SN/1024,0)%> KB</P></TD>
</TR>

<TR>
<TD width=150><P align=center class=normal>Available:</P></TD>
<TD width=150><P align=center class=normal><%= formatnumber(1000-(SN/1024),0)%> KB</P></TD>
</TR>

<TR>
<TD width=150 bgColor=#c0c0c0><P align=center class=normal>Total:</P></TD>
<TD width=150 bgColor=#c0c0c0><P align=center class=normal>1,000 KB</P></TD>
</TR>

</TABLE>
</DIV>
</center>

</body>
</html>
<!-- #include file = "Util/fin.int" -->