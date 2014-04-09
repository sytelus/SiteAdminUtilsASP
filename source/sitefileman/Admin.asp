<%
' Programmed by webmaster@pitic.net
if Session("admgr") = hex(date) then
If Request.form("Modif") = "" then%>
<!-- #include file = "Util/intro.int" -->
<HTML>
<HEAD>
<META HTTP-EQUIV="REFRESH" CONTENT="600; URL=index.asp?timeout=Si">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Web Site Administration</TITLE>
</HEAD>
<!-- #include file = "Util/style.int" -->
<body background="Util/notebook.gif">
<p align=center class=grande>Web Site</p>
<P align=center class=normalb>Web Site Administration</P>
<FORM name="Users" method="post" action="admin.asp">
<div align="center"><table border="1" bordercolor="#000000" cellspacing="0" cellpadding="0" width=650 bgcolor="#FFCC00">
<tr>
<td width="20"  align="middle" bgcolor="#c0c0c0" class=normalb>X</td>
<td width="150" align="middle" bgcolor="#c0c0c0" class=normalb>&nbsp;User</td>
<td width="200" align="middle" bgcolor="#c0c0c0" class=normalb>&nbsp;Name</td>
<td width="100" align="middle" bgcolor="#c0c0c0" class=normalb>&nbsp;Date</td>
<td width="180" align="middle" bgcolor="#c0c0c0" class=normalb>&nbsp;E-Mail</td>
</tr>
<%
DirDB = Session("DirDB")
Set Conn = Createobject("ADODB.Connection")
Set Rst = Createobject("ADODB.Recordset")
Conn.Provider = "Microsoft.JET.OLEDB.4.0"
Conn.Open DirDB
Rst.CursorType = adOpenKeyset
Rst.ActiveConnection = conn
Rst.Open "SELECT * FROM Usuarios ;"
q = 0
Do While Not rst.EOF
if not rst.Fields(0).Value = "admin" then%>
<tr>
<td width="20"  align="middle" class=normal><input type="checkbox" name="<%= rst.Fields(0).Value%>" value="ON"></td>
<td width="150" align="left"   class=normal>&nbsp;<A href="<%= "../" & rst.Fields(0).Value%>"><%= rst.Fields(0).Value%></a></td>
<td width="200" align="left"   class=normal>&nbsp;<%= rst.Fields(2).Value%></td>
<td width="100" align="left"   class=normal>&nbsp;<%= rst.Fields(3).Value%></td>
<td width="180" align="left"   class=normal>&nbsp;<A href="mailto:<%= rst.Fields(4).Value%>"><%= rst.Fields(4).Value%></a></td>
</tr>
<%end if
rst.MoveNext
loop
conn.Close
set rst = nothing
set conn = nothing
%>
<tr>
<td width="20"  align="middle" bgcolor="#c0c0c0" class=normalb>X</td>
<td width="150" align="middle" bgcolor="#c0c0c0" class=normalb>&nbsp;Users</td>
<td width="200" align="middle" bgcolor="#c0c0c0" class=normal>&nbsp;</td>
<td width="100" align="middle" bgcolor="#c0c0c0" class=normalb>&nbsp;</td>
<td width="180" align="middle" bgcolor="#c0c0c0" class=chica>&nbsp;</td>
</tr>
</table></div><br>
<center>
<input name="Nue" type="submit" value="New User">
<input name="Del" type="submit" value="Delete User">
<input name="Edi" type="submit" value="Edit User">
<input name="Modif" type="hidden" value="True">
</FORM>
<FORM align="center" method="post" name="LogOff" action="index.asp">
<input type=submit name=Cerrar value="Close Session">
<input name="LogOff" type="hidden" value="1">
<input name="Usuario" type="hidden" value="<%= Nombre%>">
</FORM>
</center>
</BODY>
</HTML>
<!-- #include file = "Util/fin.int" -->
<%
else
select case Request.Form.Item(Request.Form.Count-1)
case "New User"
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="REFRESH" CONTENT="600; URL=index.asp?timeout=Si">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Web Site Administration</TITLE>
</HEAD>
<!-- #include file = "Util/style.int" -->
<body background="Util/notebook.gif">
<p align=center class=grande>Web Site</p>
<P align=center class=normalb>Web Site Administration</P>
<FORM name="Users" method="post" action="admin.asp">
<div align="center"><table border="1" bordercolor="#000000" cellspacing="0" cellpadding="0">
<tr><td>
<div align="center"><table border="0" bordercolor="#000000" cellspacing="0" cellpadding="0" width=450 bgcolor="#FFCC00">
<tr><td>
<center><LABEL class=normalb>User Data</label></center>
</tr></td>
<tr><td><table border=0 width=450 cellspacing=0 cellpadding=0><tr>
<td width=150 align=right><label class=normal>Name:&nbsp;</label></td>
<td width=300><input type=text name=Name size=27><label class=chica>Example. John Smith</label></td>
</tr></table></td></tr>
<tr><td><table border=0 width=450 cellspacing=0 cellpadding=0><tr>
<td width=150 align=right><label class=normal>Login:&nbsp;</label></td>
<td width=300><input type=text name=Login size=8><label class=chica>Max 8 Letters. Example. jsmith</label></td>
</tr></table></td></tr>
<tr><td><table border=0 width=450 cellspacing=0 cellpadding=0><tr>
<td width=150 align=right><label class=normal>Password:&nbsp;</label></td>
<td width=300><input type=password name=Password size=6><label class=chica>Max 6 Characters.</label></td>
</tr></table></td></tr>
<tr><td><table border=0 width=450 cellspacing=0 cellpadding=0><tr>
<td width=150 align=right><label class=normal>Re-Password:&nbsp;</label></td>
<td width=300><input type=password name=RePassword size=6><label class=chica>Max 6 Characters.</label></td>
</tr></table></td></tr>
<tr><td><table border=0 width=450 cellspacing=0 cellpadding=0><tr>
<td width=150 align=right><label class=normal>E-Mail:&nbsp;</label></td>
<td width=300><input type=text name=EMail size=25></td>
</tr></table></td></tr>
<tr><td  align=center>
<br>
<INPUT name=submit1 type=submit value=Register>
<INPUT name=Modif type=hidden value=True>
<br><br>
</tr></td>
</table></td></tr></div></table>
</form>
</body>
</html>
<%
Case "Register"
Dim RegDat()
redim regdat(5)
regdat(1) = Request.Form("Name")
regdat(2) = Request.Form("Login")
regdat(3) = Request.Form("Password")
regdat(4) = Request.Form("RePassword")
regdat(5) = Request.Form("EMail")
If regdat(1) = "" then Mensaje = Mensaje & "Not valid name<br>"
If regdat(2) = "" then Mensaje = Mensaje & "Not valid Login<br>"
If regdat(3) = "" then Mensaje = Mensaje & "Not valid Password<br>"
If regdat(5) = "" then Mensaje = Mensaje & "Not valid E-Mail<br>"
if len(regdat(2)) >= 9 then Mensaje = Mensaje & "Not valid Login, Max 8 Characters<br>"
if not regdat(3) = "" then if not regdat(3) = regdat(4) then Mensaje = Mensaje & "The password dont match<br>"
if len(regdat(3)) >= 7 then Mensaje = Mensaje & "Not valid Password, Max 6 Characters<br>"
Hab = true
For a = 1 to len(regdat(2))
	B = mid(regdat(2),a,1)
	select case b
	case "\","/",":","?","<",">","|",chr(34)
	Hab = false
	end select
next
If Hab = False then Mensaje = Mensaje & "Not valid Characters in 'LogIn'<br>"
Hab = false
For a = 1 to len(regdat(5))
	B = mid(regdat(5),a,1)
	select case b
	case "@","."
	Hab = true
	end select
next
If not regdat(5) = "" then If Hab = False then Mensaje = Mensaje & "Not valid E-Mail<br>"
Set Fso = Createobject("Scripting.FileSystemObject")
UPath = Left(Session("Inter"), InStrRev(Session("Inter"), "\", Len(Session("Inter")) - 1)) & RegDat(2)
DirChk = fso.FolderExists(UPath)
If DirChk = true then Mensaje = Mensaje & "Login in use, Pick another Please<br>"
if not mensaje = "" then
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="REFRESH" CONTENT="600; URL=index.asp?timeout=Si">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Web Site Administration</TITLE>
</HEAD>
<!-- #include file = "Util/style.int" -->
<body background="Util/notebook.gif">
<p align=center class=grande>Web Site</p>
<P align=center class=normalb>Web Site Administration</P>
<p align=center class=normal><%= Mensaje%></p>
<P align=center><input type=button name=Back onclick=window.history.back(-1); value="Back to Form">
</body>
</html>
<%
else
DirDB = Session("DirDB")
Set Conn = Createobject("ADODB.Connection")
Conn.Provider = "Microsoft.JET.OLEDB.4.0"
Conn.Open DirDB
Conn.execute "INSERT INTO Usuarios (Nombre, Contra, NombreCom, Fecha, Mail) VALUES('" & RegDat(2) & "', '" & RegDat(3) & "', '" & RegDat(1) & "', '" & formatdatetime(date,2) & "', '" & RegDat(5) & "')"
Conn.close
fso.CreateFolder(UPath)
fso.CopyFile Left(Session("Inter"), InStrRev(Session("Inter"), "\", Len(Session("Inter")) - 1)) & "Admin\Util\cono.gif", UPath & "\" 
fso.CopyFile Left(Session("Inter"), InStrRev(Session("Inter"), "\", Len(Session("Inter")) - 1)) & "Admin\Util\flecha_der.gif", UPath & "\"
fso.CopyFile Left(Session("Inter"), InStrRev(Session("Inter"), "\", Len(Session("Inter")) - 1)) & "Admin\Util\flecha_izq.gif", UPath & "\"
fso.CopyFile Left(Session("Inter"), InStrRev(Session("Inter"), "\", Len(Session("Inter")) - 1)) & "Admin\Util\nubes.jpg", UPath & "\"
Set Nuev = fso.CreateTextFile(UPath & "\Default.htm",true)
Nuev.writeline "<HTML><HEAD><TITLE>Page under Construction</TITLE></HEAD>"
Nuev.writeline "<BODY background=""nubes.jpg"">"
Nuev.writeline "<P align=center><b><font size=""7"">Page under Construction</FONT></b></P>"
Nuev.writeline "<center><img border=""0"" src=""flecha_izq.gif"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img border=""0"" src=""cono.gif"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img border=""0"" src=""flecha_der.gif""></center>"
Nuev.writeline "<P align=center>" & RegDat(1) & "</P>"
Nuev.writeline "<P align=center>" & formatdatetime(date,2) & "</P>"
Nuev.writeline "<P align=center><a href=""mailto:" & RegDat(5) & """>E-Mail</a></P>"
Nuev.writeline "</BODY></HTML>"
Nuev.close
Set Nuev = nothing
Response.Redirect "admin.asp"
end if
case "Delete User"
'===================================================================================
'===================================================================================
'===================================================================================
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
<FORM name="Files" method="post" action"admin.asp">
<center>
<P class=normalb>Delete User</P>
<table border="1" bordercolor="#000000" cellspacing="0" cellpadding="0" bgcolor="#FFCC00"
<%DirDB = Session("DirDB")
Set Conn = Createobject("ADODB.Connection")
Set Rst = Createobject("ADODB.Recordset")
Conn.Provider = "Microsoft.JET.OLEDB.4.0"
Conn.Open DirDB
Rst.CursorType = adOpenKeyset
Rst.ActiveConnection = conn
Rst.Open "SELECT * FROM Usuarios ;"
Num = 0
q = true
Do While Not rst.EOF
if not rst.Fields(0).Value = "admin" then
If Request.Form(rst.Fields(0).Value) = "ON" then
If q = true then 
Num = Num + 1
NUser = rst.Fields(0).Value%>
<TR>
<td width="150" align="center" class=normalb><%= rst.Fields(0).Value%></td>
</tr>
<%q = false
end if
end if
end if
rst.MoveNext
loop
conn.Close
set rst = nothing
set conn = nothing%>
</table><br>
<input name="NUser" type="hidden" value="<%= NUser%>">
<%If Not Num = 0 then%>
<input name="DoRen" type="submit" value="Delete User">
<%End If%>
<input name="Modif" type="hidden" value="True">
</center>
</FORM>
<P class=normal>&nbsp;</P>
<P class=normal align=center>The Folder: <b>'<%= NUser%>'</b> change the name to: <b>'Baja_<%= NUser%>'</b></P><br>
</body>
</html>
<!-- #include file = "Util/fin.int" -->
<%
case "Delete User"
DirDB = Session("DirDB")
Set Conn = Createobject("ADODB.Connection")
Conn.Provider = "Microsoft.JET.OLEDB.4.0"
Conn.Open DirDB
Conn.execute "DELETE * FROM Usuarios WHERE Nombre = '" & Request.Form("NUser") & "' ;"
Conn.close
set Conn = nothing
Set Fso = Createobject("Scripting.FileSystemObject")
UserDir = Left(Session("Inter"), InStrRev(Session("Inter"), "\", Len(Session("Inter")) - 1)) & Request.Form("NUser")
fso.GetFolder(UserDir).Name = "Baja_" & Request.Form("NUser")
Set Fso = nothing
Response.Redirect "admin.asp"

case "Edit User"
	Response.Write "Page under Construction"
end select
end if
Else
	Response.Redirect "index.asp?timeout=Si&Msg=NonAuthorized Page"
End If
%>