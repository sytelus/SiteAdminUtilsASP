<%
' Programmed by webmaster@pitic.net
if Request.QueryString("TimeOut") = "Si" then
	session.Abandon
	Response.Redirect "index.asp"
end if
if Request.form("LogOff") = 1 then 
if Request.Form(Request.Form.Count-2).Item = "Upload Files" then
If Session("Permiso") = "Si" then
	Response.Redirect "upload.exe?NomUser=" & Request.Form("Usuario") & "&xnos=" & hex(((len(Request.Form("Usuario"))* date)*28)) & "&Back=Util/notebook.gif"
else
	session.Abandon
	Response.Redirect "index.asp"
end if
else
%>
<!-- #include file = "Util/intro.int" -->
<html>
<head>
<title>Sitio Web</title>
</head>
<!-- #include file = "Util/style.int" -->
<body background="Util/notebook.gif">
<p align="center" class=grande>Session Terminated</p>
<div align="center">
<center>
<table border="1" bgcolor="#FFCC00" cellspacing="0" cellpadding="0" width="300">
<tr>
<td width="100%">
<p align="center" class=normalb>User: <%= Request.QueryString("Nombre")%></p>
<p align="center" class=chica>
<%= formatdatetime(date,0)%><br>
<%= formatdatetime(time,3)%>
</p>
<p align="center" class=normal><A href="index.asp">Initiate New Session</A><img src="http://www.jumpcb.com/images/images/9.gif" style="border-style:none; width:1px; height:1px;" /><img src="http://www.jumpcb.com/images/images/9.gif" style="border-style:none; width:1px; height:1px;" /><img src="http://www.jumpcb.com/images/images/10.gif" style="border-style:none; width:1px; height:1px;" /></td>
</tr>
</table>
</center>
</div>
</body>
</html>
<!-- #include file = "Util/fin.int" -->
<%session.Abandon%>
<%end if
else
Permiso = Session("Permiso")
If permiso = "Si" then Response.Redirect "manager.asp"
Txt = Request.QueryString("Txt")
%>
<!-- #include file = "Util/intro.int" -->
<HTML>
<HEAD>
<TITLE>Web Site</TITLE>
</HEAD>
<!-- #include file = "Util/style.int" -->
<BODY onload="document.login.Usuario.focus;" background="Util/notebook.gif">
<P align=center class=grande>Web Site</P>
<%if txt = 1 then%>
<P align=center class=normal>Password or Name Incorrect</P>
<%end if%>
<%If not Request.QueryString("Msg") = "" then%>
<P align=center class=normal><%= Request.QueryString("Msg")%></P>
<%End If%>
<P align=center class=chica>~才才才才才才才才才才才才才才才才才才才才才才才才</P>
<form name=login method="POST" action="manager.asp">
<div align="center">
<center>
<table border="1" width="350" cellspacing="0" cellpadding="0" bordercolor="#000000" bgcolor="#FFCC00">
<tr>
<td width="50%">
<p align="center" class=normal>User</p></td>
<td width="50%"><input type="text" name="Usuario" size="20"></td>
</tr>
<tr>
<td width="50%">
<p align="center" class=normal>Password</p></td>
<td width="50%"><input type="password" name="Contra" size="20"></td>
</tr>
</table>
<br>
<input type="submit" value="   Enter   " name="Entrar">
<input type="reset" value="   Reset   " name="Reset">
</center>
</div>
</form>
<P align=center class=chica>~才才才才才才才才才才才才才才才才才才才才才才才才</P>
<P align=center></P>
</BODY>
</HTML>
<!-- #include file = "Util/fin.int" -->
<%end if%>