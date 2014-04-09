<%Option Explicit%>
<!--#INCLUDE FILE="clsWhois.asp"-->
<%
Dim StrDomain
Dim StrWhois
Dim ObjWhois

StrDomain = Request.QueryString("Domain")

If Not StrDomain = "" Then
	Set ObjWhois = New clsWhois
	StrWhois = ObjWhois.Lookup(StrDomain)
	Set ObjWhois = Nothing
End If
%>
<HTML>
	<HEAD>
		<TITLE>Whois Lookup With XMLHTTP</TITLE>
	</HEAD>
	<BODY>
		<H1>Whois Lookup with XMLHTTP</H1>
		<P>
			Type in the domain that you would like to perform a WHOIS lookup on.
		</P>
		<FORM>
			<INPUT name="Domain" value="<%=Server.HTMLEncode(StrDomain)%>" size="25"><BR>
			<INPUT type="submit" value="Whois">
		</FORM>
		Created by <A href="http://www.lewismoten.com">Lewis Moten</A>.
		<HR>
		<FONT size="2">
		<XMP><%=StrWhois%></XMP>
		</FONT>
	</BODY>
</HTML>