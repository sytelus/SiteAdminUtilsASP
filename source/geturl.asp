<!--#INCLUDE FILE="clsHTMLParser.asp"-->
<%
Dim StrURL
Dim StrHTML
Dim ObjParser

StrURL = Request.Form("URL")
If Not StrURL = "" Then
	Set ObjParser = New clsHTMLParser
	With ObjParser
		StrHTML = .GetURL(StrURL)
		Response.Write StrHTML
	End With
	Set ObjParser = Nothing
End If
%>
<H1>Get URL</H1>
<P>
	This page will get you another web page without making you actually go there. So you stay protected
	under your firewall and company rules and still view any page YOU want (this is also called HTTP proxy). However this is very simple one so you might not get complecated
	pages and you might not be able to click anywhere like normal (but you can "calculate" link and request
	that page manually).
</P>
<FORM method="post" action="geturl.asp">
	<INPUT size="50" name="URL" value="<%=StrURL%>"><BR>
	<INPUT type="Submit" value="Get URL">
</FORM>
Created by <A href="http://www.lewismoten.com">Lewis Moten</A>.<BR>
Modified by <A href="http://www.ShitalShah.com">Shital Shah</A>.<BR>
<BR>
