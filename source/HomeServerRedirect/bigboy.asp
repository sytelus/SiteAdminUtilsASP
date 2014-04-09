<!-- #INCLUDE FILE="bigboy_ip.inc" -->

<%
	Dim sPreviousURL 
	sPreviousURL = Request.ServerVariables("HTTP_REFERER")
	Dim sThisURL
	sThisURL = "http://" & (Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING"))
	Dim sURLOnBigBoy
	sURLOnBigBoy = sBIGBOY_URL & "/" & Request.QueryString("file")
	
	If sBIGBOY_URL = "" Then
		%>
	<font face="Verdana" size="+1" color="#820000">Our BigBoy Server Hosting Big Files Is Down</font>
	<HR>
	<br>
	The file you were looking for is placed on BigBoy server (part of <a href="www.ShitalShah.com" title="Personal site of Shital Shah">ShitalShah.com</a>) because it's too large or it requires special 
	handling. We can't assure when this server will come up again but it should happen somewhere around mid of September, 2001. If you are in desperate need of
	this file you may try <a href="contact.asp">contacting us</a>, but be aware: there are no promises!<br>
	<br>
	<strong><font size="3" color="#000080">Should I Continue Browsing? Yes. You won't be able to access really huge files, however.
	<br><br>
	<a href="<% =sPreviousURL %>" title="Back to previous page">Back To Previous Page</a>
	</font></strong>
	<br>
	<br>
	File you were looking for is: <% = Request.QueryString("file") %>
	<br>
	<br>
		URL you were trying to reach: <a href="<% =sThisURL %>" title="This URL - You can bookmark it for later use!"><% =sThisURL %></a>. You can use it to
		bookmark this page and visit it later.
	<br>
	<br>
	<br>
	<input onClick="window.close()" type="Button" value="Close This Window">
	&nbsp;&nbsp;&nbsp;&nbsp;
	<input onClick="window.external.addFavorite('<% =sThisURL %>','File I Was Looking On ShitalShah.com')" type="Button" value="Add This To Favorite To Check It Out Later">
	&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="<% =sPreviousURL %>" title="Back to previous page">Back</a>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="www.ShitalShah.com" title="Start page of personal web site of Shital Shah">Home</a>
	&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="contact.asp">Contact</a>
		<%
	Else
		Call Response.Redirect(sURLOnBigBoy)
	End If
 %>