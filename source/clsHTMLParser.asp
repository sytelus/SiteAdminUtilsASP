<%
' HTML Parser
'
'	Has ability to retrieve a document from another web site and parse meta data
'	such as Title, Description, and Keywords.
'
' Author: Lewis Moten
' Email: lewis@moten.com
' URL: http://www.lewismoten.com
'
' (C) Copyright 2001 Lewis Moten, All rights reserved.
'
' ------------------------------------------------------------------------------
Class clsHTMLParser
' ------------------------------------------------------------------------------
	Private mStrHTML
	Private mObjRegExp
	Private mObjMatches
	Private mObjMatch
	Public Title
	Public Keywords
	Public Description
' ------------------------------------------------------------------------------
	Public Property Let HTML(ByRef pStrHTML)
		mStrHTML = pStrHTML

		Set mObjRegExp = New RegExp
		mObjRegExp.IgnoreCase = True

		Call ParseTitle()
		Call ParseDescription()
		Call ParseKeywords()

		Set mObjMatch = Nothing
		Set mObjMatches = Nothing
		Set mObjRegExp = Nothing

	End Property
' ------------------------------------------------------------------------------
	Public Property Get HTML()
		HTML = mStrHTML
	End Property
' ------------------------------------------------------------------------------
	Private Sub ParseTitle()
		Title = ""
		mObjRegExp.Pattern = "<TITLE>([^<]*)</TITLE>"
		Set mObjMatches = mObjRegExp.Execute(mStrHTML)
		If mObjMatches.Count = 0 Then Exit Sub
		Title = mObjMatches.item(0).Value
		Title = Replace(Title, "<TITLE>", "", 1, -1, vbTextCompare)
		Title = Replace(Title, "</TITLE>", "", 1, -1, vbTextCompare)
	End Sub
' ------------------------------------------------------------------------------
	Private Sub ParseDescription()
		Description = ""
		mObjRegExp.Pattern = "<META[^>]+(name=""description""|content=""([^""]*)"")[^>]+(name=""description""|content=""([^""]*)"")[^>]*>"
		Set mObjMatches = mObjRegExp.Execute(mStrHTML)
		If mObjMatches.Count = 0 Then Exit Sub
		Description = mObjMatches.item(0).Value
		Description = Mid(Description, InStr(1, Description, "content=""", vbTextCompare) + 9)
		Description = Mid(Description, 1, InStr(1, Description, """", vbTextCompare) -1)
	End Sub
' ------------------------------------------------------------------------------
	Private Sub ParseKeywords()
		Keywords = ""
		mObjRegExp.Pattern = "<META[^>]+(name=""keywords""|content=""([^""]*)"")[^>]+(name=""keywords""|content=""([^""]*)"")[^>]*>"
		Set mObjMatches = mObjRegExp.Execute(mStrHTML)
		If mObjMatches.Count = 0 Then Exit Sub
		Keywords = mObjMatches.item(0).Value
		Keywords = Mid(Keywords, InStr(1, Keywords, "content=""", vbTextCompare) + 9)
		Keywords = Mid(Keywords, 1, InStr(1, Keywords, """", vbTextCompare) -1)
	End Sub
' ------------------------------------------------------------------------------
	Public Function GetURL(ByRef pStrURL)

		Dim lObjSpider
		Dim strText

		If pStrURL = "" Then Exit Function

		On Error Resume Next

		' Different variations of XML objects
		'Set lObjSpider = Server.CreateObject ("MSXML2.XMLHTTP.3.0")
		'Set lObjSpider = Server.CreateObject ("MSXML2.ServerXMLHTTP")
		Set lObjSpider = Server.CreateObject ("Microsoft.XMLHTTP")
		

		' Could not create Internet Control
		If Err Then
			GetURL = "Error: " & Err.Description
			Exit Function
		End If
		
		On Error Goto 0

		With lObjSpider
			.Open "GET", pStrURL, False, "", ""
			.Send
			GetURL = .ResponseText
		End With
		Set LobjSpider = Nothing

		HTML = GetURL
		
	End Function
' ------------------------------------------------------------------------------
End Class
' ------------------------------------------------------------------------------
%>
