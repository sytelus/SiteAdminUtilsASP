<%
Class clsWhois

	Public URL
	Public CropAfter
	Public CropBefore
	
	Private Sub Class_Initialize()
		URL = "http://www.networksolutions.com/cgi-bin/whois/whois?STRING="
		CropAfter = "<PRE>"
		CropBefore = "</PRE>"
	End Sub
	
	Function LookUp(pStrDomain)

		Dim lStrURL
		Dim lObjRegExp
		Dim lStrLookUp
		
		If pStrDomain = "" Then Exit Function

		lStrLookUp = GetURL(URL & Server.URLEncode(pStrDomain))
		
		Dim lLngStart
		Dim lLngEnd
		
		If Not(CropAfter = "" Or CropBefore = "") Then
			lLngStart = InStr(1, lStrLookUp, CropAfter, vbTextCompare)
			If lLngStart = 0 Then
				lStrLookUp = ""
			Else
				lLngStart = lLngStart + Len(CropAfter) + 1
				lLngEnd = InStr(lLngStart, lStrLookUp, CropBefore, vbTextCompare)
				If lLngEnd = 0 Then
					lStrLookUp = ""
				Else
					lLngEnd = lLngEnd - lLngStart
					lStrLookUp = Mid(lStrLookUp, lLngStart, lLngEnd)
				End If
			End If
		End If
				
		' remove HTML tags
		Set lObjRegExp = New RegExp
		lObjRegExp.Global = True
		lObjRegExp.Pattern = "<[^>]*>"
		lStrLookUp = lObjRegExp.Replace(lStrLookUp, "")
		Set lObjRegExp = Nothing
		
		LookUp = lStrLookUp

	End Function

	Public Function GetURL(ByRef pStrURL)

		Dim lObjSpider
		Dim strText

		If pStrURL = "" Then Exit Function

		On Error Resume Next

		' Different variations of XML objects
		'Set lObjSpider = Server.CreateObject ("MSXML2.XMLHTTP.3.0")
		'Set lObjSpider = Server.CreateObject ("MSXML2.ServerXMLHTTP")
		Set lObjSpider = Server.CreateObject ("Microsoft.XMLHTTP")
			
		' Could not create Spider
		If Err Then Exit Function
			
		On Error Goto 0

		With lObjSpider
			.Open "GET", pStrURL, False, "", ""
			.Send
			GetURL = .ResponseText
		End With
		Set lObjSpider = Nothing

	End Function
	
End Class
%>