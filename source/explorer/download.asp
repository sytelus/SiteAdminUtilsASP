<%
Response.Buffer = True
Response.Clear

strFileName = Request.QueryString("file")
strFileType = Request.QueryString("type")
'All of these should work
'if strFileType = "" then strFileType = "bad/type"
'if strFileType = "" then strFileType = "application/octet-stream"
if strFileType = "" then strFileType = "application/download"
strFile = Right(strFileName, Len(strFileName) - InStrRev(strFileName,"\"))

Response.ContentType = strFileType
Response.AddHeader "Content-Disposition", "attachment; filename=""" & strFile &""""

Set Stream = Server.CreateObject("ADODB.Stream")
Stream.Open
Stream.LoadFromFile strFileName
Contents = Stream.ReadText
Response.BinaryWrite Contents
Stream.Close
Set Stream = Nothing

Response.End
%>