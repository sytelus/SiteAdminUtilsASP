Attribute VB_Name = "Main0"
'Programmed By webmaster@pitic.net
Public Sub Cgi_Main()
    If Ind = "True" Then
        If Vali = True Then
            SendHeader2 "Sent File", LCase(App.EXEName) & ".exe", "600"
            Send "<body bgcolor=""#FFFFFF"" background=""" & BackG & """>"
            Estilo
            Send "<p align=center class=grande>Sent File</p>"
            Send "<p align=center class=normalb>User: " & NmUser & "</p>"
            Send "<div align=""center""><center>"
            Send "<table border=""1"" bgcolor=""#FFCC00"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
            Send "<tr><td width=""100%"">"
            Sep
            Send "<FORM name=""env"" action""upload.exe"">"
            Send "<P align=center><font class=normal>Name:</font><br><font class=chica>" & Nombre_Archivo & "</font</P>"
            Send "<P align=center><font class=normal>Size:</font><br><font class=chica>" & Format((lContentLength / 1024), "#,##0.00") & " KB</font</P>"
            Send "<P align=center><font class=normal>Type:</font><br><font class=chica>" & ContenidoType & "</font</P>"
            Send "<P align=center><font class=normal>Date:</font><br><font class=chica>" & FormatDateTime(Date, vbShortDate) & " " & Time & "</font</P>"
            Send "<br>"
            Send "<input name=""NomUser"" type=""hidden"" value=""" & NmUser & """>"
            Send "<input name=""xnos"" type=""hidden"" value=""" & Per & """>"
            Send "<input name=""Back"" type=""hidden"" value=""" & BackG & """>"
            Send "<center><input name=""uplo"" type=""submit"" value=""Back to Send Files Form""></center>"
            Sep
            Send "</td></tr></table></center></form></div>"
        Else
            SendHeader2 "Not Authorized File", LCase(App.EXEName) & ".exe", "600"
            Send "<body bgcolor=""#FFFFFF"" background=""" & BackG & """>"
            Estilo
            Send "<p align=center class=grande>Not Authorized File</p>"
            Send "<p align=center class=normalb>User: " & NmUser & "</p>"
            Send "<div align=""center""><center>"
            Send "<FORM name=""env"" action""upload.exe"">"
            Send "<table border=""1"" bgcolor=""#FFCC00"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
            Send "<tr><td width=""100%"">"
            Sep
            Send "<P align=center class=normalb>Not Authorized File in Type or Size</P>"
            Send "<P align=center class=normal>It is not allowed to upload EXE or System files</P>"
            Send "<P align=center class=normal>It is not allowed to upload files biger then 1.2MB</P>"
            Send "<input name=""NomUser"" type=""hidden"" value=""" & NmUser & """>"
            Send "<input name=""xnos"" type=""hidden"" value=""" & Per & """>"
            Send "<input name=""Back"" type=""hidden"" value=""" & BackG & """>"
            Send "<center><input name=""uplo"" type=""submit"" value=""Back to Send Files Form""></center>"
            Sep
            Send "</td></tr></table></center></form></div>"
        End If
    Else
        SendHeader "Send Files Form"
        Estilo
        Send "<body bgcolor=""#FFFFFF"" background=""" & BackG & """>"
        Send "<FORM name=""Enviar"" enctype=""multipart/form-data"" action=""" & LCase(App.EXEName) & ".exe?UpFile=True&NomUser=" & NmUser & "&xnos=" & Per & "&Back=" & BackG & """ method=""POST"">"
        Send "<p align=center><STRONG> </STRONG></p>"
        Send "<p align=center class=grande>Send Files Form</p>"
        Send "<p align=center class=normalb>User: " & NmUser & "</p>"
        Send "<div align=""center""><center>"
        Send "<table border=""1"" bgcolor=""#FFCC00"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
        Send "<tr><td width=""100%"">"
        Sep
        Send "<P align=center class=normal>Select the file to send.</P>"
        Send "<Center><INPUT name=fName type=file></CENTER>"
        Send "<CENTER><FONT class=chica>The time of process change based on the size of the file to send.</FONT><br><br>"
        Send "<INPUT name=button2 type=submit value=~Send~><br><br>"
        Sep
        Send "</CENTER></td></tr></table></center></div>"
        Send "</FORM>"
    End If
    Send "<p align=center>&nbsp;</p>"
    Send "<p align=center><A href=""manager.asp"">Back to Web Site Administration</A></p>"
    Send "<p align=center>&nbsp;</p>"
    SendFooter
    End
End Sub
Public Sub NoAut()
    SendHeader2 "Not Authorized", "timeout.asp", "5"
    Send "<body bgcolor=""#FFFFFF"" background=""../Images/reja.gif"">"
    Estilo
    Send "<p align=center class=grande>Not Authorized</p>"
    Send "<div align=""center""><center>"
    Send "<table border=""1"" bgcolor=""#FFFFCC"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
    Send "<tr><td width=""100%"">"
    Sep
    Send "<P align=center class=normalb>Not Authorized</P>"
    Sep
    Send "</td></tr></table></center></div>"
    Send "<p align=center>&nbsp;</p>"
    SendFooter
End Sub
Private Sub Estilo()
    Send "<Style Type=""text/css"">"
    Send "<!--"
    Send "    .grande{font-family:verdana,helvetica,arial,sans serif;color:#000000;font-size:20pt;}"
    Send "    .normalB{font-family:verdana,helvetica,arial,sans serif;color:#000000;font-size:10pt;font-weight:bold;}"
    Send "    .normal{font-family:verdana,helvetica,arial,sans serif;color:#000000;font-size:10pt;}"
    Send "    .chica{font-family:verdana,helvetica,arial,sans serif;color:#000000;font-size:8pt;}"
    Send "    a:link{color:#0000FF;}"
    Send "    a:visited{color:#0000FF;}"
    Send "    a:hover{color:#ff3300;}"
    Send "-->"
    Send "</Style>"
End Sub
Private Sub Sep()
    Send "<P align=center class=chica>~才才才才才才才才才才才才才才才才才才才才才才才才才才才才才</P>"
End Sub
