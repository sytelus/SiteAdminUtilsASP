Attribute VB_Name = "CGI4VB"
' CGI routines used with VB (32bit) using STDIN / STDOUT.
'
' Version: 1.5 (September 1997)
'
' Author:  Kevin O'Brien <obrienk@pobox.com>
'                        <obrienk@ix.netcom.com>
'
'Function LeerFile() and Function Guardar() and other Modify webmaster@pitic.net
Option Explicit
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Const FILE_BEGIN = 0&
Public CGI_Accept            As String: Public CGI_AuthType         As String
Public CGI_ContentLength     As String: Public CGI_ContentType      As String
Public CGI_Cookie            As String: Public CGI_GatewayInterface As String
Public CGI_PathInfo          As String: Public CGI_PathTranslated   As String
Public CGI_QueryString       As String: Public CGI_Referer          As String
Public CGI_RemoteAddr        As String: Public CGI_RemoteHost       As String
Public CGI_RemoteIdent       As String: Public CGI_RemoteUser       As String
Public CGI_RequestMethod     As String: Public CGI_ScriptName       As String
Public CGI_ServerSoftware    As String: Public CGI_ServerName       As String
Public CGI_ServerPort        As String: Public CGI_ServerProtocol   As String
Public CGI_UserAgent         As String
Public lContentLength As Long
Public hStdIn         As Long
Public hStdOut        As Long
Public sErrorDesc     As String
Public sEmail         As String
Public sFormData      As String
Public FileNombre As String
Public ContenidoType As String
Public Valor  As String
Public ArchPos As String
Public Nombre_Archivo As String
Public Ind
Type pair
  Name As String
  Value As String
End Type
Public tPair() As pair
Public Ruta As String
Public Fso, TxtStream
Public NmUser As String
Public Vali As Boolean
Public SVali As String
Public Per As String
Public BackG As String
Global FwdURL As String
Global DEFX As String
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Sub Main()
    On Error GoTo ErrorRoutine
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Vali = True
    InitCgi
    GetFormData
    BackG = GetCgiValue("Back")
    Ind = GetCgiValue("UpFile")
    NmUser = GetCgiValue("NomUser")
    Per = GetCgiValue("xnos")
    If NmUser = "" Then
        NoAut
        End
    End If
    If Not Per = CStr(Hex(((Len(NmUser) * Date) * 28))) Then
        NoAut
        End
    End If
    If Ind = "True" Then
        LeerFile
        Guardar
    End If
    Cgi_Main
EndPgm:
       End
ErrorRoutine:
       sErrorDesc = Err.Description & " Error Number = " & Str$(Err.Number)
       ErrorHandler
       Resume EndPgm
End Sub
Sub InitCgi()
    hStdIn = GetStdHandle(STD_INPUT_HANDLE)
    hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
    sEmail = "webmaster@pitic.net"
    CGI_Accept = Environ("HTTP_ACCEPT")
    CGI_AuthType = Environ("AUTH_TYPE")
    CGI_ContentLength = Environ("CONTENT_LENGTH")
    CGI_ContentType = Environ("CONTENT_TYPE")
    CGI_Cookie = Environ("HTTP_COOKIE")
    CGI_GatewayInterface = Environ("GATEWAY_INTERFACE")
    CGI_PathInfo = Environ("PATH_INFO")
    CGI_PathTranslated = Environ("PATH_TRANSLATED")
    CGI_QueryString = Environ("QUERY_STRING")
    CGI_Referer = Environ("HTTP_REFERER")
    CGI_RemoteAddr = Environ("REMOTE_ADDR")
    CGI_RemoteHost = Environ("REMOTE_HOST")
    CGI_RemoteIdent = Environ("REMOTE_IDENT")
    CGI_RemoteUser = Environ("REMOTE_USER")
    CGI_RequestMethod = Environ("REQUEST_METHOD")
    CGI_ScriptName = Environ("SCRIPT_NAME")
    CGI_ServerSoftware = Environ("SERVER_SOFTWARE")
    CGI_ServerName = Environ("SERVER_NAME")
    CGI_ServerPort = Environ("SERVER_PORT")
    CGI_ServerProtocol = Environ("SERVER_PROTOCOL")
    CGI_UserAgent = Environ("HTTP_USER_AGENT")
    lContentLength = Val(CGI_ContentLength)
    ReDim tPair(0)
End Sub
Sub GetFormData()
    Dim sBuff        As String
    Dim lBytesRead   As Long
    Dim rc           As Long
    If CGI_RequestMethod = "POST" Then
       sBuff = String(lContentLength, Chr$(0))
       Do While Len(sFormData) < lContentLength
          rc = ReadFile(hStdIn, ByVal sBuff, lContentLength, lBytesRead, ByVal 0&)
          sFormData = sFormData & Left$(sBuff, lBytesRead)
       Loop
    End If
    StorePairs CGI_QueryString
End Sub
Sub LeerFile()
Dim TBytes, URLData, PodIni, PosFin, DBoundPos, DBound, Lugar, Nombre
Dim PosFile, PosDBound, FileCount, FileName, FieldCount, ContentType, Value
TBytes = lContentLength
URLData = sFormData
PodIni = 1
PosFin = InStr(PodIni, URLData, Chr(13))
DBound = Mid(URLData, PodIni, PosFin - PodIni)
DBoundPos = InStr(1, URLData, DBound)
Do Until (DBoundPos = InStr(URLData, DBound & "--"))
    Lugar = InStr(DBoundPos, URLData, "Content-Disposition")
    Lugar = InStr(Lugar, URLData, "name=")
    PodIni = Lugar + 6
    PosFin = InStr(PodIni, URLData, Chr(34))
    Nombre = Mid(URLData, PodIni, PosFin - PodIni)
    PosFile = InStr(DBoundPos, URLData, "filename=")
    PosDBound = InStr(PosFin, URLData, DBound)
    If PosFile <> 0 And (PosFile < PosDBound) Then
        FileCount = FileCount + 1
        PodIni = PosFile + 10
        PosFin = InStr(PodIni, URLData, Chr(34))
        FileName = Mid(URLData, PodIni, PosFin - PodIni)
        FileNombre = FileName
        Lugar = InStr(PosFin, URLData, "Content-Type:")
        PodIni = Lugar + 14
        PosFin = InStr(PodIni, URLData, Chr(13))
        ContentType = Mid(URLData, PodIni, PosFin - PodIni)
        ContenidoType = ContentType
        PodIni = PosFin + 4
        PosFin = InStr(PodIni, URLData, DBound) - 2
        Value = Mid(URLData, PodIni, PosFin - PodIni)
    Else
        FieldCount = FieldCount + 1
        Lugar = InStr(Lugar, URLData, Chr(13))
        PodIni = Lugar + 4
        PosFin = InStr(PodIni, URLData, DBound) - 2
        Value = Mid(URLData, PodIni, PosFin - PodIni)
    End If
    If Not Value = "~Send~" Then Valor = Valor & Value
    DBoundPos = InStr(DBoundPos + LenB(DBound), URLData, DBound)
Loop
ArchPos = InStrRev(FileNombre, "\")
Nombre_Archivo = Right(FileNombre, Len(FileNombre) - ArchPos)
End Sub
Sub Guardar()
Dim Buc, Nue As String, Txt As String, A As Integer, Ext As String
    Ruta = Left(App.Path, InStrRev(App.Path, "\"))
    Ruta = Ruta & NmUser & "\"
    Nue = ""
    For A = 1 To Len(Nombre_Archivo)
        Txt = Mid(Nombre_Archivo, A, 1)
        Txt = LCase(Txt)
        Select Case Txt
        Case "á": Txt = "a"
        Case "é": Txt = "e"
        Case "í": Txt = "i"
        Case "ó": Txt = "o"
        Case "ú": Txt = "u"
        End Select
        Nue = Nue & Txt
    Next
    Nombre_Archivo = Nue
    Ext = Fso.GetExtensionName(Nombre_Archivo)
    Select Case LCase(Ext)
    Case "htm", "html", "txt", "zip", "mid", "jpg", "jpeg", "gif", "bmp"
        Vali = True
    Case Else
        Vali = False
    End Select
    If lContentLength > 800000 Then: Vali = False
    If Vali = True Then
        Set TxtStream = Fso.CreateTextFile(Ruta & Nombre_Archivo, True)
        TxtStream.Write (Valor)
        TxtStream.Close
    End If
End Sub
Sub StorePairs(sData As String)
Dim pointer As Long, n As Long, delim1 As Long, delim2 As Long, lastPair As Long, lPairs As Long
lastPair = UBound(tPair)
delim1 = 0
Do
  delim1 = InStr(delim1 + 1, sData, "=")
  If delim1 = 0 Then Exit Do
  lPairs = lPairs + 1
Loop
If lPairs = 0 Then Exit Sub
ReDim Preserve tPair(lastPair + lPairs) As pair
pointer = 1
For n = (lastPair + 1) To UBound(tPair)
   delim1 = InStr(pointer, sData, "=")
   If delim1 = 0 Then Exit For
   tPair(n).Name = UrlDecode(Mid$(sData, pointer, delim1 - pointer))
   delim2 = InStr(delim1, sData, "&")
   If delim2 = 0 Then delim2 = Len(sData) + 1
   tPair(n).Value = UrlDecode(Mid$(sData, delim1 + 1, delim2 - delim1 - 1))
   pointer = delim2 + 1
Next n
End Sub
Public Function UrlDecode(ByVal sEncoded As String) As String
Dim pos As Long
If sEncoded = "" Then Exit Function
pos = 0
Do
   pos = InStr(pos + 1, sEncoded, "+")
   If pos = 0 Then Exit Do
   Mid$(sEncoded, pos, 1) = " "
Loop
pos = 0
On Error GoTo errorUrlDecode
Do
   pos = InStr(pos + 1, sEncoded, "%")
   If pos = 0 Then Exit Do

   Mid$(sEncoded, pos, 1) = Chr$("&H" & (Mid$(sEncoded, pos + 1, 2)))
   sEncoded = Left$(sEncoded, pos) _
             & Mid$(sEncoded, pos + 3)
Loop
On Error GoTo 0
UrlDecode = sEncoded
Exit Function
errorUrlDecode:
If Err.Number = 13 Then
   Err.Clear
   Err.Raise 65001, , "Invalid data passed to UrlDecode() function."
Else
   Err.Raise Err.Number
End If
Resume Next
End Function
Function GetCgiValue(cgiName As String) As String
Dim n As Integer
For n = 1 To UBound(tPair)
    If UCase$(cgiName) = UCase$(tPair(n).Name) Then
       If GetCgiValue = "" Then
          GetCgiValue = tPair(n).Value
       Else
          GetCgiValue = GetCgiValue & ";" & tPair(n).Value
       End If
    End If
Next n
End Function
Sub SendHeader(sTitle As String)
    Send "Status: 200 OK"
    Send "Content-type: text/html" & vbCrLf
    Send "<!-- Inicio Sitio Web -->"
    Send "<HTML><HEAD><TITLE>" & sTitle & "</TITLE>"
    Send "</HEAD>"
End Sub
Sub SendHeader2(sTitle As String, sDef As String, sTime As String)
    FwdURL = "<META HTTP-EQUIV=""REFRESH"" CONTENT=""" & sTime & ";"
    FwdURL = FwdURL & " "
    FwdURL = FwdURL & "URL="
    FwdURL = FwdURL & sDef & """"
    FwdURL = FwdURL & ">"
    Send "Status: 200 OK"
    Send "Content-type: text/html" & vbCrLf
    Send "<!-- Inicio Sitio Web -->"
    Send "<HTML><HEAD><TITLE>" & sTitle & "</TITLE>"
    Send FwdURL
    Send "</HEAD>"
End Sub
Sub SendFooter()
    Send "</BODY></HTML>"
    Send "<!-- Fin Sitio Web -->"
End Sub
Sub Send(s As String)
    Dim rc            As Long
    Dim lBytesWritten As Long
    s = s & vbCrLf
    rc = WriteFile(hStdOut, s, Len(s), lBytesWritten, ByVal 0&)
End Sub
Sub SendB(s As String)
    Dim rc            As Long
    Dim lBytesWritten As Long
    rc = WriteFile(hStdOut, s, Len(s), lBytesWritten, ByVal 0&)
End Sub
Sub ErrorHandler()
    Dim rc As Long
    On Error Resume Next
    rc = SetFilePointer(hStdOut, 0&, 0&, FILE_BEGIN)
    SendHeader "Error Interno"
    Send "<H1>Error en " & CGI_ScriptName & "</H1>"
    Send "Ocurrio el Siguiente Error:"
    Send "<PRE>" & sErrorDesc & "</PRE>"
    Send "<I>Por favor</I> anote en que pagina estaba en el momento en ocurrio el error, "
    Send "para porder identificarlo correctamente. "
    Send "Contacte al Administrador de Sistema: "
    Send "<A HREF=""mailto:" & sEmail & """>"
    Send "<ADDRESS>&lt;" & sEmail & "&gt;</ADDRESS></A>"
    SendFooter
    rc = SetEndOfFile(hStdOut)
End Sub
