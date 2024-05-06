Attribute VB_Name = "FairWeb"
Option Explicit

'Public gFwChkVer As Boolean
Public gFwChkActive As Boolean

Type WebTag_t
   Name   As String
   Value As String
End Type

Type WebComm_t
   Rc          As Integer

   Prod        As String
   AppName     As String
   Url         As String
   ExtraInfo   As String      ' optional
   Param       As String      ' optional

   RBuf        As String
   Tag(7)      As WebTag_t
End Type

' Tags
Public Const TG_VER = 0
Public Const TG_DATE = 1
Public Const TG_UMSG = 2
Public Const TG_MSG = 3
Public Const TG_QURL = 4
Public Const TG_URL = 5
Public Const TG_PURL = 6
Public Const TG_UNR = 7

Private lWebComm As WebComm_t

Public gChecked As Boolean

Public Function FwWebGetTag(WebComm As WebComm_t, ByVal TagName As String) As String
   Dim i As Integer

   FwWebGetTag = ""
   
   For i = 0 To UBound(WebComm.Tag)
      If StrComp(WebComm.Tag(i).Name, TagName, vbTextCompare) = 0 Then
         FwWebGetTag = WebComm.Tag(i).Value
         Exit For
      End If
   Next i

End Function


Public Function FwWebAppVer(ByVal Prod As String, Info() As WebTag_t) As String
   Dim RBuf As String, Page As String
   Dim Rc As Long, Dt As Long, i As Integer, j As Integer, k As Integer
   Dim HostFair1 As String, UrlFair1 As String

'   Debug.Print "Host=" & FwEncrypt1("            www.fairware.cl      ", 97125)
'   Debug.Print "Host=" & FwEncrypt1("            https://servicioslp.thomsonreuters.cl      ", 97125)
'   Debug.Print "Page=" & FwEncrypt1("            /LastVer.asp?P=      ", 97125)
'    HostFair1 = Trim(FwDecrypt1("682767286A2D716A6DA6A9A0A99D979BA831F3EDE864388D633A926B45207C5937", 97125))
    HostFair1 = Trim(FwDecrypt1("8B345E89356290736083E2C3B29374E6CCB7A28A79E7DD487AEBE1D1C6C2C5C3B5ACB1A966A4D8D4D4D1CF5A2E83593088613B96724F2D", 97125))
    UrlFair1 = Trim(FwDecrypt1("35743475377A3E667CA49C929BD8D25AF2ED6A7AA064388D633A926B45207C5937", 97125))
   
   If W.InDesign Then
'      HostFair1 = "192.168.220.11"
'      UrlFair1 = "/Fairware/LastVer.asp?P="
      HostFair1 = "servicioslp.thomsonreuters.cl"
      UrlFair1 = "/LastVer.asp?P="
   End If

   FwWebAppVer = ""
   
   Page = FwWebReadPage(HostFair1, UrlFair1 & Prod, 5000)
   
   If Page = "" Then
      Exit Function
   End If
      
   Page = ReplaceStr(Page, "\n", vbCrLf)
   Page = ReplaceStr(Page, "\t", vbTab)
   
   RBuf = ""
   For i = 0 To UBound(Info)
      j = InStr(1, Page, "<" & Info(i).Name & ">", vbTextCompare)
      If j > 0 Then
         j = j + 2 + Len(Info(i).Name)
         
         k = InStr(j + 2, Page, "</" & Info(i).Name & ">", vbTextCompare)
      
         If k > 0 Then
            Info(i).Value = Mid(Page, j, k - j)
            RBuf = RBuf & Info(i).Value
         End If
      End If
   Next i
           
   FwWebAppVer = RBuf
           
End Function


Public Function FwWebAppVer_old(ByVal Prod As String, Info() As WebTag_t) As String
   Dim hInternetSession As Long, hInternetConnect As Long, hHttpOpenRequest As Long
   Dim bDoLoop As Boolean
   Dim sReadBuffer As String * 50, lNumberOfBytesRead As Long
   Dim Buf As String, RBuf As String
   Dim Rc As Long, Dt As Long, i As Integer, j As Integer, k As Integer
   Const UrlFair1 = "ww" & "w.f" & "air" & "war" & "e.cl"
   Const UrlFair2 = "/LastVer.asp?P="
'   Const UrlFair1 = "fairware.ath.cx"
'   Const UrlFair2 = "/Fairware/LastVer.asp?P="

   FwWebAppVer_old = ""
   
   hInternetSession = InternetOpen("FairVer", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
   If hInternetSession = 0 Then
      Exit Function
   End If
   
   DoEvents
   
   Rc = 5000
   Rc = InternetSetOption(hInternetSession, INTERNET_OPTION_CONNECT_TIMEOUT, Rc, Len(Rc))
   
   Rc = 1
   Rc = InternetSetOption(hInternetSession, INTERNET_OPTION_IGNORE_OFFLINE, Rc, Len(Rc))
   
   DoEvents
   
   hInternetConnect = InternetConnect(hInternetSession, UrlFair1, INTERNET_DEFAULT_HTTP_PORT, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
   If hInternetConnect = 0 Then
      Exit Function
   End If
   
   DoEvents

   hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", UrlFair2 & Prod, "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION, 0)
   If hHttpOpenRequest = 0 Then
      Exit Function
   End If

   DoEvents

   Buf = ""
   Rc = 0
   Rc = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, Buf, Rc)
   
   DoEvents
   
   Buf = ""
   If Rc Then
      Do
         sReadBuffer = vbNullString
         bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
         Buf = Buf & Left(sReadBuffer, lNumberOfBytesRead)
         
         DoEvents
      Loop While lNumberOfBytesRead
   End If

   InternetCloseHandle (hHttpOpenRequest)
   InternetCloseHandle (hInternetConnect)
   InternetCloseHandle (hInternetSession)
   
   Buf = ReplaceStr(Buf, "\n", vbCrLf)
   Buf = ReplaceStr(Buf, "\t", vbTab)
   
   RBuf = ""
   For i = 0 To UBound(Info)
      j = InStr(1, Buf, "<" & Info(i).Name & ">", vbTextCompare)
      If j > 0 Then
         j = j + 2 + Len(Info(i).Name)
         
         k = InStr(j + 2, Buf, "</" & Info(i).Name & ">", vbTextCompare)
      
         If k > 0 Then
            Info(i).Value = Mid(Buf, j, k - j)
            RBuf = RBuf & Info(i).Value
         End If
      End If
   Next i
           
   FwWebAppVer_old = RBuf
           
End Function
' Verifica si hay una versión más reciente
' por si hay problemas con la conexión probamos sólo un par de veces en el día
' si bMsg = true es porque lo solicitó el usuario
Public Function FwCheckVersion(Frm As Form, ByVal bMsg As Boolean, ByVal AppName As String, ByVal Url As String) As Boolean
   Dim d As Double, n As Integer, H As Long, Rut As String

   Call AddDebug("FwCheckVer1 " & AppName & ", " & FwGetPcCode() & ", " & gAppCode.Demo)

   FwCheckVersion = False
   If gChecked = False Then ' Basta con una vez
      
      d = GetIniString(gIniFile, "Config", "I1", "0")
      n = GetIniString(gIniFile, "Config", "I2", "0")
      H = GetIniString(gIniFile, "Config", "I3", "0")
      
      ' si ya intentó más de dos veces y no pudo, asumumos que si, mañana será otro día
      If bMsg = False And d = CLng(Int(Now)) And n > 2 Then ' intentamos hasta dos veces en el día
         gChecked = True
      
      Else
         If d <> CLng(Int(Now)) Or bMsg = True Then
            n = 0
            H = Hour(Now) * 60 + Minute(Now)
         End If
         
         If Hour(Now) * 60 + Minute(Now) >= H Then
            Call SetIniString(gIniFile, "Config", "I1", CLng(Int(Now)))
            Call SetIniString(gIniFile, "Config", "I2", n + 1)
            Call SetIniString(gIniFile, "Config", "I3", Int(H + 70 + (20 * Rnd))) ' prueba en 95 minutos
            
            If gAppCode.Rut <> "" Then
               Rut = gAppCode.Rut
            Else
               Rut = GetIniString(gCfgFile, "Config", "RUT")
            End If
                        
            'gChecked = FwCheckVer(Frm, "FairPay1", App.Title, "http://www.fairware.cl/FairPay.asp", , "&r=" & GetIniString(gCfgFile, "Config", "RUT") & "&cpc=" & FwGetPcCode() & "&d=" & Abs(gAppCode.Demo) & "&ver=" & W.Version & "&fver=" & Format(W.FVersion, "yyyymmdd"), bMsg)
            gChecked = FwCheckVer(Frm, AppName, App.Title, Url, , "&r=" & Rut & "&cpc=" & FwGetPcCode() & "&d=" & Abs(gAppCode.Demo) & "&ver=" & W.Version & "&fver=" & Format(W.FVersion, "yyyymmdd"), bMsg)
            FwCheckVersion = gChecked
            Call AddDebug("FwCheckVer2 " & gAppCode.Demo)
         End If
         
      End If
      
   End If
      
End Function

Public Function FwCheckVer(Frm As Form, ByVal Prod As String, ByVal AppName As String, ByVal Url As String, Optional ByVal ExtraInfo As String = "", Optional ByVal WebInfo As String = "", Optional ByVal bMsg As Boolean = 1) As Boolean
   Dim bAnsw As Boolean, i As Integer

   lWebComm.Prod = Prod
   lWebComm.AppName = AppName
   lWebComm.Url = Url
   lWebComm.ExtraInfo = ExtraInfo
   lWebComm.Param = WebInfo
   lWebComm.RBuf = ""

   Call FwWebComm(Frm, lWebComm, bMsg)

   bAnsw = (lWebComm.RBuf <> "") ' hubo respuesta ??
   FwCheckVer = bAnsw

   If bAnsw = True And gAppCode.Demo = False Then
      ' 8 sep 2006
      
      If lWebComm.Tag(TG_UNR).Value = ".s." Then
         Call FwUnRegister
         gAppCode.Demo = True
      End If
      
'      For i = 0 To UBound(lWebComm.Tag)
'         If lWebComm.Tag(i).Name = "unr" Then
'            If lWebComm.Tag(i).Value = ".s." Then
'               Call FwUnRegister
'               gAppCode.Demo = True
'            End If
'            Exit For
'         End If
'      Next i
   End If

End Function

Public Sub FwWebComm(Frm As Form, WebComm As WebComm_t, Optional ByVal bMsg As Boolean = 1)
   Dim PVer As String, PFecha As String, Msg As String, Url As String
   'Dim Info(7) As WebInfo_t
   Dim Rc As Long, i As Integer, WebInfo As String
   
   For i = 0 To UBound(WebComm.Tag)
      WebComm.Tag(i).Value = ""
      WebComm.Tag(i).Name = ""
   Next i
      
   WebComm.Tag(TG_VER).Name = "ver"
   WebComm.Tag(TG_DATE).Name = "date"
   WebComm.Tag(TG_UMSG).Name = "UMsg"
   WebComm.Tag(TG_MSG).Name = "Msg"
   WebComm.Tag(TG_QURL).Name = "QUrl"
   WebComm.Tag(TG_URL).Name = "Url"
   WebComm.Tag(TG_PURL).Name = "PUrl"
   WebComm.Tag(TG_UNR).Name = "unr"
   
   WebInfo = WebComm.Param & "&pc=" & GetComputerName() & "&mac=" & GetMac()

   WebComm.RBuf = FwWebAppVer(WebComm.Prod & WebInfo, WebComm.Tag())
   
   If bMsg = True And WebComm.RBuf = "" And modWinInet.gWLastDllError = ERROR_FILE_NOT_FOUND Then
      MsgBox1 "No se pudo verificar si hay una actualización." & vbCrLf & "Verifique en Internet Explorer que no esté trabajando desconectado (Offline).", vbExclamation
      Exit Sub
   End If
   
   If WebComm.RBuf <> "" And bMsg = True Then ' se conectó y hay que informar ?
            
      PVer = App.Major & "." & App.Minor & "." & App.Revision
      PFecha = Right(Trim(App.ProductName), 8)
      
      WebComm.Tag(TG_PURL).Value = ReplaceStr(WebComm.Tag(TG_PURL).Value, " ", "%20") & "&cv=" & PVer & "&cd=" & PFecha & WebComm.ExtraInfo

      ' nueva versión ?
      If NormVer(WebComm.Tag(TG_VER).Value) > NormVer(PVer) Or WebComm.Tag(TG_DATE).Value > PFecha Then
         Msg = WebComm.Tag(TG_UMSG).Value & vbLf & WebComm.Tag(TG_MSG).Value
         
         If WebComm.Tag(TG_QURL).Value <> "" And WebComm.Tag(TG_URL).Value <> "" Then
            Msg = Msg & WebComm.Tag(TG_QURL).Value
                        
            If MsgBox1(Utf8Ansi(Msg), vbQuestion Or vbYesNo) = vbYes Then
               Rc = ShellExecute(Frm.hWnd, "open", WebComm.Tag(TG_URL).Value, "", "", 1)
            End If
         ElseIf Len(Msg) > 2 Then
            MsgBox1 Utf8Ansi(Msg), vbInformation
         End If
         
      ElseIf Trim(WebComm.Tag(TG_MSG).Value) <> "" Then
         MsgBox1 Utf8Ansi(WebComm.Tag(TG_MSG).Value), vbInformation
      End If
      
   End If

End Sub

' Consulta en fairware.cl si tiene licencia y el nivel
' 10 jul 2012
Public Function FwChkActive(ByVal When As Integer) As Integer
   Static RcChecked As Integer, wsErr As Long, wsDescr As String, oDemo As Boolean
   Dim Info As String, IP As String, Path As String, Page As String, Resp As String, Rc As Integer, MsgIni As String, AskIni As String
   Dim iAux As Long, SAux As String, PcCode As String, i As Integer, Chk As Long
   
   oDemo = gAppCode.Demo
   
   If gAppCode.Demo Then   ' sólo se controlan los activos
      FwChkActive = vbYes
      Exit Function
   End If
   
   If gFwChkActive Then ' Ya verificó en la sesión
      FwChkActive = RcChecked
      Exit Function
   End If
   
   If gAppCode.Rut = "" Then
      FwChkActive = vbYes
      Call AddLog("Demo: falta el Rut")
      Exit Function
   End If
         
   If gAppCode.Name = "" Then
      MsgBox "Falta asignar gAppCode.Name con APP_NAME", vbCritical
      Exit Function
   End If
   
'   Debug.Print FwEncrypt1("   www.fairware.cl    ", 63197)
'   Debug.Print FwEncrypt1("   https://servicioslp.thomsonreuters.cl    ", 63197)
'   Debug.Print FwEncrypt1("    /FwProgInfo_.asp    ", 93591)
'   IP = Trim(FwDecrypt1("915B2672BFC14D8F9B9BADAAADBAD062ADB0B4399674", 63197))
   IP = Trim(FwDecrypt1("42762B6198847AA68E78F0DAC4BFAEA296877F767569A49E9D969499A5ACA7A7B5367C4380858E949B2F8C6A", 63197))
   Path = Trim(FwDecrypt1("3C844D97E2DED865A6B7B9C54A75838A8FB7B96D39957250", 93591))
   
   PcCode = FwGetPcCode()
   
'   If W.InDesign Then
'      gAppCode.Rut = "78089800-3" ' Rucantu S.A.
'      W.PcName = "ADQUISICIONES"
'      PcCode = "YBIHZMMMUIU"
'   End If
   
   Info = "?an=" & gAppCode.Name & "&rc=" & gAppCode.Rut & "&pc=" & W.PcName & "&mc=" & W.Mac & "&cp=" & PcCode
   Call AddDebug("368: Info=[" & ReplaceStr(Mid(Info, 2), "&", "; ") & "]")
   
   Info = Info & "&pv=" & W.Version & "&pf=" & W.FVersion
   
   If W.InDesign Then
      Info = Info & "&ts=1"
   End If
   
   Info = Info & "&fu=" & gAppCode.FUsoVersion ' FPrimUso debe ir despues de ts
   
   iAux = GenClave(Info & "#", 74731)
'   Info = Info & "&ck=" & iAux
   
'   If w.InDesign Then
'      IP = "localhost"
'      Path = "/fairware/FwProgInfo_.asp"
''      Path = "/FwProgInfo__.asp"
'   End If
   
   Debug.Print "url: " & IP & Path & Info & "&ck=" & iAux
   Page = FwWebReadPage(IP, Path & Info & "&ck=" & iAux)
   If Len(Page) < 10 Then
      FwChkActive = vbYes
      gAppCode.Demo = (Val(GetIniString(gIniFile, "Config", "WaitPrt4", "0")) Mod 2) ' si no logró conectarse recordamos la última respuesta
      Call AddLog("393: Demo: Antes: " & oDemo & ", después: " & gAppCode.Demo & ", Pc: " & W.PcName & ", PcCode: " & PcCode & ";")
      Exit Function
   End If
   
   Resp = FwGetXmlTag(Page, "Info")
   
   If Len(Resp) < 10 Then
'      Call AddDebug("400: WPage=[" & Page & "]")
      FwChkActive = vbYes
      gAppCode.Demo = (Val(GetIniString(gIniFile, "Config", "WaitPrt4", "0")) Mod 2) ' si no logró conectarse recordamos la última respuesta
      Call AddLog("403: Demo: Antes: " & oDemo & ", después: " & gAppCode.Demo & ", Pc: " & W.PcName & ", PcCode: " & PcCode & ".")
      Exit Function
   End If
   
'   Call AddDebug("C_Resp=[" & Resp & "]")

   wsErr = Val(FwGetXmlTag(Resp, "Err"))
   wsDescr = FwGetXmlTag(Resp, "Descr")
   
   gAppCode.Demo = Val(FwGetXmlTag(Resp, "D" & "e" & "m" & "o"))
   gAppCode.UnReg = Val(FwGetXmlTag(Resp, "UnReg"))
   
   If gAppCode.Demo Then
      Call AddDebug("416: Demo: [" & Page & "]")
   End If
   
   SAux = FwGetXmlTag(Resp, "Level")
   If SAux <> "" Then
      iAux = Val(SAux)
      If iAux > 0 And iAux <> gAppCode.NivProd Then
         For i = 0 To UBound(gAppCode.Nivel)
            If gAppCode.Nivel(i).Id <= 0 Then
               Exit For
            ElseIf iAux = gAppCode.Nivel(i).Id Then
               gAppCode.NivProd = iAux
               Exit For
            End If
         Next i
      End If
   End If
   
    ' verificamos que sea el sitio de Fairware el que responde
   Info = Info & "&Demo=" & Abs(gAppCode.Demo) & "&Level=" & SAux

   Chk = Val(FwGetXmlTag(Resp, "Chk"))
   iAux = GenClave(Info & "#", 93157)
   
   If Chk <> iAux Then
      gAppCode.Demo = 1
   End If
   
   If gAppCode.NivProd = 0 Then
      gAppCode.NivProd = gAppCode.NivDef
   End If
   
   gAppCode.MinMsg = Val(FwGetXmlTag(Resp, "MinMsg"))
   gAppCode.Msg = FwGetXmlTag(Resp, "Msg")
   
   MsgIni = FwGetXmlTag(Resp, "MsgIni")
   AskIni = FwGetXmlTag(Resp, "AskIni")
   
   gAppCode.Msg = ReplaceStr(gAppCode.Msg, "\n", vbCrLf)
   MsgIni = ReplaceStr(MsgIni, "\n", vbCrLf)
   AskIni = ReplaceStr(AskIni, "\n", vbCrLf)

   RcChecked = vbYes
   gFwChkActive = True

   ' Se guarda por si la próxima vez no hay conexión a Internet
   Call SetIniString(gIniFile, "Config", "WaitPrt4", IIf(gAppCode.Demo, 5, 2))

   If gAppCode.Demo And gAppCode.UnReg Then
      Call FwUnRegister
   End If

   If When = 0 Then
      If MsgIni <> "" Then  ' al inicio del programa
         MsgBox1 Utf8Ansi(MsgIni), vbExclamation
      End If
   
      If AskIni <> "" Then  ' al inicio del programa
         RcChecked = MsgBox1(Utf8Ansi(AskIni) & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbExclamation Or vbYesNo Or vbDefaultButton2)
      End If
   End If

'   Call AddDebug("C_Demo=[" & gAppCode.Demo & "]")

   If oDemo = False And gAppCode.Demo Then
      Call AddLog("481: Demo: Antes: " & oDemo & ", después: " & gAppCode.Demo & ", Pc: " & W.PcName & ", PcCode: " & PcCode)
   End If

   FwChkActive = RcChecked
   
End Function
