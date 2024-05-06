Attribute VB_Name = "ModAcepta"
Option Explicit

' Requiere modWinINet.bas

'Datos fijos de Acepta
'Private Const URL_ACEPTA = "https://escritorio.acepta.com"
Private Const URL_ACEPTA = "https://fairware.acepta.com"


'URL de Produccion Tecnoback
Private Const URL_TECNOBACK = "https://api.tecnoback.com/"
' Para certificación
Private Const URL_ACEPTACERT = "https://escritorio-cert.acepta.com"
Private Const URL_TECNOBACKCERT = "https://apicert.tecnoback.com/"

' Lo que viene despues de la URL
Private Const PURL_LOGIN = "/pivote.php"
Private Const PURL_RECIBIDOS = "/ext.php?r=https://fairware.acepta.com/appDinamicaClasses/index.php%3Fapp_dinamica=buscarNEW_recibidos%26identificador=NO_BUSCAR%26"

'Rut y Clave de desarrollo Tecnoback Utilizar en Fuente
Private Const RUT_USR_TECNOBACKDESA = "77049060-K"
Private Const CLAVE_USR_TECNOBACKDESA = "ca248bdc97d6c37fc195968eaa43dc7b"

'Rut y Clave de Produccion Tecnoback
Private Const RUT_USR_TECNOBACK = "79755150-3"
Private Const CLAVE_USR_TECNOBACK = "9a0cdbb3f5671f792264d404c0d8cb4f"

Public RutFirma As String

Public Const RUT_EMP_ACEPTA = 96919050
Public Const RUT_EMP_TECNOBACK = 77049060
Private Const RUT_USR_ACEPTA = 17597643
Private Const CLAVE_USR_ACEPTA = "dani2105"
Private Const CLAVE_CERT_ACEPTA = "dani2116"
Private Const MAIL_EMISOR_ACEPTA = "franca@fairware.cl"

Private Const URL_FAIR = "http://www.fairware.cl"

Private Const AERR_OK = 1
Private Const AERR_NORESP = 1000    ' No hay respuesta desde Acepta
Private Const AERR_ERR1 = 1001      ' Error de algun tipo
Private Const AERR_BADUSER = 1002   ' Usuario o clave incorrecto
Private Const AERR_ERR3 = 1003      ' Error de algun tipo
Public Const AERR_MOTOR = 1004      ' Se fue por Timeout al firmar
Private Const AERR_NOPARAM = 1005   ' Falta algun parámetro

Public Type AcpLogin_t

   ' Datos de entrada
   Host        As String
   Url         As String   ' Url base para emitir facturas
   UrlRec      As String   ' Url base para recibidos
   
   'RutUsr      As Long     ' sin DV
   RutUsr      As String     ' sin DV
   Pasw        As String   ' Para autenticarse
   PaswCert    As String   ' Para firmar (clave del certificado)
   RutEmpr     As Long     ' sin DV
   
   MailEmisor  As String
   SubjMail    As String
   
   ' Datos Calculados
   SessionId   As String
   md5Pasw     As String
   
   UrlRecibidos As String
End Type

Public Type AcpDTE_t

   ' Entrada
   xml         As String   ' del DTE
   MailRecep   As String   ' Mail del receptor
   bExport     As Boolean  ' 2 oct 2020: se agrega para no reclamar por el mail del receptor cuando es exportación
   IdDTE       As Long     ' 31 dic 2020: para identificarlo en el caso que falle la firma

   ' Salida
   Folio       As Long
   UrlDoc      As String
   ObsDTE      As String
   
End Type

' Estado
Public Const ET_INFO As Integer = 50
Public Const ET_OK As Integer = 1
Public Const ET_WARN As Integer = 2
Public Const ET_ERR As Integer = 3

' Errores
Public Const AERR_AUT_OK = 1     ' Ok para autenticar
Public Const AERR_FIRM_ERR = 2   ' Error al firmar
Public Const AERR_FIRM_OK = 3    ' Ok para firmar
Public Const AERR_SESIONINVALIDA = 666 'indica que está conectado al escritorio de Acepta mienstras está friamndo un documento

Public Type DetTraza_t
   Tipo     As String
   UrlTipo  As String
   Estado   As Integer
   Fecha    As Double
   Obs      As String
End Type

Public Type AcpTraza_t

   ' Entrada
   Url           As String
   
   ' Salida
   Doc            As String
   Folio          As Long
   Fecha          As Double
   
   nErr           As Integer
   nWarn          As Integer
   
   nAvisos        As Integer
   nAcepta        As Integer
   nSII           As Integer
   nIntercam      As Integer
   nControl       As Integer
   nMandat        As Integer
   nIECV          As Integer
   UrlData        As String
   
   Avisos(20)     As DetTraza_t
   Acepta(20)     As DetTraza_t
   SII(20)        As DetTraza_t
   Intercam(20)   As DetTraza_t
   Controll(20)   As DetTraza_t
   Mandat(20)     As DetTraza_t
   IECV(20)       As DetTraza_t
   
End Type

Public Type AcpTrazaEvento_t

   ' Entrada
   RutEmisor       As String
   TipoDTE         As String
   Folio           As String
   canal           As String
   RUTRecep        As String

   ' Salida
   Doc            As String
   'Folio          As Long
   Fecha          As Double
   
   nErr           As Long
   nWarn          As Long
   
   nAvisos        As Integer
   nAcepta        As Integer
   nSII           As Integer
   nIntercam      As Integer
   nControl       As Integer
   nMandat        As Integer
   nIECV          As Integer
   
   Avisos(10)     As DetTraza_t
   Acepta(10)     As DetTraza_t
   SII(10)        As DetTraza_t
   Intercam(10)   As DetTraza_t
   Controll(10)   As DetTraza_t
   Mandat(10)     As DetTraza_t
   IECV(10)       As DetTraza_t
   
End Type

Public Sub AcpTest()
   Dim Lg As AcpLogin_t, Rc As Long, DTE As AcpDTE_t
   Dim xml As String
   
'   curl -X POST --data '{"rut":17597643-4,"pass":"dani2105","cod_md5":"c18447e58b032bae5ff2206b424af97d","empresa":96919050,"tipo_tx":"sesion_externa","__FLAG_AUTENTIFICACION__":"SI"}'  https://escritorio_cert.acepta.com/pivote.php
   
   If W.InDesign = False Then
      Exit Sub
   End If
   
   If W.InDesign Then
      Lg.Url = URL_ACEPTACERT & PURL_LOGIN
   Else
      Lg.Url = URL_ACEPTA & PURL_LOGIN
   End If
      
   Lg.RutUsr = 17597643
   Lg.Pasw = "dani2105"
   Lg.RutEmpr = 96919050

   Rc = AcpAutenticar(Lg)
   
   If Rc = 0 Then
   
      Lg.PaswCert = "dani2116"
      Lg.MailEmisor = "pam@fairware.cl"
      
      DTE.xml = "<?xml version=""1.0""?>" & vbCrLf
      DTE.xml = DTE.xml & "<DTE xmlns=""http://www.sii.cl/SiiDte"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" version=""1.0"">"
'      Xml = Xml & "<Documento ID=""FT34"">"  ' Para Acepta no debe ir el ID=...
      DTE.xml = DTE.xml & "<Documento>"
      DTE.xml = DTE.xml & "<Encabezado><IdDoc><TipoDTE>34</TipoDTE><Folio/><FchEmis>2012-03-30</FchEmis></IdDoc>"
      ' Para pruebas el emisor debe ser Acepta
      DTE.xml = DTE.xml & "<Emisor><RUTEmisor>96919050-8</RUTEmisor><RznSoc>Acepta S.A.</RznSoc><GiroEmis>Computaci&amp;oacute;n</GiroEmis><Acteco>123456</Acteco><DirOrigen>Manquehue Sur 520, of. 428</DirOrigen><CmnaOrigen>Las Condes</CmnaOrigen><CiudadOrigen>Santiago</CiudadOrigen></Emisor>"
      DTE.xml = DTE.xml & "<Receptor><RUTRecep>76005641-3</RUTRecep><RznSocRecep>INMOBILIARIA PASEO LAS CONDES S.A.</RznSocRecep><GiroRecep>Inmobiliaria</GiroRecep><Contacto/><DirRecep>Avda. 11 de Septiembre Nº 1860, Of. 61</DirRecep><CmnaRecep>Providencia</CmnaRecep><CiudadRecep>Santiago</CiudadRecep></Receptor>"
      DTE.xml = DTE.xml & "<Totales><MntNeto>0</MntNeto><MntExe>9126071</MntExe><MntTotal>9126071</MntTotal></Totales></Encabezado><Detalle><NroLinDet>1</NroLinDet><CdgItem><TpoCodigo/><VlrCodigo/></CdgItem><IndExe>1</IndExe><NmbItem>PPTO. 1282</NmbItem><DscItem>PROYECTO CIRCULO CABOT I</DscItem><QtyItem>1</QtyItem><PrcItem>9126071</PrcItem><MontoItem>9126071</MontoItem></Detalle>"
'      Dte.Xml = Dte.Xml & "<TmstFirma>0</TmstFirma>"  ' Tag no reconocido por el SII
      DTE.xml = DTE.xml & "</Documento></DTE>"
      
      DTE.MailRecep = "franca@fairware.cl"
      
      Rc = AcpFirmarDte(Lg, DTE)
   End If
   
   End
End Sub

Public Function AcpProcesar(ByVal IdDTE As Long, ByVal XmlDTE As String, ByVal MailReceptor As String, ByVal ObsDoc As String, Folio As Long, UrlDoc As String)
   Dim Lg As AcpLogin_t, Rc As Long, AcpDTE As AcpDTE_t
   
   Call AddDebug("AcpProcesar: Inicio: IdDTE=" & IdDTE & ", Folio=" & Folio & ", Rc=" & Rc)
   
   Rc = AcpAutenticar(Lg)
      
   Call AddDebug("AcpProcesar: Despues de Autenticar - Folio=" & Folio & ", Rc=" & Rc)
   Call AddDebug("AcpProcesar: XmlDTE =[" & XmlDTE & "]")
      
   If Rc = 0 Then
      
      AcpDTE.xml = XmlDTE
      AcpDTE.IdDTE = IdDTE
      
      AcpDTE.MailRecep = MailReceptor
      AcpDTE.ObsDTE = ObsDoc
      AcpDTE.Folio = 0 ' 9 feb 2021: por siaca
      
      Rc = AcpFirmarDte(Lg, AcpDTE)
      
      If Rc = 0 And AcpDTE.Folio > 0 Then
         Folio = AcpDTE.Folio
         UrlDoc = AcpDTE.UrlDoc
      
      Else
         Call AddLog("AcpProcesar: FirmarDTE=" & Rc & ", idDTE=" & IdDTE)
         Folio = 0
         UrlDoc = ""

      End If
      
   End If
   
   AcpProcesar = Rc
   
End Function

Public Function AcpAutenticar(Login As AcpLogin_t) As Long
   Dim Req As String, Resp As String, Glosa As String, Error As Long, Mensaje As String
   Dim Md5 As ClsMd5
                  
   Call AddDebug("AcpAutenticar: Inicio: Usr:[" & gConectData.Usuario & "]")
            
   Call AcpInit(Login)
   
   If Login.RutUsr = "" Then
      MsgBox1 "Falta ingresar el RUT del usuario.", vbExclamation
      AcpAutenticar = AERR_NOPARAM
      Call AddLog("AcpAutenticar: Falta ingresar el RUT del usuario de conexión con Acepta.")
      Exit Function
   End If
      
   Call AddDebug("AcpAutenticar: Despues de AcpInit")
   
   Set Md5 = New ClsMd5
   Login.md5Pasw = LCase(Md5.DigestStrToHexStr(Login.Pasw))
   Set Md5 = Nothing
   
   Call AddDebug("AcpAutenticar: Despues de ClsMd5")
      
'   Login.md5Pasw = Login.Pasw
   
   Req = ""
   Call MkJSon(Req, "usuario", Login.RutUsr)     'Ambiente Certificación 17.597.643-4
   Call MkJSon(Req, "clave", Login.Pasw)      'Ambiente Certificación dani2105
'   Call MkJSon(Req, "rut", Login.RutUsr)     'Ambiente Certificación 17.597.643-4
'   Call MkJSon(Req, "pass", Login.Pasw)      'Ambiente Certificación dani2105
'   Call MkJSon(Req, "cod_md5", Login.md5Pasw)
'   Call MkJSon(Req, "empresa", Login.RutEmpr)
'   Call MkJSon(Req, "tipo_tx", "sesion_externa")
'   Call MkJSon(Req, "__FLAG_AUTENTIFICACION__", "SI", , True)
         
   Call AddDebug("AcpAutenticar: Req=[" & Req & "]")
   
   'Resp = FwPostPage(Login.Url, Req, "text/html")
   Call AddLog("AcpAutenticar: URL: " & Login.Url & "params: " & Req)
   Resp = FwPostPage(Login.Url, Req, "application/json")
   
   If Len(Resp) < 5 Then
      AcpAutenticar = AERR_NORESP
      Call AddLog("AcpAutenticar: no hay respuesta desde el servicio de autenticación.")
      Exit Function
   End If
   
   Debug.Print "Resp: [" & Resp & "]"
   Call AddDebug("AcpAutenticar: Resp=[" & Resp & "]")
   

   Error = Val(Replace(JSonValue(Resp, "status"), Chr(34), ""))
   Glosa = JSonValue(Resp, "message")
   
 
   If Error = AERR_OK Then ' OK
      Login.SessionId = Trim(JSonValue(Resp, "session_id"))
      
      If Login.SessionId = "" Then
         Call AddLog("AcpAutenticar: Resp=[" & Resp & "]")
         MsgBox1 "Error en la autenticación del usuario." & vbCrLf & Glosa & " (A)", vbExclamation
         AcpAutenticar = AERR_BADUSER
         Exit Function
      End If
      ' 17 dic 2020: Se analiza la respuesta para ver si tiene saldo
      Mensaje = GetMensaje(JSonValue(Resp, "url_default"))
      Call AddLog("AcpAutenticar: Mensaje=[" & Mensaje & "]")
      
      If InStr(1, Mensaje, " no tiene saldo disponible ", vbTextCompare) > 0 Then
         Call AddLog("AcpAutenticar: Sin Saldo Resp=[" & Resp & "]")
      End If

      AcpAutenticar = 0
      'Login.UrlRecibidos = Login.UrlRec & "session_id=" & Login.SessionId & "%26rutCliente=" & Login.RutEmpr

   Else
      If Error = 0 Then ' para que el error no sea cero
         Error = AERR_ERR1
      End If
      
      Call AddLog("AcpAutenticar: Resp=[" & Resp & "]")
      
      MsgBox1 "Error " & Error & ", " & Glosa & " (T).", vbExclamation
      AcpAutenticar = Error
   End If
           
End Function
Public Function AcpFirmarDte(Login As AcpLogin_t, DTE As AcpDTE_t) As Long
   Dim Req As String, Resp As String, Glosa As String, Error As Long
   Dim DatosAdjuntos As String
   
   If Len(DTE.xml) < 20 Or Login.SessionId = "" Then
      Call AddLog("AcpFirmarXml: largo XML < 20 o SessionId vacío.")
      AcpFirmarDte = AERR_NOPARAM
      Exit Function
   End If
   
   DTE.Folio = -1
   DTE.UrlDoc = ""
   
   Req = ""
'   Call MkJSon(Req, "pass", Login.PaswCert)        'Ambiente Certificación dani2116
'   Call MkJSon(Req, "session_id", Login.SessionId)
'   Call MkJSon(Req, "tipo_tx", "firmar_dte_xml")
'   Call MkJSon(Req, "aplicacion", "DTE")
'   Call MkJSon(Req, "rutCliente", Login.RutEmpr)
'   Call MkJSon(Req, "ID_UNICO_TX", DTE.IdDTE)   ' 31 dic 2020: para identificarlo en caso que falle la firma
''   Call MkJSon(Req, "XML_DTE_FIRMA", Cb64.EncodeTxt(XmlDTE))

   Call MkJSon(Req, "session_id", Replace(Login.SessionId, Chr(34), ""))      'Ambiente Certificación dani2116
   Call MkJSon(Req, "rut_signer", gEmpresa.RutFirma)
   Call MkJSon(Req, "encoding_xml", "UTF-8")
   Call MkJSon(Req, "xml_b64", DTE.xml)
   Call MkJSon(Req, "FOLIO_EXTERNO", "")
   Call MkJSon(Req, "ID_PROXY", DTE.IdDTE)   ' 31 dic 2020: para identificarlo en caso que falle la firma

   
'   DatosAdjuntos = GenAdjunto(Login, DTE)
'   If DatosAdjuntos = "" Then
'      Call AddLog("AcpFirmarXml: Mail del Receptor inválido: '" & DTE.MailRecep & "'")
'      MsgBox1 "Mail del Receptor inválido:" & vbCrLf & vbCrLf & DTE.MailRecep, vbExclamation
'      Exit Function
'   End If
   
'   Call MkJSon(Req, "XML_DTE", Str2Hex(DTE.xml & DatosAdjuntos))
'   Call MkJSon(Req, "__FLAG_RESPUESTA_NO_HTTP__", "SI", , True)
      
   Call AddDebug("AcpFirmarXml: Req=[" & Req & "]")
   
   If W.InDesign Then
      Login.Url = URL_TECNOBACKCERT & "v1/emitir/xml"
   Else
      Login.Url = URL_TECNOBACK & "v1/emitir/xml"
   End If
   'Resp = FwPostPage(Login.Url, Req, "text/html")
   Resp = FwPostPage(Login.Url, Req, "application/json")
   
   Debug.Print "Resp: [" & Resp & "]"
   
   If Len(Resp) < 5 Then
      AcpFirmarDte = AERR_NORESP
      Call AddLog("AcpFirmarXml: no hay respuesta desde el servicio de firma.")
      Exit Function
   End If
   
   Error = Val(Replace(JSonValue(Resp, "status"), Chr(34), ""))
   Glosa = JSonValue(Resp, "message")
   
   If Error = AERR_AUT_OK Then ' OK
      AcpFirmarDte = 0
      
      DTE.Folio = Val(Replace(JSonValue(Resp, "Folio"), Chr(34), ""))
      
      DTE.UrlDoc = Trim(Replace(ReplaceStr(JSonValue(Resp, "url_pdf"), "\/", "/"), Chr(34), ""))

   ElseIf Error = AERR_SESIONINVALIDA Then

      Call AddLog("AcpFirmarXml: Reg=[" & AcpRemovePass(Req) & "]")
      Call AddLog("AcpFirmarXml: Resp=[" & Resp & "]")
      
      MsgBox1 "Error " & Error & ", " & Glosa & " (T)." & vbCrLf & vbCrLf & "Verifique que nadie esté conectado en el Escritorio de ACEPTA.", vbExclamation
      
      AcpFirmarDte = Error

   Else
      If Error = 0 Then ' para que el error no sea cero
         Error = AERR_ERR3
      End If
      
      Call AddLog("AcpFirmarXml: Reg=[" & AcpRemovePass(Req) & "]")
      Call AddLog("AcpFirmarXml: Resp=[" & Resp & "]")
      
      If Error = AERR_FIRM_ERR And InStr(1, Glosa, "llamar al motor", vbTextCompare) > 0 Then ' Error al llamar al motor
         Error = AERR_MOTOR
      End If
      
      MsgBox1 "Error " & Error & ", " & Glosa & " (T).", vbExclamation
      
      AcpFirmarDte = Error
   End If

End Function

Public Function AcpPrevisualizar(ByVal XmlDTE As String, ByVal MailReceptor As String, ByVal ObsDoc As String, PreHtmlDTE As String, ByVal bExport As Boolean)
   Dim Lg As AcpLogin_t, Rc As Long, AcpDTE As AcpDTE_t
   
   Call AddDebug("AcpPrevisualizar: Inicio   Rc=" & Rc)
   
   Rc = AcpAutenticar(Lg)
      
   Call AddDebug("AcpPrevisualizar: Despues de Autenticar  Rc=" & Rc)
   Call AddDebug("AcpPrevisualizar: XmlDTE =[" & XmlDTE & "]")
      
   If Rc = 0 Then
      
      AcpDTE.xml = XmlDTE
      AcpDTE.bExport = bExport
      
      AcpDTE.MailRecep = MailReceptor
      AcpDTE.ObsDTE = ObsDoc
      If W.InDesign Then
        Lg.Url = URL_TECNOBACKCERT & "v1/preview-emitir/xml"
      Else
        Lg.Url = URL_TECNOBACK & "v1/preview-emitir/xml"
      End If
      Rc = AcpGetHTMLDte(Lg, AcpDTE, PreHtmlDTE)
            
   End If
   
   AcpPrevisualizar = Rc
   
End Function

Public Function AcpEventos(trazaEvento As AcpTrazaEvento_t, response As String)
   Dim Lg As AcpLogin_t, Rc As Long, AcpDTE As AcpDTE_t
   
   Call AddDebug("AcpEventos: Inicio   Rc=" & Rc)
   
   Rc = AcpAutenticar(Lg)
      
   Call AddDebug("AcpEventos: Despues de Autenticar  Rc=" & Rc)
   'Call AddDebug("AcpEventos: XmlDTE =[" & XmlDTE & "]")
      
   If Rc = 0 Then
    
    If W.InDesign Then
      Lg.Url = URL_TECNOBACKCERT & "v1/get-logger"
    Else
      Lg.Url = URL_TECNOBACK & "v1/get-logger"
    End If
      Rc = AcpGetEventosDte(Lg, trazaEvento, response)
            
   End If
   
   AcpEventos = Rc
   
End Function
Public Function AcpGetHTMLDte(Login As AcpLogin_t, DTE As AcpDTE_t, PreHtmlDTE As String) As Long
   Dim Req As String, Resp As String, Glosa As String, Error As Long
   Dim DatosAdjuntos As String, Resp2 As String, HtmlHex As String
   
   AcpGetHTMLDte = 0
   
   If Len(DTE.xml) < 20 Or Login.SessionId = "" Then
      Call AddLog("AcpFirmarXml: largo XML < 20 o SessionId vacío.")
      AcpGetHTMLDte = AERR_NOPARAM
      Exit Function
   End If
   
   DTE.Folio = -1
   DTE.UrlDoc = ""
   PreHtmlDTE = ""
   
   Req = ""
'   Call MkJSon(Req, "pass", Login.PaswCert)        'Ambiente Certificación dani2116
'   Call MkJSon(Req, "session_id", Login.SessionId)
'   Call MkJSon(Req, "tipo_tx", "firmar_dte_xml")
'   Call MkJSon(Req, "aplicacion", "DTE")
'   Call MkJSon(Req, "rutCliente", Login.RutEmpr)
''   Call MkJSon(Req, "XML_DTE_FIRMA", Cb64.EncodeTxt(XmlDTE))

   Call MkJSon(Req, "session_id", Replace(Login.SessionId, Chr(34), ""))      'Ambiente Certificación dani2116
   Call MkJSon(Req, "rut_signer", gEmpresa.RutFirma)
   Call MkJSon(Req, "encoding_xml", "UTF-8")
   Call MkJSon(Req, "xml_b64", DTE.xml)
   Call MkJSon(Req, "FOLIO_EXTERNO", "")
   Call MkJSon(Req, "ID_PROXY", DTE.IdDTE)   ' 31 dic 2020: para identificarlo en caso que falle la firma


   
'   DatosAdjuntos = GenAdjunto(Login, DTE)
'   If DatosAdjuntos = "" And DTE.bExport = False Then ' 2 oct 2020: se agrega  And DTE.bExport = False
'      Call AddLog("AcpFirmarXml: Mail del Receptor inválido: '" & DTE.MailRecep & "'")
'      MsgBox1 "Mail del Receptor inválido:" & vbCrLf & vbCrLf & DTE.MailRecep, vbExclamation
'      Exit Function
'   End If
'
'   Call MkJSon(Req, "XML_DTE", Str2Hex(DTE.xml & DatosAdjuntos))
''   Call MkJSon(Req, "__FLAG_RESPUESTA_NO_HTTP__", "SI", , True)
'   Call MkJSon(Req, "FLAG_PRE_VISUALIZACION", "SI", , True)
      
   Call AddDebug("AcpFirmarXml: Req=[" & Req & "]")
   
   
   'Resp = FwPostPage(Login.Url, Req, "text/html")
   Resp = FwPostPage(Login.Url, Req, "application/json")
   Debug.Print "Resp: [" & Resp & "]"
   
   If Len(Resp) < 5 Then
      AcpGetHTMLDte = AERR_NORESP
      Call AddLog("AcpFirmarXml: no hay respuesta desde el servicio de firma.")
      Exit Function
   End If
   
   Error = Val(Replace(JSonValue(Resp, "status"), Chr(34), ""))
   Glosa = JSonValue(Resp, "message")
   
   If Error = AERR_AUT_OK Then ' OK
      AcpGetHTMLDte = 0
      
'      Resp2 = JSonValue(Resp, "RESPUESTA")    'no le pone la llave (}) final
'      HtmlHex = JSonValue(Resp2 & "}", "HTML")
'      PreHtmlDTE = Hex2Str(HtmlHex)
      PreHtmlDTE = JSonValue(Resp, "url_pdf")
      
   ElseIf Error = AERR_SESIONINVALIDA Then

      Call AddLog("AcpFirmarXml: Reg=[" & AcpRemovePass(Req) & "]")
      Call AddLog("AcpFirmarXml: Resp=[" & Resp & "]")
      
      MsgBox1 "Error " & Error & ", " & Glosa & " (A)." & vbCrLf & vbCrLf & "Verifique que nadie esté conectado en el Escritorio de ACEPTA.", vbExclamation
      
      AcpGetHTMLDte = Error

   Else
      If Error = 0 Then ' para que el error no sea cero
         Error = AERR_ERR3
      End If

      Call AddLog("AcpFirmarXml: Reg=[" & AcpRemovePass(Req) & "]")
      Call AddLog("AcpFirmarXml: Resp=[" & Resp & "]")
      
      MsgBox1 "Error " & Error & ", " & Glosa & " (A).", vbExclamation
      
      AcpGetHTMLDte = Error
   End If

End Function
Public Function AcpGetEventosDte(Login As AcpLogin_t, trazaEvento As AcpTrazaEvento_t, response As String) As Long
   Dim Req As String, Resp As String, Glosa As String, Error As Long
   Dim DatosAdjuntos As String, Resp2 As String, HtmlHex As String
   
   AcpGetEventosDte = 0
   
   If Login.SessionId = "" Then
      Call AddLog("AcpGetEventosDte:  SessionId vacío.")
      AcpGetEventosDte = AERR_NOPARAM
      Exit Function
   End If
   
   
   Req = ""
'   Call MkJSon(Req, "pass", Login.PaswCert)        'Ambiente Certificación dani2116
'   Call MkJSon(Req, "session_id", Login.SessionId)
'   Call MkJSon(Req, "tipo_tx", "firmar_dte_xml")
'   Call MkJSon(Req, "aplicacion", "DTE")
'   Call MkJSon(Req, "rutCliente", Login.RutEmpr)
''   Call MkJSon(Req, "XML_DTE_FIRMA", Cb64.EncodeTxt(XmlDTE))

   Call MkJSon(Req, "RUTEmisor", trazaEvento.RutEmisor)
   Call MkJSon(Req, "TipoDTE", trazaEvento.TipoDTE)
   Call MkJSon(Req, "Folio", trazaEvento.Folio)
   Call MkJSon(Req, "canal", trazaEvento.canal)
   Call MkJSon(Req, "RUTRecep", trazaEvento.RUTRecep)
   Call MkJSon(Req, "session_id", Replace(Login.SessionId, Chr(34), ""))   ' 31 dic 2020: para identificarlo en caso que falle la firma

   
'   DatosAdjuntos = GenAdjunto(Login, DTE)
'   If DatosAdjuntos = "" And DTE.bExport = False Then ' 2 oct 2020: se agrega  And DTE.bExport = False
'      Call AddLog("AcpGetEventosDte: Mail del Receptor inválido: '" & DTE.MailRecep & "'")
'      MsgBox1 "Mail del Receptor inválido:" & vbCrLf & vbCrLf & DTE.MailRecep, vbExclamation
'      Exit Function
'   End If
'
'   Call MkJSon(Req, "XML_DTE", Str2Hex(DTE.xml & DatosAdjuntos))
''   Call MkJSon(Req, "__FLAG_RESPUESTA_NO_HTTP__", "SI", , True)
'   Call MkJSon(Req, "FLAG_PRE_VISUALIZACION", "SI", , True)
      
   Call AddDebug("AcpGetEventosDte: Req=[" & Req & "]")
   
   'Login.Url = URL_TECNOBACKCERT & "v1/preview-emitir/xml"
   'Resp = FwPostPage(Login.Url, Req, "text/html")
   Resp = FwPostPage(Login.Url, Req, "application/json")
   Debug.Print "Resp: [" & Resp & "]"
   
   If Len(Resp) < 5 Then
      AcpGetEventosDte = AERR_NORESP
      Call AddLog("AcpGetEventosDte: no hay respuesta desde el servicio de firma.")
      Exit Function
   End If
   
   Error = Val(Replace(JSonValue(Resp, "status"), Chr(34), ""))
   Glosa = JSonValue(Resp, "message")
   
   If Error = AERR_AUT_OK Then ' OK
      AcpGetEventosDte = 0
      response = Resp
'      Resp2 = JSonValue(Resp, "RESPUESTA")    'no le pone la llave (}) final
'      HtmlHex = JSonValue(Resp2 & "}", "HTML")
'      PreHtmlDTE = Hex2Str(HtmlHex)
'      PreHtmlDTE = JSonValue(Resp, "url_pdf")
      
   ElseIf Error = AERR_SESIONINVALIDA Then

      Call AddLog("AcpGetEventosDte: Reg=[" & AcpRemovePass(Req) & "]")
      Call AddLog("AcpGetEventosDte: Resp=[" & Resp & "]")
      
      MsgBox1 "Error " & Error & ", " & Glosa & " (T)." & vbCrLf & vbCrLf & "Verifique que nadie esté conectado en el Escritorio de ACEPTA.", vbExclamation
      
      AcpGetEventosDte = Error

   Else
      If Error = 0 Then ' para que el error no sea cero
         Error = AERR_ERR3
      End If

      Call AddLog("AcpGetEventosDte: Reg=[" & AcpRemovePass(Req) & "]")
      Call AddLog("AcpGetEventosDte: Resp=[" & Resp & "]")
      
      MsgBox1 "Error " & Error & ", " & Glosa & " (T).", vbExclamation
      
      AcpGetEventosDte = Error
   End If

End Function

Public Sub MkJSon(JSON As String, ByVal Item As String, ByVal Value As String, Optional ByVal bNum As Boolean = 0, Optional ByVal bClose As Boolean = 0)

   JSON = JSON & ",""" & Item & """:"
   
   If bNum Then
      JSON = JSON & Value
   Else
      JSON = JSON & """" & Value & """"
   End If
   
   If bClose Then
      JSON = "{" & Mid(JSON, 2) & "}"
   End If
   
End Sub

Private Function JSonValue(ByVal JSON As String, ByVal Item As String) As String
   Dim Aux As String, i As Long, j As Long
   
   Aux = """" & Item & """:"

   i = InStr(JSON, Aux)

   If i > 0 Then
      j = InStr(i + Len(Aux), JSON, ",")
      If j <= 0 Then
         j = InStr(i + Len(Aux), JSON, "}")
      End If

      If j > 0 Then
         Aux = Mid(JSON, i + Len(Aux), j - i - Len(Aux))
         
         If Left(Aux, 1) = """" And Right(Aux, 1) = """" Then
            Aux = Mid(Aux, 2, Len(Aux) - 2)
         End If
         
         JSonValue = Aux
         
      End If
   End If

End Function

Private Function GenAdjunto(Login As AcpLogin_t, DTE As AcpDTE_t) As String
   Dim xml As String, Dato As String

   DTE.MailRecep = Trim(DTE.MailRecep)

   If ValidEmail(DTE.MailRecep) = False And DTE.bExport = False Then   ' 2 oct 2020: Si es Exportación, no hay mail del Receptor
      Exit Function
   End If

   If DTE.MailRecep <> "" Then
      xml = "<DatoAdjunto nombre=""MailReceptor"">" & DTE.MailRecep & "</DatoAdjunto>"
      
      Dato = AddTag("NombreDA", "Mail_Receptor")
      Dato = Dato & AddTag("ValorDA", DTE.MailRecep)
      xml = xml & AddTag("DatosAdjuntos", Dato)
      
      Dato = AddTag("NombreDA", "Subject_Mail")
      Dato = Dato & AddTag("ValorDA", IIf(Login.SubjMail = "", "Envio Documento Electronico", Login.SubjMail))
      xml = xml & AddTag("DatosAdjuntos", Dato)
      
      If ValidEmail(Login.MailEmisor) Then
         Dato = AddTag("NombreDA", "Mail_Emisor")
         Dato = Dato & AddTag("ValorDA", Login.MailEmisor)
         xml = xml & AddTag("DatosAdjuntos", Dato)
      End If
   End If
   
   If DTE.ObsDTE <> "" Then
      Dato = AddTag("NombreDA", "Observacion")
      Dato = Dato & AddTag("ValorDA", Left(Ansi2XmlTxt(DTE.ObsDTE), 100))   'largo 100??
      xml = xml & AddTag("DatosAdjuntos", Dato)
   End If

   GenAdjunto = xml
   
End Function

Public Function AcpShowDTE(Frm As Form, ByVal UrlDTE As String, Optional ByVal Provisorio As Boolean = False) As Integer
   Dim Page As String
   Dim Buf As String
   Dim Rc As Integer
   
   AcpShowDTE = 0
   
   If UrlDTE = "" Then
      AcpShowDTE = -1
      Exit Function
   End If
   
   If Provisorio Then
      MsgBox1 "ATENCIÓN: " & vbCrLf & vbCrLf & "Esta es una visualización PROVISORIA del DTE, dado que aún no ha sido aceptado por el SII." & vbCrLf & vbCrLf & "El proceso que realiza el SII puede tomar varios minutos." & vbCrLf & vbCrLf & "Puede revisar el estado del DTE en la lista de DTE Emitidos.", vbInformation
   End If
   
   DoEvents
      
   Rc = ShellExecute(Frm.hWnd, "open", UrlDTE, "", "", 1)
   DoEvents
      
End Function

Public Function AcpShowEstadoDTE(ByVal IdDTE As Long, ByVal UrlDTE As String, idEstado As Integer, TxtEstado As String, trazaEvento As AcpTrazaEvento_t) As Integer
   Dim Page As String
   Dim Buf As String
   Dim Rc As Integer
   Dim Tr As AcpTraza_t
   Dim Frm As FrmDetTrazaDTE

   AcpShowEstadoDTE = 0
   
   If UrlDTE = "" Then
      AcpShowEstadoDTE = -1
      Exit Function
   End If
   
   'para ver estado DTE cambiar "v01" por "traza" en la URL del documento
'   Buf = ReplaceStr(UrlDTE, "v01", "traza")
   
   Tr.Url = UrlDTE
   'Tr.Url = URL_TECNOBACKCERT & "v1/get-estado-sii"
   If W.InDesign Then
    Tr.Url = URL_TECNOBACKCERT & "v1/get-logger"
   Else
    Tr.Url = URL_TECNOBACK & "v1/get-logger"
   End If
   
   If GetJTraza(Tr, trazaEvento) >= 0 Then
      
      Set Frm = New FrmDetTrazaDTE
      Call Frm.FView(IdDTE, Tr, idEstado, TxtEstado)
      Set Frm = Nothing
   Else
      MsgBox1 "Error al obtener información del documento.", vbInformation
   End If
   
'   DoEvents
      
'   Rc = ShellExecute(Frm.hWnd, "open", Buf, "", "", 1)
'   DoEvents
      
End Function


Public Function GetTraza_old1(Traza As AcpTraza_t) As Integer
   Dim Page As String, fld As String, Ini As Long, Fin As Long, p As Long, Tag As String, Fld1 As String
   Dim r As Integer, a As Integer, Tit As String, Nr As Integer, Det As DetTraza_t, nTot As Integer
   Dim Host As String, Path As String, Tits(5) As String, iTit As Integer, T As Integer, i As Integer
   
   ' https://
   r = InStr(12, Traza.Url, "/", vbBinaryCompare)
   Host = Left(Traza.Url, r - 1)
   
   Host = ReplaceStr(Host, "https", "http") ' No funciona con https
   
   Path = Mid(Traza.Url, r)
   
   Page = FwWebReadPage(Host, Path)

   Traza.nErr = 0
   Traza.nWarn = 0
   Traza.nAvisos = 0
   Traza.nAcepta = 0
   Traza.nSII = 0
   Traza.nIntercam = 0
   Traza.nControl = 0
   Traza.nMandat = 0
   Traza.nIECV = 0
   
   Tag = "<td class=""panelTitle"">"
   Ini = InStr(1, Page, Tag, vbTextCompare)
   Fin = InStr(Ini, Page, "</td>", vbTextCompare)
   p = Fin
   
   If Ini <= 0 Then
      Exit Function
   End If
   
   Ini = Ini + Len(Tag)
   fld = Mid(Page, Ini, Fin - Ini)

   Tag = "<span class=""document"" title=""Ver documento almacenado"">"
   Ini = InStr(1, Page, Tag, vbTextCompare)
   Fin = InStr(Ini, Page, "</span>", vbTextCompare)
   
   If Ini <= 0 Then
      Exit Function
   End If
   
   Ini = Ini + Len(Tag)
   fld = Mid(Page, Ini, Fin - Ini)
   
   Traza.Doc = fld
   
   ' Lista de Items a ana
   Tits(0) = "Avisos"
   Tits(1) = "ACEPTA"
   Tits(2) = "SII"
   Tits(3) = "Intercambio"
   Tits(4) = "Controller"
       
   ' Inicio de los datos
   For a = 1 To 50
      
      Tag = "<tr class=""ui-widget-header"">"
      Ini = InStr(p, Page, Tag, vbTextCompare)
      If Ini <= 0 Then
         Exit For
      End If
      
      Ini = Ini + Len(Tag)
      Fin = InStr(Ini, Page, "</tr>", vbTextCompare)
      fld = Mid(Page, Ini, Fin - Ini)
      fld = FwGetXmlTag(fld, "td", 1)
      fld = ReplaceStr(fld, vbLf, "")
      fld = Trim(ReplaceStr(fld, vbCr, ""))
      
      T = -1
      For i = 0 To UBound(Tits)
         
         If Len(Tits(i)) > 0 Then
            Ini = InStr(1, fld, Tits(i) & " (", vbTextCompare)
            If Ini > 0 Then
               T = i
               Exit For
            End If
         Else
            Exit For
         End If
      Next
      
      If T = -1 Then
         p = Fin
         GoTo Next_a
      End If
      
      Nr = Val(Mid(fld, Ini + Len(Tits(T) & " (")))
      p = Fin
      
      For r = 1 To Nr
      
         Tag = "<tr id=""j_idt58:" & a - 1 & ":j_idt65:" & r - 1 & "_row_" & r - 1 & """ class=""ui-widget-content"">"
      
         Ini = InStr(p, Page, Tag, vbTextCompare)
         If Ini <= 0 Then
            Exit For
         End If
         
         Ini = Ini + Len(Tag)
         Fin = InStr(Ini, Page, "</tr>", vbTextCompare)
         p = Fin
         fld = Mid(Page, Ini, Fin - Ini)
         
         ' Tipo
         Fld1 = FwGetXmlTag(fld, "td", 1)
                  
         If Left(Fld1, 2) = "<a" Then
            
            Tag = " href="""
            Ini = InStr(2, Fld1, Tag, vbTextCompare)
            
            If Ini > 0 Then
               Ini = Ini + Len(Tag)
            End If
            
            Fin = InStr(Ini, Fld1, """", vbTextCompare)
         
            Det.UrlTipo = Mid(Fld1, Ini, Fin - Ini)
               
            Fld1 = FwGetXmlTag(fld, "a", 1)
            
            If Left(Fld1, 5) = "<span" Then
               Fld1 = FwGetXmlTag(fld, "span", 1)
            End If
         End If
 
         Det.Tipo = Fld1
                       
         ' Icon
         Fld1 = FwGetXmlTag(fld, "td", 2)
         If InStr(1, Fld1, "info.png", vbTextCompare) Then           ' Azul - Info
            Det.Estado = ET_INFO
         ElseIf InStr(1, Fld1, "Level_1.png", vbTextCompare) Then    ' Verde
            Det.Estado = ET_OK
         ElseIf InStr(1, Fld1, "Level_2.png", vbTextCompare) Then    ' Amarillo
            Det.Estado = ET_WARN
            Traza.nWarn = Traza.nWarn + 1
         ElseIf InStr(1, Fld1, "Level_3.png", vbTextCompare) Then    ' Rojo
            Det.Estado = ET_ERR
            Traza.nErr = Traza.nErr + 1
         End If
         
         ' Fecha
         Fld1 = FwGetXmlTag(fld, "td", 3)
         Det.Fecha = GetDate(Fld1, "dmy", True)
         
         ' Obs
         Fld1 = FwGetXmlTag(fld, "td", 4)
         Det.Obs = Fld1
            
         Select Case T
            Case 0:
               Traza.Avisos(r - 1) = Det
               Traza.nAvisos = r
            Case 1:
               Traza.Acepta(r - 1) = Det
               Traza.nAcepta = r
            Case 2:
               Traza.SII(r - 1) = Det
               Traza.nSII = r
            Case 3:
               Traza.Intercam(r - 1) = Det
               Traza.nIntercam = r
            Case 4:
               Traza.Mandat(r - 1) = Det
               Traza.nMandat = r
            Case 5:
               Traza.Controll(r - 1) = Det
               Traza.nControl = r
            Case 6:
               Traza.IECV(r - 1) = Det
               Traza.nIECV = r
            Case Else:
               Exit For
         End Select

         nTot = nTot + 1
      Next r
      
Next_a:
   Next a

   GetTraza_old1 = nTot

End Function

' 29 nov 2017: por cambio en página de Acepta
Public Function GetJTraza(Traza As AcpTraza_t, trazaEvento As AcpTrazaEvento_t) As Integer
   Dim Page As String
   Dim Det As DetTraza_t
   Dim Host As String, Path As String, i As Integer
   Dim jTraza As Object
   Dim response As String
   
   If Len(Traza.Url) < 5 Then
      GetJTraza = -1
      Exit Function
   End If
   
   Traza.nErr = 0
   Traza.nWarn = 0
   Traza.nAvisos = 0
   Traza.nAcepta = 0
   Traza.nSII = 0
   Traza.nIntercam = 0
   Traza.nControl = 0
   Traza.nMandat = 0
   Traza.nIECV = 0
   Traza.Doc = ""
   Traza.Folio = 0
            
'   Path = "/ca4webv3?url=" & Traza.Url & "&accion=traza&FLAG_ESCRITORIO=SI"
'   Host = "http://motor-prod.acepta.com"
   
   Path = "/v1/get-logger"
   Host = "https://apicert.tecnoback.com"
   'Host = "https://api.tecnoback.com"
      
   'traza con errores de Findea para verlo con Pablo
   '  URL Traza: http://fairware1811.acepta.com/traza/00000000_1721614177_2215363636_40669354_?k=c6645c366c379d40edc2ec40cd58507a
   '  URL DTE: http://fairware1811.acepta.com/v01/00000000_1721614177_2215363636_40669354_?k=c6645c366c379d40edc2ec40cd58507a
   Page = AcpEventos(trazaEvento, response)
   'Page = FwWebReadPage(Host, Path)

   ' El Json se puede revisar en https://jsonformatter.curiousconcept.com/

'   If Len(Page) < 3 Then
'      GetJTraza = -1
'      Exit Function
'   End If
   Page = response

   'Set jTraza = json.parse(Page)
'   If jTraza Is Nothing Then
'      Call AddLog("GetJTraza: Page=" & Page)
'      GetJTraza = -2
'      Exit Function
'   End If
   
'   Debug.Print "[" & JSON.toString(jTraza) & "]"
'Dim a As Long
'Dim b As String
'   a = Val(Replace(JSonValue(Page, "folio"), Chr(34), ""))
'   b = JSonValue(Page, "tipo")
   Set jTraza = JSON.parse(Page)
   Dim jEventos, Ev, keys
   Dim key As String
   
   On Error Resume Next ' 26 ene 2021: se agrega verificación, porque si está vacío se cae al asignar
   Set Ev = jTraza.Item("logger")
   If Ev Is Nothing Then
      Call AddLog("GetJTraza: Header vacío [" & Page & "]")
'      MsgBox1 "No es posible obtener infomación del documento.", vbExclamation
      GetJTraza = -3
      Exit Function
   End If

   On Error GoTo 0

'   Traza.Doc = Ev.Item("nodelabel")
'   Traza.Folio = Val(Ev.Item("nodeid"))
'   Traza.Fecha = GetDate(Ev.Item("nodetimestamp"), "ymd", True)
   
   Traza.Doc = JSonValue(Page, "tipo")
   Traza.Folio = Val(Replace(JSonValue(Page, "folio"), Chr(34), ""))
   Traza.Fecha = GetDate(JSonValue(Page, "fecha"), "ymd", True)
   Traza.UrlData = Replace(JSonValue(Page, "url_temp_logger"), Chr(34), "")

   Set jEventos = jTraza.Item("logger")
   For i = 1 To jEventos.Count
      Set Ev = jEventos.Item(i)
      
      Debug.Print i & ":" & toString(Ev)
            
'      key = Ev.Item("icono")
'      If InStr(1, key, "info.png", vbTextCompare) Then           ' Azul - Info
'         Det.Estado = ET_INFO
'      ElseIf InStr(1, key, "Level_1.png", vbTextCompare) Then    ' Verde
'         Det.Estado = ET_OK
'      ElseIf InStr(1, key, "Level_2.png", vbTextCompare) Then    ' Amarillo
'         Det.Estado = ET_WARN
'         Traza.nWarn = Traza.nWarn + 1
'      ElseIf InStr(1, key, "Level_3.png", vbTextCompare) Then    ' Rojo
'         Det.Estado = ET_ERR
'         Traza.nErr = Traza.nErr + 1
'      ElseIf InStr(1, key, "primary", vbTextCompare) Then    ' Verde
'         Det.Estado = ET_OK
'      End If

      key = Ev.Item("icono")
      If InStr(1, key, "info.png", vbTextCompare) Then           ' Azul - Info
         Det.Estado = ET_INFO
      ElseIf InStr(1, key, "green", vbTextCompare) Then    ' Verde
         Det.Estado = ET_OK
      ElseIf InStr(1, key, "yellow", vbTextCompare) Then    ' Amarillo
         Det.Estado = ET_WARN
         Traza.nWarn = Traza.nWarn + 1
      ElseIf InStr(1, key, "red", vbTextCompare) Then    ' Rojo
         Det.Estado = ET_ERR
         Traza.nErr = Traza.nErr + 1
      ElseIf InStr(1, key, "primary", vbTextCompare) Then    ' Verde
         Det.Estado = ET_OK
      End If
      
'      Det.Fecha = GetDate(Ev.Item("processeddate"), "ymd", True)
'      Det.Tipo = Ev.Item("description")
'      Det.UrlTipo = IIf(IsNull(Ev.Item("url")), "", Ev.Item("url"))
'      Det.Obs = Ev.Item("comment2")

      Det.Fecha = GetDate(Ev.Item("fecha_insert"), "ymd", True)
      Det.Tipo = Ev.Item("evento")
      Det.UrlTipo = IIf(IsNull(Ev.Item("url_pdf")), "", Ev.Item("url_pdf"))
      Det.Obs = Ev.Item("message")
      
      key = Ev.Item("categoria")
      If StrComp(Left(key, 5), "OTROS", vbTextCompare) = 0 Then
         Traza.Avisos(Traza.nAvisos) = Det
         Traza.nAvisos = Traza.nAvisos + 1
      ElseIf StrComp(key, "Acepta", vbTextCompare) = 0 Then
         Traza.Acepta(Traza.nAcepta) = Det
         Traza.nAcepta = Traza.nAcepta + 1
      ElseIf StrComp(key, "SII", vbTextCompare) = 0 Then
         Traza.SII(Traza.nSII) = Det
         Traza.nSII = Traza.nSII + 1
      ElseIf StrComp(key, "INTERCAMBIO", vbTextCompare) = 0 Then
         Traza.Intercam(Traza.nIntercam) = Det
         Traza.nIntercam = Traza.nIntercam + 1
      ElseIf StrComp(key, "REGLAS", vbTextCompare) = 0 Then
         Traza.Controll(Traza.nControl) = Det
         Traza.nControl = Traza.nControl + 1
      ElseIf StrComp(key, "MANDATO", vbTextCompare) = 0 Then
         Traza.Mandat(Traza.nMandat) = Det
         Traza.nMandat = Traza.nMandat + 1
'      ElseIf StrComp(key, "OTROS", vbTextCompare) = 0 Then
'         Traza.IECV(Traza.nIECV) = Det
'         Traza.nIECV = Traza.nIECV + 1
      Else
         Debug.Print "GetJTraza: Grupo '" & key & "' no incluido."
      End If
   Next i
   
   GetJTraza = jEventos.Count

End Function

' 29 nov 2017: por cambio en página de Acepta
Public Function GetTraza_old2(Traza As AcpTraza_t) As Integer
   Dim Page As String, fld As String, Ini As Long, Fin As Long, p As Long, Tag As String, Fld1 As String
   Dim r As Integer, a As Integer, Tit As String, Nr As Integer, Det As DetTraza_t, nTot As Integer
   Dim Host As String, Path As String, Tits(5) As String, iTit As Integer, T As Integer, i As Integer
   Dim Eventos As String
   
   ' https://
   r = InStr(12, Traza.Url, "/", vbBinaryCompare)
   Host = Left(Traza.Url, r - 1)
   
   Host = ReplaceStr(Host, "https", "http") ' No funciona con https
   
   Path = Mid(Traza.Url, r)
   
   Page = FwWebReadPage(Host, Path)

   Traza.nErr = 0
   Traza.nWarn = 0
   Traza.nAvisos = 0
   Traza.nAcepta = 0
   Traza.nSII = 0
   Traza.nIntercam = 0
   Traza.nControl = 0
   Traza.nMandat = 0
   Traza.nIECV = 0
   
'   Tag = "<td class=""panelTitle"">"
'   Ini = InStr(1, Page, Tag, vbTextCompare)
'   Fin = InStr(Ini, Page, "</td>", vbTextCompare)
'   p = Fin
   
   Tag = "<table id=""example"""
   Ini = InStr(1, Page, Tag, vbTextCompare)
   Fin = InStr(Ini, Page, "</table>", vbTextCompare)
   p = Fin
   
   If Ini <= 0 Then
      Exit Function
   End If
   
   Eventos = Mid(Page, Ini, Fin + Len("</table>") - Ini)

   Ini = Ini + Len(Tag)
   fld = Mid(Page, Ini, Fin - Ini)

   Tag = "<div class=""panel-heading panel_principal"">"
   Ini = InStr(1, Page, Tag, vbTextCompare)
   
   If Ini <= 0 Then
      Exit Function
   End If
   Fin = InStr(Ini, Page, "</div>", vbTextCompare)
   
   Ini = Ini + Len(Tag)
   fld = Mid(Page, Ini, Fin - Ini)
   
   Traza.Doc = FwGetXmlTag(fld, "a")
      
   ' Lista de Items a ana
   Tits(0) = "Avisos"
   Tits(1) = "ACEPTA"
   Tits(2) = "SII"
   Tits(3) = "Intercambio"
   Tits(4) = "Controller"
       
   ' Inicio de los datos
   p = 1
   For a = 1 To 50
      
      Tag = "<tr class=""item_grupo"">"
      Ini = InStr(p, Eventos, Tag, vbTextCompare)
      If Ini <= 0 Then
         Exit For
      End If
      
      Ini = Ini + Len(Tag)
      Fin = InStr(Ini, Eventos, "</tr>", vbTextCompare)
      fld = Mid(Eventos, Ini, Fin - Ini)
      fld = FwGetXmlTag(fld, "td", 1)
      fld = ReplaceStr(fld, vbLf, "")
      fld = Trim(ReplaceStr(fld, vbCr, ""))
      
      T = InStr(fld, "<i ")
      If T > 0 Then
         fld = Trim(Left(fld, T - 1))
      End If
      
      T = -1
      For i = 0 To UBound(Tits)
         
         If Len(Tits(i)) > 0 Then
            Ini = InStr(1, fld, Tits(i) & " (", vbTextCompare)
            If Ini > 0 Then
               T = i
               Exit For
            End If
         Else
            Exit For
         End If
      Next
      
      If T = -1 Then
         p = Fin
         GoTo Next_a
      End If
      
      Nr = Val(Mid(fld, Ini + Len(Tits(T) & " (")))
      p = Fin
      
      For r = 1 To Nr
      
         fld = FwGetXmlTag(Eventos, "tr", r, p)
      
'         Tag = "<tr id=""j_idt58:" & a - 1 & ":j_idt65:" & r - 1 & "_row_" & r - 1 & """ class=""ui-widget-content"">"
      
'         Ini = InStr(p, Eventos, Tag, vbTextCompare)
'         If Ini <= 0 Then
'            Exit For
'         End If
'
'         Ini = Ini + Len(Tag)
'         Fin = InStr(Ini, Eventos, "</tr>", vbTextCompare)
'         p = Fin
'         Fld = Mid(Eventos, Ini, Fin - Ini)
         
         ' Tipo
         Fld1 = FwGetXmlTag(fld, "td", 1)
                  
         If Left(Fld1, 2) = "<a" Then
            
            Tag = " href="""
            Ini = InStr(2, Fld1, Tag, vbTextCompare)
            
            If Ini > 0 Then
               Ini = Ini + Len(Tag)
            End If
            
            Fin = InStr(Ini, Fld1, """", vbTextCompare)
         
            Det.UrlTipo = Mid(Fld1, Ini, Fin - Ini)
               
            Fld1 = FwGetXmlTag(fld, "a", 1)
            
            If Left(Fld1, 5) = "<span" Then
               Fld1 = FwGetXmlTag(fld, "span", 1)
            End If
         End If
 
         Det.Tipo = IIf(Left(Fld1, 1) = "<", FwGetXmlTag(Fld1, "strong", 1), Fld1)
                       
         ' Icon
         Fld1 = FwGetXmlTag(fld, "td", 2)
         If InStr(1, Fld1, "info_azul", vbTextCompare) Then           ' Azul - Info
            Det.Estado = ET_INFO
         ElseIf InStr(1, Fld1, "cir_verde", vbTextCompare) Then    ' Verde
            Det.Estado = ET_OK
         ElseIf InStr(1, Fld1, "amarillo", vbTextCompare) Then    ' Amarillo
            Det.Estado = ET_WARN
            Traza.nWarn = Traza.nWarn + 1
         ElseIf InStr(1, Fld1, "rojo", vbTextCompare) Then    ' Rojo
            Det.Estado = ET_ERR
            Traza.nErr = Traza.nErr + 1
         End If
         
         ' Fecha
         Fld1 = FwGetXmlTag(fld, "td", 3)
         Det.Fecha = GetDate(Fld1, "ymd", True)
         
         ' Obs
         Fld1 = FwGetXmlTag(fld, "td", 4)
         Det.Obs = Fld1
            
         Select Case T
            Case 0:
               Traza.Avisos(r - 1) = Det
               Traza.nAvisos = r
            Case 1:
               Traza.Acepta(r - 1) = Det
               Traza.nAcepta = r
            Case 2:
               Traza.SII(r - 1) = Det
               Traza.nSII = r
            Case 3:
               Traza.Intercam(r - 1) = Det
               Traza.nIntercam = r
            Case 4:
               Traza.Mandat(r - 1) = Det
               Traza.nMandat = r
            Case 5:
               Traza.Controll(r - 1) = Det
               Traza.nControl = r
            Case 6:
               Traza.IECV(r - 1) = Det
               Traza.nIECV = r
            Case Else:
               Exit For
         End Select

         nTot = nTot + 1
      Next r
      
Next_a:
   Next a

   GetTraza_old2 = nTot

End Function

Public Function StrEstadoTraza(ByVal Estado As Integer) As String

   Select Case Estado
   
      Case ET_INFO:
         StrEstadoTraza = "Información"
      Case ET_OK:
         StrEstadoTraza = "OK"
      Case ET_WARN:
         StrEstadoTraza = "Advertencia"
      Case ET_ERR:
         StrEstadoTraza = "Error"
         
   End Select
   
End Function

Private Sub AcpInit(Lg As AcpLogin_t)

   Lg.RutUsr = 0
   
   If W.InDesign Then
      Lg.Host = URL_ACEPTACERT
      Lg.Host = URL_TECNOBACKCERT
   Else
      Lg.Host = URL_ACEPTA
      'produccion
      'Lg.Host = URL_TECNOBACK
      'Desarrollo
      Lg.Host = URL_TECNOBACK
   End If

   If W.InDesign Or gEmpresa.Rut = RUT_EMP_ACEPTA Then
      
'       Lg.RutUsr = RUT_USR_ACEPTA
'       Lg.Pasw = CLAVE_USR_ACEPTA
       Lg.RutUsr = RUT_USR_TECNOBACKDESA
       Lg.Pasw = CLAVE_USR_TECNOBACKDESA

'      Lg.RutEmpr = RUT_EMP_ACEPTA
'      Lg.PaswCert = CLAVE_CERT_ACEPTA
'      Lg.MailEmisor = MAIL_EMISOR_ACEPTA
   Else
      
      If IsNumeric(gConectData.Usuario) = False Then
         Call AddLog("El Usuario de Conexión no es un RUT (" & gConectData.Usuario & ")")
         Exit Sub
      End If
       'Desarrollo
       Lg.RutUsr = RUT_USR_TECNOBACK
       Lg.Pasw = CLAVE_USR_TECNOBACK
         
'      Lg.RutUsr = Val(gConectData.Usuario)
'      Lg.Pasw = gConectData.Clave
'      Lg.RutEmpr = gEmpresa.Rut
'      Lg.PaswCert = gConectData.ClaveCert
'      Lg.MailEmisor = gConectData.MailEmisor
   End If
   
   'Lg.Url = Lg.Host & "/pivote.php"
   'Lg.UrlRec = Lg.Host & "/ext.php?r=" & Lg.Host & "/appDinamicaClasses/index.php%3Fapp_dinamica=buscarNEW_recibidos%26identificador=NO_BUSCAR%26"
   Lg.Url = Lg.Host & "v1/get-session-id"
   
   Call AddDebug("AcpInit: User=" & Lg.RutUsr)
   Call AddDebug("AcpInit: Pasw=" & Left(Lg.Pasw, 2) & "...")
   Call AddDebug("AcpInit: Rut=" & Lg.RutEmpr)

End Sub

Public Sub AcpShowRecibidos(Frm As Form)
   Dim Lg As AcpLogin_t

   If AcpAutenticar(Lg) = 0 Then
      
      Call ShellExecute(Frm.hWnd, "open", Lg.UrlRecibidos, "", "", 1)
      
   End If

End Sub
'obtiene lista de archivos de facturas del libro de compras
Public Function AcpListArchFCompra(ByVal Rut As Long, Files() As String, nFiles As Integer) As Long
   Dim Buf As String, Rc As Long, Path As String, k As Long, Fn As String, p As Long, F As Integer
   
   Buf = "[--" & Rut & "#" & CLng(Int(Now)) & "$L==>##"
   k = GenClave(Buf, 1311753)
   
   Path = "/DirAcepta_.asp?r=" & Rut & "&o=L&k=" & k & "&e=" & gEmpresa.Rut

   Buf = FwWebReadPage(URL_FAIR, Path)
   
   If Left(Buf, 7) <> "!Files=" Then
      AcpListArchFCompra = 11    'error
      Exit Function
   End If
   
   p = InStr(1, Buf, vbCrLf, vbBinaryCompare)
   If p <= 0 Then
      AcpListArchFCompra = 12    'error
      Exit Function
   End If
   
   p = p + 2
   
   F = 0
   ReDim Files(10)
   Do
      Fn = Trim(NextField2(Buf, p, vbCrLf))

      If Len(Fn) < 10 Then
         Exit Do
      End If

      If F > UBound(Files) Then
         ReDim Preserve Files(F + 5)
      End If

      Files(F) = Fn
      F = F + 1
   Loop

   nFiles = F
   AcpListArchFCompra = 0
   
End Function
'obtiene lista de archivos de facturas del libro de compras
Public Function AcpCountArchFCompra(ByVal Rut As Long) As Integer
   Dim Files() As String
   Dim nFiles As Integer, Rc As Long

   Rc = AcpListArchFCompra(Rut, Files(), nFiles)
   If Rc = 0 And nFiles > 0 Then
      AcpCountArchFCompra = nFiles
   End If

End Function

'Obtiene el contenido de un archivo a un buffer
Public Function AcpArchFCompra(ByVal Rut As Long, ByVal Fn As String, Buf As String) As Long
   Dim Rc As Long, Path As String, k As Long, p As Long, q As Long
   
   Buf = "[--" & Rut & "#" & CLng(Int(Now)) & "$c==>" & Fn & "##"
   k = GenClave(Buf, 1311753)

   Path = "/DirAcepta_.asp?r=" & Rut & "&o=c&f=" & Fn & "&k=" & k & "&e=" & gEmpresa.Rut

   Buf = FwWebReadPage(URL_FAIR, Path)
   
   If Left(Buf, 12) <> "!Contenido [" Then
      AcpArchFCompra = 11
      Exit Function
   End If
   
   p = InStr(1, Buf, vbCrLf, vbBinaryCompare)
   If p > 0 Then
      q = InStrRev(Buf, "!FinContenido!" & vbCrLf, -1, vbBinaryCompare)
      If q > 0 Then
         Buf = Mid(Buf, p + 2, q - p - 2) ' +2 por vbCrLf
      Else
         AcpArchFCompra = 12
      End If
   Else
      AcpArchFCompra = 13
   End If
   
End Function
'elimina el archivo especificado
Public Function AcpDelArchFCompra(ByVal Rut As Long, ByVal Fn As String) As Long ' , Buf As String) As Long
   Dim Buf As String, Rc As Long, Path As String, k As Long, p As Long, q As Long
   
   Buf = "[--" & Rut & "#" & CLng(Int(Now)) & "$c==>" & Fn & "##"
   k = GenClave(Buf, 1311753)

   Path = "/DirAcepta_.asp?r=" & Rut & "&o=k&f=" & Fn & "&k=" & k

   Buf = FwWebReadPage(URL_FAIR, Path)
   
   If Left(Buf, 9) <> "!Delete [" Then
      AcpDelArchFCompra = 11
      Exit Function
   End If
      
End Function

' Obtiene el mensaje que viene en hexa desde url_default
Private Function GetMensaje(ByVal Url As String) As String
   Dim i As Integer, j As Integer, HMsg As String
   
   i = InStr(Url, "&__mensaje=")
   If i > 0 Then
      i = i + 11
      j = InStr(i, Url, "&") ' por si viene algo después
      
      If j > 0 Then
         HMsg = Mid(Url, i, j - i)
      Else
         HMsg = Mid(Url, i)
      End If
      
      GetMensaje = Trim(Hex2Str(HMsg))
   End If

End Function

' 23 dic 2020: para que la clave no quede visible en el Log
Private Function AcpRemovePass(ByVal Req As String) As String
   Dim i As Long, j As Long, Tag As String, Pass As String

   Tag = """pass"":"""
   i = InStr(1, Req, Tag, vbTextCompare)
   
   If i <= 0 Then
      AcpRemovePass = Req
      Exit Function
   End If
   
   i = i + Len(Tag)
   j = InStr(i, Req, """", vbBinaryCompare)
   If i <= 0 Then
      AcpRemovePass = Req
      Exit Function
   End If

'   AcpRemovePass = Left(Req, i) & "..." & Mid(Req, j - 1)

   Pass = Mid(Req, i, j - i)
   Pass = FwEncrypt1(Pass, 98173)
   AcpRemovePass = Left(Req, i - 1) & Pass & Mid(Req, j)

End Function
