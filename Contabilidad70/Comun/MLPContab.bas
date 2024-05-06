Attribute VB_Name = "MLPContab"
Option Explicit

'******************* NO MODIFICAR **********************

'2868088
'Password DB
Public Const SG_SEGCFG = "FW6T9R54WX3A"  'archivo cfg para eliminar clave
'Public Const SG_SEGCFG_NEW = "FW6T9R64WX4A"  'archivo cfg para eliminar clave
'fin 2868088

'2868088
'password LexContab.mdb
Public Const PASSW_PREFIX = "Fw#42+"   'prefijo password empresa (sigue RUT sin puntos, ni guión, ni dígito verificador
Public Const PASSW_PREFIX_NEW = "Fw#54+"   'prefijo password empresa (sigue RUT sin puntos, ni guión, ni dígito verificador
'fin 2868088


'2868088
Public Const PASSW_LEXCONT = "Fw#420!&+"
Public Const PASSW_LEXCONT_NEW = "Fw#529!&+"
Public Const PASSW_LEXCONT_NEW2 = "Fw#540!&+"
'fin 2868088

'2850275
Public Const SG_PASSW_FAIRPAY = "oP,*/'#2j7h7_$3"
'fin 2850275

#If DATACON = 1 Then


'Public Const FAIRCONT_CODE = 78239315    ' Hasta la versión 3 - 11 sep 2012
Public Const FAIRCONT_CODE = 3179512      ' Version 4 con IFRS
Public Const APP_NAME = "LpContab4"

' Para generar el código de red
' Public Const PC_SEED = 765432        ' Hasta la versión 3 - 11 sep 2012
Public Const PC_SEED = 391719          ' Versión 4 con IFRS

#Else

Public Const FAIRCONT_CODE = 5717895      ' Version para SQL Server
Public Const APP_NAME = "LpContabSql"
Public Const PC_SEED = 319175          ' Versión SQL Server jul 2019

#End If


'Public Const APP_URL = "http://www.fairware.cl/LpContab2.asp"
Public Const APP_URL = "https://www.hyperrenta.cl/?page_id=4058"
Public Const APP_FULLNAME = "LP Contabilidad"

Public Const APP_DEMO = False

'******************* NO MODIFICAR **********************

' Informacion para el archivo de licencias
Public gLicFile As String
Public Const KEY_CRYP = 7827141
Public gCantLicencias As Integer  ' cantidad de licencias autorizadas, se llena en ChkInscPC

' Versiones de Mantenciones
'Public Const VER_2005 = 1
'Public Const VER_2005M = 500  ' sólo por el 2005
'Public Const VER_DEMO = VER_2005M   ' La última disponible para los usuarios en DEMO
'Public Const VER_2006 = 600
'Public Const VER_2006M = 650
'Public Const VER_2007 = 700
'Public Const VER_2007M = 750
'Public Const VER_2008 = 800
'Public Const VMANT_ALL = 99999999#
'Public Const VMANT_2005 = VER_2005M ' compatibilidad

' Pam: 13 dic 2010: Licencias
Public Const VER_ILIM = 800   ' Como la actual
Public Const VER_5EMP = 700   ' 5 empresas
Public Const VER_DEMO = 600   ' 3 empresas

#If DATACON = 2 Then
Public Const VER_50EMP = 705     ' máximo 50 empresas
Public Const VER_100EMP = 710    ' máximo 100 empresas
Public Const VER_200EMP = 720    ' máximo 200 empresas
Public Const VER_400EMP = 740    ' máximo 400 empresas
Public Const VER_800EMP = 780    ' máximo 800 empresas
#End If

Public gMaxEmpLicencia As Integer

'cantidad máxima de comprobantes para la versión Demo
Public Const MAX_COMPDEMO = 20

'cantidad máxima de documentos para la versión Demo
Public Const MAX_DOCDEMO = 50

' *********************************************************************



Public Sub InitLexComun()
   Dim i As Integer
      
   'en versión 1.0.15 del 28 Oct. 2005 se libera exportación a FUT
   'con fecha 14 sept. 2017 se elimina exportación a FUT hasta nuevo aviso, por indicación de Cristofer Elgueta
   gFunciones.ExpFUT = False
   
   'en versión 1.0.17 del 4 Ene 2006 se libera exportación a Certif
   gFunciones.ExpHRCertificados = True
   
   ' en versión 1.0.21 del 6 de abr 2005, se libera export a F22
   gFunciones.ExpHRForm22 = True
   
   'funciones nuevas para año 2006
   gFunciones.ActivoFijo = True              'Entregado
   gFunciones.RazFinancieras = True          'Entregado
   gFunciones.OtrosInformes = True           'Entregado
   gFunciones.DetDocReten = True             'entregado
   gFunciones.DetSaldoApertura = True        'Entregado
   gFunciones.ComprobanteResumido = True     'Entregado
   gFunciones.ExpImpLibrosAux = True         'entregado
   gFunciones.ExpImpLibrosAuxFile = True     'entregado
   gFunciones.ExpPlanCuentas = True          'Entregado
   
   gFunciones.PrtCheque = True               'Entregado
   gFunciones.ImportRemu = True              'Entregado
   
   gFunciones.IFRS = True                    'Entregado
   gFunciones.IFRS_BalanceTributario = True  'Entregado
   gFunciones.IFRS_Ejecutivo = True          'Entregado
   
   gFunciones.NuevoTraspasoIVA = True       'entregado
   
   gFunciones.NuevoTraspasoForm22 = True    'entregado
      
   gFunciones.ImportComprobantes = True      'Entregado
   gFunciones.ImportRetenciones = True       'Entregado
   gFunciones.AuditoriaInterna = True        'Entregado
   gFunciones.ControlContrib = True          'Entregado
   
   gFunciones.ExpLibCompVentasSII = True     'Entregado
   
   gFunciones.ProporcionalidadIVA = True    'Entregado
   
   gFunciones.ActFijoFinanciero = True       'entregado
   gFunciones.RepActFijoFinanciero = True    'entregado
   
   gFunciones.LibroCaja = True               'entregado 28 ene 2016
   
   gFunciones.DocCuotas = True               'desarrollo desde 6 jul 2017, entregado a testing 1 ago 2017
   gFunciones.OtrosIngEgresos = True         'desarrollo desde 1 ago 2017
   
   gFunciones.AjustesExtraLibCaja = True    'desarrollo desde 1 dic 2017
   
   'gLexContab = "L" & "e" & "x" & "is" & "Ne" & "x" & "is" & " C" & "on" & "ta" & "bi" & "l" & "id" & "ad"
   gLexContab = "L" & "e" & "g" & "al" & "Pu" & "b" & "li" & "s" & "hi" & "ng" & " C" & "on" & "ta" & "bi" & "l" & "id" & "ad"

   'App.HelpFile = W.AppPath & "\LexContabilidad.hlp"
   App.HelpFile = W.AppPath & "\LPContabilidad.hlp"
   
   'gIniFile = "LexContab.ini"
   On Error Resume Next
   MkDir ("C:\TReuters")
   gIniFile = "C:\TReuters\LPContab.ini"
   If Not ExistFile(gIniFile) Then
      Call CopyOldIniFile("LPContab.Ini")
   End If
   On Error GoTo 0
   
   'gCfgFile = W.AppPath & "\LexContab.cfg"
   gCfgFile = W.AppPath & "\LPContab.cfg"
   gAdmUser = "administ"
   gValidRut = True

   gAppCode.Code = FAIRCONT_CODE
   gAppCode.Name = APP_NAME
   gAppCode.Title = App.Title
   gAppCode.TVerif = 1 ' LexContab
'   gAppCode.emailSop = "soporte@legalpublishing.cl"     'cambio a solicitud de Carlo Maturana 11/10/16
'   gAppCode.emailInfo = "soporte@legalpublishing.cl"
   gAppCode.emailSop = "soporte.chile@thomsonreuters.com"
   gAppCode.emailInfo = "soporte.chile@thomsonreuters.com"
   
   gAppCode.Contacto = "LegalPublishing"
   gAppCode.TxInsc1 = "Gracias por probar nuestro producto. Si usted desea adquirirlo, por favor contáctese con Legal Publishing a los teléfonos (56-2) 510 5100, (56) 600 700 8000."
   gAppCode.TxInsc2 = "Para obtener el Código de Usuario: utilice el botón [Solicitud de Codigo de Usuario] o utilice el botón [Copiar datos] y luego péguelos en un email dirigido a " & gAppCode.emailSop & "."
   gAppCode.IniFile = gIniFile
   gAppCode.CfgFile = gCfgFile
   
   
   gAppCode.NivDef = VER_5EMP ' el más limitado
   
   ' pam: 13 dic 2010
   i = 0
   gAppCode.Nivel(i).id = VER_ILIM
   gAppCode.Nivel(i).Desc = "Sin límite de empresas"
   
   i = i + 1
   gAppCode.Nivel(i).id = VER_5EMP
   gAppCode.Nivel(i).Desc = "Máximo cinco empresas"
         
      
#If DATACON = 2 Then

   i = i + 1
   gAppCode.Nivel(i).id = VER_50EMP
   gAppCode.Nivel(i).Desc = "Máximo 50 empresas"

   i = i + 1
   gAppCode.Nivel(i).id = VER_100EMP
   gAppCode.Nivel(i).Desc = "Máximo 100 empresas"

   i = i + 1
   gAppCode.Nivel(i).id = VER_200EMP
   gAppCode.Nivel(i).Desc = "Máximo 200 empresas"

   i = i + 1
   gAppCode.Nivel(i).id = VER_400EMP
   gAppCode.Nivel(i).Desc = "Máximo 400 empresas"

   i = i + 1
   gAppCode.Nivel(i).id = VER_800EMP
   gAppCode.Nivel(i).Desc = "Máximo 800 empresas"

#End If
      
   i = i + 1
   gAppCode.Nivel(i).id = 0
   gAppCode.Nivel(i).Desc = ""   ' fin de la lista
   
   
   'Call GetExtInfo("html", gHtmExt)
   Call GetExtInfo(".html", gHtmExt) 'PS le agregué ., porque no reconocia el iexplorer.exe     FCA 1/09/2021
   
   Call FindPrinter(GetIniString(gIniFile, "Config", "Printer"), True)
   
   On Error Resume Next
   
End Sub

