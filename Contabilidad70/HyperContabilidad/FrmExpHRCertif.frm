VERSION 5.00
Begin VB.Form FrmExpHRCertif 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar a HR-Certificados Honorarios (Declaración Jurada 1879)"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Exp 
      Caption         =   "Exportar"
      Height          =   315
      Left            =   5880
      TabIndex        =   7
      Top             =   480
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   5880
      TabIndex        =   6
      Top             =   840
      Width           =   1275
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Período"
      Height          =   975
      Left            =   1260
      TabIndex        =   1
      Top             =   420
      Width           =   4395
      Begin VB.TextBox Tx_Ano 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   420
         Width           =   855
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   5
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   4
         Top             =   480
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   300
      Picture         =   "FrmExpHRCertif.frx":0000
      ScaleHeight     =   570
      ScaleWidth      =   585
      TabIndex        =   0
      Top             =   480
      Width           =   585
   End
End
Attribute VB_Name = "FrmExpHRCertif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Const HR_RUTFMT = "00,000,000"

Const HR_HONORARIOS = 1
Const HR_PARTICIP10 = 2
Const HR_PARTICIP20 = 3

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_Exp_Click()
   Dim Mes As Integer
   Dim Msg As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TipoContrib As Integer
   
   Q1 = "SELECT TipoContrib FROM Empresa "
   Q1 = Q1 & " WHERE Id = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      TipoContrib = vFld(Rs("TipoContrib"))
   End If
   
   Call CloseRs(Rs)
   
   If TipoContrib <= 0 Then
      MsgBox1 "Antes de realizar la exportación debe definir el tipo de contribuyente." & vbCrLf & vbCrLf & " Utilice el botón Empresa disponible en la ventana principal del sistema.", vbExclamation
      Exit Sub
   End If

   Mes = ItemData(Cb_Mes)
   If Mes < 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   Msg = "¡ATENCION!" & vbLf & "Esta exportación reemplazará los valores actuales en el producto HR-Certificados"
   
   If Mes > 0 Then
      Msg = Msg & " para el mes de " & Cb_Mes & "."
   Else
      Msg = Msg & " para todos los meses del año."
   End If
   
   Msg = Msg & vbLf & vbLf & "Antes de realizar la exportación, asegúrese que ningún usuario tenga abierto el producto HR-Certificados con la empresa " & gEmpresa.RazonSocial & " para el año comercial " & gEmpresa.Ano & "."
   Msg = Msg & vbLf & vbLf & "¿ Desea continuar ?"
   If MsgBox1(Msg, vbExclamation Or vbYesNo Or vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If

   MousePointer = vbHourglass
   Bt_Exp.Enabled = False
   DoEvents
   
   
'   If ExportCertif(Mes) = 0 Then
'
'
'      Msg = "¡ATENCIÓN!" & vbLf & "Para terminar el proceso, ahora debe abrir el producto HR-Certificados y realizar la Numeración de Certificados de Honorarios."
'      Msg = Msg & vbLf & vbLf & "Empresa: " & gEmpresa.RazonSocial & vbLf & "Año Comercial: " & gEmpresa.Ano
'      MsgBox1 Msg, vbInformation
'
      Call Export_DJ1879(Mes, True)
      
'   Else
'      MsgBox1 "No se pudo realizar la exportación.", vbExclamation
'
'   End If
   
   Bt_Exp.Enabled = True
   MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()

   Call AddItem(Cb_Mes, "(todos)", 0)
   
   Call FillMes(Cb_Mes, GetMesActual())
   
   Cb_Mes.ListIndex = 0

   Tx_Ano = gEmpresa.Ano

End Sub

Public Function ExportCertif(ByVal Mes As Integer) As Integer
#If DATACON = 1 Then       'Access
   Dim DbName As String
   Dim Rc As Integer
   Dim TblCli As String
   Dim TblHonoPart As String
   Dim TblHonorario As String
   Dim CertPrefix As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Rut As String
   Dim Rs1 As Recordset
   Dim TipoContrib As Integer
   
   ExportCertif = 0
   
   If gLinkF22 = False Then
      MsgBox1 "No se encontraron los archivos correspondientes al producto HR-Certificados en " & vbLf & W.AppPath & "\..\PAR", vbExclamation
      ExportCertif = -1
      Exit Function
   End If
   
   Q1 = "SELECT TipoContrib FROM Empresa"
   Q1 = Q1 & " WHERE Id = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      TipoContrib = vFld(Rs("TipoContrib"))
   End If
   
   Call CloseRs(Rs)


   ' Linkeamos las tablas de Certificados

   DbName = gHRPath & "\RUTS\" & Right("000000000" & vFmtRut(gEmpresa.Rut), 8) & "\CerA" & gEmpresa.Ano + 1 & ".MDB"
   CertPrefix = "CER_"
   
   Rc = True
   
   TblCli = CertPrefix & "CLIENTE"
   Rc = Rc And LinkMdbTable(DbMain, DbName, Mid(TblCli, Len(CertPrefix) + 1), TblCli, , False)
   TblHonoPart = CertPrefix & "HONOPART_CERT"
   Rc = Rc And LinkMdbTable(DbMain, DbName, Mid(TblHonoPart, Len(CertPrefix) + 1), TblHonoPart, , False)
   TblHonorario = CertPrefix & "HONORARIO"
   Rc = Rc And LinkMdbTable(DbMain, DbName, Mid(TblHonorario, Len(CertPrefix) + 1), TblHonorario, , False)

   If Rc Then ' existen las tablas, ahora vemos si están los campos
   
      Q1 = "SELECT RUT, MES, TIPO, BRUTO, IMPTO, DATO_USUARIO FROM " & TblHonorario
      Set Rs = OpenRs(DbMain, Q1, False)
      If Rs Is Nothing Then
         Rc = False
      End If
   
      Call CloseRs(Rs)
   
   End If

   If Rc = False Then
      ExportCertif = -2
      If ERR = 3024 Or ERR = 3044 Or ERR = 3061 Then ' path, archivo o tablas no existen
         MsgBox1 "Aún no se ha creado este contribuyente para el año " & gEmpresa.Ano + 1 & " con el producto HR-Certificados." & vbLf & "Con el programa HR-Certificados " & gEmpresa.Ano + 1 & " abra el contribuyente con RUT " & FmtRut(vFmtCID(gEmpresa.Rut)) & "-" & DV_Rut(vFmtCID(gEmpresa.Rut)) & " y reintente nuevamente la exportación.", vbExclamation
      ElseIf ERR = 3045 Then
         MsgBox1 "Es probable que un usuario esté trabajando en HR-Certificados con esta empresa y año. " & vbNewLine & vbNewLine & "Verifique y vuelva a intentarlo.", vbExclamation + vbOKOnly
      End If
      Exit Function
   End If
   
   'eliminamos los registros que hay en la tabla de Honorarios de HR-Certificados
   Q1 = "DELETE * FROM " & TblHonorario
   If Mes > 0 Then
      Q1 = Q1 & " WHERE Mes = " & Mes
   End If
   
   Call ExecSQL(DbMain, Q1)

   'Insertamos los registros de retenciones
   Q1 = " SELECT Entidades.Rut, Entidades.Nombre, " & SqlMonthLng("FEmision") & " As Mes, Afecto, OtroImp, NumDoc, TipoRetencion, PorcentRetencion "
   Q1 = Q1 & " FROM (Documento INNER JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad) "
   'Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE " & SqlYearLng("FEmision") & " = " & gEmpresa.Ano
   If Mes > 0 Then
      Q1 = Q1 & " AND " & SqlMonthLng("FEmision") & " = " & Mes
   End If
   Q1 = Q1 & " AND Exento = 0 "
   
   If TipoContrib = CONTRIB_SAABIERTA Or TipoContrib = CONTRIB_SACERRADA Or TipoContrib = CONTRIB_SPORACCION Then
      Q1 = Q1 & " AND TipoRetencion IN (" & TR_HONORARIOS & "," & TR_DIETA & ")"
      Q1 = Q1 & " AND PorcentRetencion IN (" & IMPRET_NAC & "," & IMPRET_EXT & ")"
   
   Else  '1a cat y 2a cat
      Q1 = Q1 & " AND TipoRetencion IN (" & TR_HONORARIOS & ")"
      Q1 = Q1 & " AND PorcentRetencion IN (" & IMPRET_NAC & ")"
   
   End If
   
   Q1 = Q1 & " AND NotValidRut = 0"

   Set Rs = OpenRs(DbMain, Q1)

   Do While Not Rs.EOF
   
      Rut = Format(vFld(Rs("Rut")), HR_RUTFMT) & "-" & DV_Rut(vFld(Rs("Rut")))
   
      'insertamos el documento
      Q1 = "INSERT INTO " & TblHonorario
      Q1 = Q1 & " (Rut, Mes, Tipo, Bruto, Impto, Dato_Usuario)"
      Q1 = Q1 & " VALUES('" & Rut & "'"
      Q1 = Q1 & "," & vFld(Rs("Mes"))
      
      If vFld(Rs("TipoRetencion")) = TR_HONORARIOS Then
         Q1 = Q1 & ", " & HR_HONORARIOS
      ElseIf vFld(Rs("PorcentRetencion")) = IMPRET_NAC Then
         Q1 = Q1 & ", " & HR_PARTICIP10
      Else
         Q1 = Q1 & ", " & HR_PARTICIP20
      End If
      
      Q1 = Q1 & "," & vFld(Rs("Afecto"))
      Q1 = Q1 & "," & vFld(Rs("OtroImp"))
      Q1 = Q1 & "," & Val(vFld(Rs("NumDoc"))) & ")"
      
      Call ExecSQL(DbMain, Q1)
      
      'insertamos la entidad en la tabla Cliente, si ya existe el RUT, actualizamos el nombre
      Q1 = "SELECT Nombre FROM " & TblCli & " WHERE Rut = '" & Rut & "'"
      Set Rs1 = OpenRs(DbMain, Q1)
      
      If Rs1.EOF Then
         Q1 = "INSERT INTO " & TblCli
         Q1 = Q1 & "(Rut, Nombre) VALUES('" & Rut & "', '" & vFld(Rs("Nombre")) & "')"
      Else
         Q1 = "UPDATE " & TblCli
         Q1 = Q1 & " SET Nombre = '" & vFld(Rs("Nombre")) & "'"
         Q1 = Q1 & " WHERE Rut = '" & Rut & "'"
      End If
         
      Call ExecSQL(DbMain, Q1)
      Call CloseRs(Rs1)
      
      'insertamos la entidad en la tabla HonoPart_Cert, sólo si no existe el RUT
      Q1 = "SELECT Numero FROM " & TblHonoPart & " WHERE Rut = '" & Rut & "'"
      Set Rs1 = OpenRs(DbMain, Q1)
      
      If Rs1.EOF Then
         Q1 = "INSERT INTO " & TblHonoPart
         Q1 = Q1 & "(Rut) VALUES('" & Rut & "')"
         Call ExecSQL(DbMain, Q1)
      End If
         
      Call CloseRs(Rs1)
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Call UnLinkTable(DbMain, TblCli)
   Call UnLinkTable(DbMain, TblHonoPart)
   Call UnLinkTable(DbMain, TblHonorario)
   
#End If

End Function

Public Function Export_DJ1879(ByVal Mes As Integer, Optional ByVal Msg As Boolean = True) As Long
   Dim Rc As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Rut As String
   Dim TipoContrib As Integer
   Dim Fd As Long
   Dim SFName As String, fname As String
   Dim Sep As String
   Dim Valor As Double
   Dim r As Integer
   Dim Buf As String
   Dim TipoRet As Integer
   Dim FPath As String
   
   Export_DJ1879 = 0
   
   Sep = ";"
   
   On Error Resume Next
      
'   FPath = glHRPathExportPath & "\HRDJ\"
   FPath = gHRPath & "\RUTS"
   MkDir FPath
'   FPath = gExportPath & "\HRDJ\" & gEmpresa.Rut
   FPath = FPath & "\" & Right("00000000" & vFmtCID(FmtCID(gEmpresa.Rut)), 8)    'se hace por la empresa que tiene los ruts con puntos y guión
   
   MkDir FPath

   FPath = FPath & "\ImpConta"
   MkDir FPath

   SFName = "DJ1879_" & Right("00" & Mes, 2) & Right(gEmpresa.Ano, 2) & ".csv"
   
   fname = FPath & "\" & SFName

   Fd = FreeFile
   ERR.Clear
   
   Open fname For Output As #Fd
   If ERR Then
      MsgErr fname
      Export_DJ1879 = -ERR
      Exit Function
   End If

   On Error GoTo 0
   
   Q1 = "SELECT TipoContrib FROM Empresa"
   Q1 = Q1 & " WHERE Id = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      TipoContrib = vFld(Rs("TipoContrib"))
   End If
   
   Call CloseRs(Rs)
   
   'Insertamos los registros de retenciones
   Q1 = " SELECT Entidades.Rut, Entidades.Nombre, " & SqlMonthLng("FEmision") & " As Mes, Afecto, OtroImp, NumDoc, TipoRetencion, PorcentRetencion,ValRet3Porc, Sucursales.Codigo As CodSuc "
   Q1 = Q1 & " FROM (Documento INNER JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " AND Documento.IdEmpresa = Entidades.IdEmpresa )"
   'Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc "
   Q1 = Q1 & " LEFT JOIN Sucursales ON Documento.IdSucursal = Sucursales.IdSucursal AND Documento.IdEmpresa = Sucursales.IdEmpresa "
   Q1 = Q1 & " WHERE TipoLib = " & LIB_RETEN & " AND " & SqlYearLng("FEmision") & " = " & gEmpresa.Ano
   If Mes > 0 Then
      Q1 = Q1 & " AND " & SqlMonthLng("FEmision") & " = " & Mes
   End If
   Q1 = Q1 & " AND Exento = 0 "
   
   If TipoContrib = CONTRIB_SAABIERTA Or TipoContrib = CONTRIB_SACERRADA Or TipoContrib = CONTRIB_SPORACCION Then
      Q1 = Q1 & " AND TipoRetencion IN (" & TR_HONORARIOS & "," & TR_DIETA & ")"
      
      '3056944
      'Q1 = Q1 & " AND PorcentRetencion IN (" & IMPRET_NAC & "," & IMPRET_EXT & ")"
      Q1 = Q1 & " AND PorcentRetencion IN (" & IMPRET_NAC & "," & IMPRET_EXT & "," & IMPRET_OTRO & ")"
      '3056944
      
   Else  '1a primera cat y 2a catolica
      Q1 = Q1 & " AND TipoRetencion IN (" & TR_HONORARIOS & ")"
      Q1 = Q1 & " AND PorcentRetencion IN (" & IMPRET_NAC & ")"
   
   End If
   
   Q1 = Q1 & " AND NotValidRut = 0"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano

   Set Rs = OpenRs(DbMain, Q1)
   
'   Buf = "RUT" & Sep & "Razon Social" & Sep & "Mes" & Sep & "Tipo" & Sep & "Bruto" & Sep & "Impto" & Sep & "Dato_Usuario"
   Buf = "Mes" & Sep & "RUT" & Sep & "Razon Social" & Sep & "Bruto" & Sep & "Impto" & Sep & "Tipo" & Sep & "Dato_Usuario" & Sep & "NumCert" & Sep & "Sucursal"

'   Print #Fd, Buf

   Buf = ""
   r = 0
   
   'imprimimos el archivo
   Do While Rs.EOF = False
   
' 2736180
If gEmpresa.Ano >= 2021 Then
   '3056944
    If vFld(Rs("PorcentRetencion")) = IMPRET_OTRO Then
     
     If Val(Rs("Afecto") * 0.1) = vFld(Rs("OtroImp")) Then
     
      'agregamos el documento
      Buf = vFld(Rs("Mes")) & Sep
      
      Rut = vFld(Rs("Rut")) & "-" & DV_Rut(vFld(Rs("Rut")))
      
      Buf = Buf & Rut & Sep & vFld(Rs("Nombre")) & Sep
      
      Buf = Buf & vFld(Rs("Afecto")) & Sep
      Buf = Buf & vFld(Rs("OtroImp")) & Sep
      
      If vFld(Rs("ValRet3Porc")) > 0 Then
        Buf = Buf & "SI" & Sep
        Else
        Buf = Buf & "NO" & Sep
      End If
      
      Buf = Buf & vFld(Rs("ValRet3Porc")) & Sep    'retencion3porc
     
      If vFld(Rs("TipoRetencion")) = TR_HONORARIOS Then
         TipoRet = HR_HONORARIOS
      ElseIf vFld(Rs("PorcentRetencion")) = IMPRET_NAC Then
         TipoRet = HR_PARTICIP10
      ElseIf vFld(Rs("PorcentRetencion")) = IMPRET_OTRO Then
         TipoRet = HR_PARTICIP10
      Else
         TipoRet = HR_PARTICIP20
      End If
      
      Buf = Buf & TipoRet & Sep
      
      Buf = Buf & Val(vFld(Rs("NumDoc"))) & Sep
      
     ' Buf = Buf & Sep    'NumCert Vacío
      Buf = Buf & vFld(Rs("CodSuc"))     'Sucursal
      
      'Buf = Buf & Sep    'NumCert Vacío
      'Buf = Buf & Sep
      'Buf = Buf & vFld(Rs("ValRet3Porc"))     'retencion3porc
      
      
      Print #Fd, Buf
      r = r + 1
 
     
     End If
   
    Else
    '3056944
    
      'agregamos el documento
      Buf = vFld(Rs("Mes")) & Sep
      
      Rut = vFld(Rs("Rut")) & "-" & DV_Rut(vFld(Rs("Rut")))
      
      Buf = Buf & Rut & Sep & vFld(Rs("Nombre")) & Sep
      
      Buf = Buf & vFld(Rs("Afecto")) & Sep
      Buf = Buf & vFld(Rs("OtroImp")) & Sep
      
      '2736180
      
      If vFld(Rs("ValRet3Porc")) > 0 Then
        Buf = Buf & "SI" & Sep
        Else
        Buf = Buf & "NO" & Sep
      End If
      
      Buf = Buf & vFld(Rs("ValRet3Porc")) & Sep    'retencion3porc
     
     'fin 2736180
     
      If vFld(Rs("TipoRetencion")) = TR_HONORARIOS Then
         TipoRet = HR_HONORARIOS
      ElseIf vFld(Rs("PorcentRetencion")) = IMPRET_NAC Then
         TipoRet = HR_PARTICIP10
      Else
         TipoRet = HR_PARTICIP20
      End If
      
      Buf = Buf & TipoRet & Sep
      
      Buf = Buf & Val(vFld(Rs("NumDoc"))) & Sep
      
     ' Buf = Buf & Sep    'NumCert Vacío
      Buf = Buf & vFld(Rs("CodSuc"))     'Sucursal
      
      'Buf = Buf & Sep    'NumCert Vacío
      'Buf = Buf & Sep
      'Buf = Buf & vFld(Rs("ValRet3Porc"))     'retencion3porc
      
      
      Print #Fd, Buf
      r = r + 1
      
      End If
      '3056944
            
      Rs.MoveNext
      
Else

      'agregamos el documento
      Buf = vFld(Rs("Mes")) & Sep
      
      Rut = vFld(Rs("Rut")) & "-" & DV_Rut(vFld(Rs("Rut")))
      
      Buf = Buf & Rut & Sep & vFld(Rs("Nombre")) & Sep
      
      Buf = Buf & vFld(Rs("Afecto")) & Sep
      Buf = Buf & vFld(Rs("OtroImp")) & Sep
     
      If vFld(Rs("TipoRetencion")) = TR_HONORARIOS Then
         TipoRet = HR_HONORARIOS
      ElseIf vFld(Rs("PorcentRetencion")) = IMPRET_NAC Then
         TipoRet = HR_PARTICIP10
      Else
         TipoRet = HR_PARTICIP20
      End If
      
      Buf = Buf & TipoRet & Sep
      
      Buf = Buf & Val(vFld(Rs("NumDoc"))) & Sep
      
     ' Buf = Buf & Sep    'NumCert Vacío
      Buf = Buf & vFld(Rs("CodSuc"))     'Sucursal
      
      'Buf = Buf & Sep    'NumCert Vacío
      'Buf = Buf & Sep
      Buf = Buf & vFld(Rs("ValRet3Porc"))     'retencion3porc
      
      
      Print #Fd, Buf
      r = r + 1
            
      Rs.MoveNext

End If
      
   Loop
   
   Call CloseRs(Rs)
   
   Close Fd

   If Msg Then
      If r = 0 Then
         MsgBox1 "No existen datos para generar esta Declaración Jurada.", vbInformation
         Export_DJ1879 = 0
      Else
         fname = ReplaceStr(fname, "C:\HR\LPContab\..\", "C:\HR\")
         MsgBox1 "Proceso de exportación finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & fname, vbInformation + vbOKOnly
         
         If gEmpresa.Ano < 2019 Then
            MsgBox1 "Recuerde que debe tomar el archivo csv y llevarlo al HR Importador de Certificados", vbInformation
         Else
            Call ConectHRCertif("1879", SFName, "")
         End If
         

      End If
      Export_DJ1879 = r

   End If
   
End Function


