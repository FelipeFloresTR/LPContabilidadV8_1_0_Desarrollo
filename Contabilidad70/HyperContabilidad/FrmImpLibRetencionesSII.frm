VERSION 5.00
Begin VB.Form FrmImpLibRetencionesSII 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración Honorarios SII"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   480
      TabIndex        =   18
      Top             =   2760
      Width           =   7215
      Begin VB.OptionButton Op_Automatico 
         Caption         =   "Automatico"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Op_Manual 
         Caption         =   "Manual"
         Height          =   195
         Left            =   600
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos SII"
      Height          =   735
      Left            =   600
      TabIndex        =   13
      Top             =   1560
      Width           =   7455
      Begin VB.TextBox Txt_ClaveSII 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3840
         PasswordChar    =   "x"
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox Txt_UsuarioSII 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Clave:"
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Rut: "
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.CommandButton Bt_InfoAyuda 
      Caption         =   "Consideraciones..."
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Bt_Manual 
      Caption         =   "Manual Uso"
      Height          =   795
      Left            =   6840
      Picture         =   "FrmImpLibRetencionesSII.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ver Manual Libro Electrónico de Compras"
      Top             =   480
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   540
      Picture         =   "FrmImpLibRetencionesSII.frx":06B6
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   600
      Width           =   555
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   8460
      TabIndex        =   5
      Top             =   780
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Importar"
      Height          =   315
      Left            =   8460
      TabIndex        =   4
      Top             =   420
      Width           =   1275
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Período"
      Height          =   975
      Left            =   2160
      TabIndex        =   7
      Top             =   360
      Width           =   4395
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox Tx_Ano 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   10
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   9
         Top             =   480
         Width           =   330
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Archivo"
      Height          =   1095
      Left            =   1080
      TabIndex        =   6
      Top             =   3360
      Width           =   9315
      Begin VB.TextBox Tx_FName 
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   420
         Width           =   7455
      End
      Begin VB.CommandButton Bt_Browse 
         Height          =   495
         Left            =   7800
         Picture         =   "FrmImpLibRetencionesSII.frx":0C43
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmImpLibRetencionesSII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFName As String

Dim lCtaBruto As Cuenta_t
Dim lCtaHonSinRet As Cuenta_t

Dim valida As String

Private Sub Bt_Browse_Click()

   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = ""
   FrmMain.Cm_ComDlg.InitDir = gImportPath
   FrmMain.Cm_ComDlg.Filter = "Archivos CSV (*.csv)|*.csv"
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
   If ERR = cdlCancel Then
      Exit Sub
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   ERR.Clear
   
   lFName = FrmMain.Cm_ComDlg.Filename
   
   Tx_FName = lFName
   
   DoEvents
      
End Sub

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Import_Click()
   Dim Rc As Integer
   Dim Q1 As String
   Dim Info() As String

   If DateSerial(Val(Tx_Ano), CbItemData(Cb_Mes), 1) < DateSerial(2017, 8, 1) Then
      MsgBox1 "El período a importar debe ser igual o superior al comienzo del Registro de Compras y Ventas (2017-08)", vbExclamation
      Exit Sub
   End If
   
   If GetEstadoMes(CbItemData(Cb_Mes)) <> EM_ABIERTO Then
      MsgBox1 "El mes seleccionado no está abierto.", vbExclamation
      Exit Sub
   End If
'   If Me.Op_Manual Then
'        If lFName = "" Then
'           MsgBox1 "Debe seleccionar el archivo.", vbExclamation + vbOKOnly
'           Exit Sub
'        End If
'
'        Q1 = "UPDATE ParamEmpresa "
'        Q1 = Q1 & " SET Valor = '1'"
'        Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND Tipo = 'DATOSII' AND Codigo = 1 "
'
'        Call ExecSQL(DbMain, Q1)
'
'   Else
        
        If Trim(Me.Txt_UsuarioSII.Text) = "" Then
            MsgBox1 "Favor ingresar el rut de Ingreso en el SII.", vbExclamation + vbOKOnly
            Me.Txt_UsuarioSII.SetFocus
            Exit Sub
        Else
             If Not ValidRut(Me.Txt_UsuarioSII.Text) Then
               MsgBox1 "Rut No Válido, Favor volver a ingresar", vbInformation
               Me.Txt_UsuarioSII.Text = ""
               Exit Sub
             Else
                If vFmtCID(Me.Txt_UsuarioSII.Text) <> gEmpresa.Rut Then
                    MsgBox1 "Rut No Coincide con el de la empresa en uso", vbInformation
                    Me.Txt_UsuarioSII.Text = ""
                    Exit Sub
                End If
            End If
        End If
        
        If Me.Txt_ClaveSII.Text = "" Then
            MsgBox1 "Favor ingresar la Clave de Ingreso en el SII.", vbExclamation + vbOKOnly
            Me.Txt_ClaveSII.SetFocus
            Exit Sub
        End If
        
        Q1 = "UPDATE ParamEmpresa "
        Q1 = Q1 & " SET Valor = '2'"
        Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND Tipo = 'DATOSII' AND Codigo = 1 "

        Call ExecSQL(DbMain, Q1)
        
        Q1 = "UPDATE ParamEmpresa "
        Q1 = Q1 & " SET Valor = '" & vFmtCID(Me.Txt_UsuarioSII.Text) & "'"
        Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND Tipo = 'DATOSII' AND Codigo = 2 "

        Call ExecSQL(DbMain, Q1)
        
        Q1 = "UPDATE ParamEmpresa "
        Q1 = Q1 & " SET Valor = '" & Trim(Me.Txt_ClaveSII.Text) & "'"
        Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND Tipo = 'DATOSII' AND Codigo = 3 "

        Call ExecSQL(DbMain, Q1)
        
'   End If
    If LoadDefCuentasRet(LIBRETEN_BRUTO, LIB_RETEN) = 0 Or LoadDefCuentasRet(LIBRETEN_HONORSINRET, LIB_RETEN) = 0 Then
       MsgBox1 "Favor ingresar las Cuentas del Libro de Retenciones en la Configuración de Cuentas Básicas.", vbInformation
       Exit Sub
    End If

   Rc = MsgBox1("Atención:" & vbNewLine & vbNewLine & "Se importará en el sistema mediante Integración con el SII…" & vbNewLine & vbNewLine & lFName & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
   If Rc = vbNo Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
'   If Me.Op_Manual Then
'        Call Import_LibroComprasSII(Me, lFName, Val(Tx_Ano), CbItemData(Cb_Mes))
'   Else
    'Call Import_LibroRetencionesSIIAuto(Me, lFName, Val(Tx_Ano), CbItemData(Cb_Mes), Me.Txt_UsuarioSII.Text, Me.Txt_ClaveSII.Text)
    valida = True
    Info = InformacionSii(Val(Tx_Ano), CbItemData(Cb_Mes), Me.Txt_UsuarioSII.Text, Me.Txt_ClaveSII.Text)
    If valida Then
        Call Import_LibroRetencionesSIIAuto(Val(Tx_Ano), CbItemData(Cb_Mes), lCtaHonSinRet, lCtaBruto, Me, Info)
    End If
   '   End If
   Me.MousePointer = vbDefault
      
End Sub

Private Sub Bt_InfoAyuda_Click()
MsgBox1 "Estimado " & vbNewLine & "El sistema LP Conta captura solo hasta 3500 Registros ya sea para Compras o Ventas " & vbNewLine & "Si desea capturar una cantidad mayor deber utilizar la Versión LP Conta SQL", vbExclamation
End Sub

Private Sub Bt_Manual_Click()
   Dim Rc As Long
   Dim Buf As String
   
   MousePointer = vbHourglass
   DoEvents
   
   Buf = gAppPath & "\Manual_Registo_Compras_SII.pdf"
   Rc = ExistFile(Buf)
      
   If Rc = 0 Then
      MsgBox1 "No se encontró el archivo que contiene el Manual de Importación del Registro de Compras desde el SII, por favor contáctese con su proveedor para obtenerlo.", vbExclamation
   Else

      Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
      If Rc < 32 Then
         MsgBox1 "Error " & Rc & " al abrir el archivo '" & Buf & "' que contiene el Importación del Registro de Compras desde el SII." & vbLf & "Trate de abrir este archivo con otro programa.", vbExclamation
      End If
   End If

   MousePointer = vbDefault

End Sub


Private Sub Form_Load()
   Dim MesActual As Integer
   Dim Q1 As String
   Dim Rs As Recordset

   MesActual = GetMesActual()
   
   Call FillMes(Cb_Mes)
   If MesActual > 0 Then
      Cb_Mes.ListIndex = MesActual - 1
   Else
      Cb_Mes.ListIndex = GetUltimoMesConMovs() - 1
   End If
   Me.Op_Manual.Value = True
   Tx_Ano = gEmpresa.Ano
   
   
   If gDbType = SQL_SERVER Then
        Q1 = "IF NOT EXISTS(Select * From ParamEmpresa WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND Tipo = 'DATOSII' AND Codigo = 2)"
        Q1 = Q1 & " BEGIN"
        Q1 = Q1 & " INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) VALUES (" & gEmpresa.id & "," & gEmpresa.Ano & ",'DATOSII',1,'1')"
        Q1 = Q1 & " INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) VALUES (" & gEmpresa.id & "," & gEmpresa.Ano & ",'DATOSII',2,'0')"
        Q1 = Q1 & " INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) VALUES (" & gEmpresa.id & "," & gEmpresa.Ano & ",'DATOSII',3,'0')"
        Q1 = Q1 & " END"
        Call ExecSQL(DbMain, Q1)
   Else
        Q1 = "INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) "
        Q1 = Q1 & " SELECT TOP 1 " & gEmpresa.id & " AS IdEmpresa, " & gEmpresa.Ano & " AS Ano, 'DATOSII' AS Tipo, 1 AS Codigo,1 AS Valor"
        Q1 = Q1 & " FROM ParamEmpresa"
        Q1 = Q1 & " WHERE NOT EXISTS (SELECT TOP 1 IdEmpresa, Ano, Tipo, Codigo, Valor FROM ParamEmpresa WHERE Tipo = 'DATOSII' AND Codigo = 1);"
        Call ExecSQL(DbMain, Q1)
        
        Q1 = "INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) "
        Q1 = Q1 & " SELECT TOP 1 " & gEmpresa.id & " AS IdEmpresa, " & gEmpresa.Ano & " AS Ano, 'DATOSII' AS Tipo, 2 AS Codigo,0 AS Valor"
        Q1 = Q1 & " FROM ParamEmpresa"
        Q1 = Q1 & " WHERE NOT EXISTS (SELECT TOP 1 IdEmpresa, Ano, Tipo, Codigo, Valor FROM ParamEmpresa WHERE Tipo = 'DATOSII' AND Codigo = 2);"
        Call ExecSQL(DbMain, Q1)
        
        Q1 = "INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) "
        Q1 = Q1 & " SELECT TOP 1 " & gEmpresa.id & " AS IdEmpresa, " & gEmpresa.Ano & " AS Ano, 'DATOSII' AS Tipo, 3 AS Codigo,0 AS Valor"
        Q1 = Q1 & " FROM ParamEmpresa"
        Q1 = Q1 & " WHERE NOT EXISTS (SELECT TOP 1 IdEmpresa, Ano, Tipo, Codigo, Valor FROM ParamEmpresa WHERE Tipo = 'DATOSII' AND Codigo = 3);"
        Call ExecSQL(DbMain, Q1)
   End If
   
   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa "
   Q1 = Q1 & " WHERE Idempresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND Tipo = 'DATOSII'"
   Q1 = Q1 & " Order By Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
        Select Case vFld(Rs("Codigo"))
                Case 1
                     If vFld(Rs("Valor")) = 1 Then
                        Me.Op_Manual.Value = True
                     Else
                        Me.Op_Automatico = True
                     End If
                Case 2
                     Me.Txt_UsuarioSII.Text = IIf(vFld(Rs("Valor")) = 0, "", FmtCID(Rs("Valor")))
                Case 3
                     Me.Txt_ClaveSII.Text = IIf(vFld(Rs("Valor")) = 0, "", vFld(Rs("Valor")))
         End Select
        Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   Me.Txt_UsuarioSII.Text = IIf(Trim(Me.Txt_UsuarioSII.Text) = "", FmtCID(gEmpresa.Rut), Trim(Me.Txt_UsuarioSII.Text))

End Sub

Private Sub Txt_UsuarioSII_KeyPress(KeyAscii As Integer)
Call KeyCID(KeyAscii)
End Sub

Private Function InformacionSii(lAno As Integer, lMes As Integer, Rut As String, Clave As String) As String()
'********** INICIO ****************
   Dim Params As String
   Dim Url As String
   Dim Resp As String
   Dim Termina As Boolean
   Dim Vigente As Boolean
   Dim i As Integer
   Dim x As Integer
   Dim v As Integer
   Dim Info() As String
   Dim Traza As String
   
   v = 0
   
    
    
'    If W.InDesign Then
'        Rut = "77765060-2"
'        Clave = "romal1"
'        lAno = "2023"
'        lMes = "04"
'    End If
'
   
   'Params = "rut=17533256&dv=1&referencia=https%3A%2F%2Fmisiir.sii.cl%2Fcgi_misii%2Fsiihome.cgi&411=%20&rutcntr=17533256-1&clave=Fpriet4512"
   Params = "rut=" & vFmtCID(Rut) & "&dv=" & DV_Rut(vFmtCID(Rut)) & "&referencia=https%3A%2F%2Fmisiir.sii.cl%2Fcgi_misii%2Fsiihome.cgi&411=%20&rutcntr=" & vFmtCID(Rut) & "-" & DV_Rut(vFmtCID(Rut)) & "&clave=" & Clave
   Url = URL_SII_LOGIN
   Resp = FwPostPageSII2(Url, Params, "application/x-www-form-urlencoded", SII_LOGIN)
   'La_Title = gLexContab
   If Val(InStr(1, Resp, "titulo", vbTextCompare)) > 0 Then
        'MsgBox1 Replace(Utf8Ansi(Trim(ReplaceStr(ReplaceStr(ReplaceStr(ReplaceStr(GetMensajeSII(FwGetXmlTag(Resp, "div", 1)), "<br>", ""), "<p>", ""), vbLf, ""), vbCr, ""))), "                ", vbLf), vbInformation
        MsgBox1 "Error con la informacion ingresada, Favor verificar su Clave"
        valida = False
        Exit Function
   End If

   'Params = "rut_arrastre=17533256&dv_arrastre=1&cbanoinformemensual=2018&cbmesinformemensual=07&pagina_solicitada=0&CmdXls=Ver_como_planilla"
   Params = "rut_arrastre=" & vFmtCID(Rut) & "&dv_arrastre=" & DV_Rut(vFmtCID(Rut)) & "&cbanoinformemensual=" & lAno & "&cbmesinformemensual=" & Format(lMes, "00") & "&pagina_solicitada=0&CmdXls=Ver_como_planilla"
   Url = URL_SII_RETENCIONES
   Resp = FwPostPageSII2(Url, Params, "application/x-www-form-urlencoded", SII_RETENCIONES)
   'La_Title = gLexContab
   If Val(InStr(1, Resp, "titulo", vbTextCompare)) > 0 Then
        MsgBox1 GetMensajeSII(FwGetXmlTag(Resp, "div", 1)), vbInformation
        valida = False
        'Exit Sub
   
   End If
   
   Termina = False
   i = 3
   
   Do While Not Termina
     Vigente = False
        i = i + 1
        If Not IsNumeric(FwGetXmlTag(FwGetXmlTag(Resp, "tr", i), "td", 1)) Then
            Termina = True
            Exit Do
        End If
        Traza = ""
        For x = 1 To 13
            If x < 8 Then
                Traza = Traza & FwGetXmlTag(FwGetXmlTag(Resp, "tr", i), "td", x) & ";"
                If Trim(FwGetXmlTag(FwGetXmlTag(Resp, "tr", i), "td", x)) = "VIGENTE" Then
                    Vigente = True
                End If
            Else
                Traza = Traza & FwGetXmlTag(FwGetXmlTag(FwGetXmlTag(Resp, "tr", i), "td", x), "div", 1) & ";"
            End If
        Next x
        If Trim(Traza) <> "" And Vigente Then
            ReDim Preserve Info(v)
            Info(v) = Mid(Traza, 1, Len(Traza) - 1)
            v = v + 1
        End If
   Loop

   If v > 0 Then
    InformacionSii = Info
   Else
    MsgBox1 "No tiene informacion a importar para el año " & lAno & " mes " & LCase(gNomMes(lMes)), vbInformation
    valida = False
   End If
   
   
   Url = URL_SII_LOGOUT
   Resp = FwPostPageSII2(Url, "", "application/x-www-form-urlencoded", SII_LOGOUT)
   'La_Title = gLexContab
   If Val(InStr(1, Resp, "titulo", vbTextCompare)) > 0 Then
        MsgBox1 GetMensajeSII(FwGetXmlTag(Resp, "div", 1)), vbInformation
        valida = False
        'Exit Sub
   End If
   
   
'********* FIN ***********
End Function
