VERSION 5.00
Begin VB.Form FrmImpLibComprasSII 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Registro de Compras SII (Formato CSV)"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
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
         Caption         =   "Integración SII"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   240
         Width           =   1815
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
      Left            =   480
      TabIndex        =   13
      Top             =   3460
      Width           =   7215
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
      Left            =   7920
      TabIndex        =   12
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Bt_Manual 
      Caption         =   "Manual Uso"
      Height          =   795
      Left            =   6840
      Picture         =   "FrmImpLibComprasSII.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ver Manual Libro Electrónico de Compras"
      Top             =   480
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   540
      Picture         =   "FrmImpLibComprasSII.frx":06B6
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
      Left            =   420
      TabIndex        =   6
      Top             =   1620
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
         Picture         =   "FrmImpLibComprasSII.frx":0C43
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmImpLibComprasSII"
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
   If Me.Op_Manual Then
        If lFName = "" Then
           MsgBox1 "Debe seleccionar el archivo.", vbExclamation + vbOKOnly
           Exit Sub
        End If
        
        Q1 = "UPDATE ParamEmpresa "
        Q1 = Q1 & " SET Valor = '1'"
        Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND Tipo = 'DATOSII' AND Codigo = 1 "

        Call ExecSQL(DbMain, Q1)
        
   Else
        
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
        
   End If

   Rc = MsgBox1("Atención:" & vbNewLine & vbNewLine & "Se importará en el sistema mediante Integración con el SII…" & vbNewLine & vbNewLine & lFName & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
   If Rc = vbNo Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   If Me.Op_Manual Then
        Call Import_LibroComprasSII(Me, lFName, Val(Tx_Ano), CbItemData(Cb_Mes))
   Else
        valida = True
        Info = InformacionSii(Val(Tx_Ano), CbItemData(Cb_Mes), Me.Txt_UsuarioSII.Text, Me.Txt_ClaveSII.Text)
        If valida Then
            Call Import_LibroComprasSIIAuto(Me, lFName, Val(Tx_Ano), CbItemData(Cb_Mes), Info)
        End If
        'Call Import_LibroComprasSIIAuto(Me, lFName, Val(Tx_Ano), CbItemData(Cb_Mes), Me.Txt_UsuarioSII.Text, Me.Txt_ClaveSII.Text)
   End If
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
   
   Tx_Ano = gEmpresa.Ano
   Me.Op_Manual.Value = True
   'comentar para pasar a produccion req SII
'   Op_Automatico.Enabled = False
'   Frame3.visible = False
'   Frame1.visible = False
'   Frame1.Enabled = False

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
                        Frame2.Enabled = True
                        Frame1.Enabled = False
                     Else
                        Me.Op_Automatico = True
                        Frame2.Enabled = False
                        Frame1.Enabled = True
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


Private Sub Op_Automatico_Click()
    Me.Tx_FName.Text = ""
    Frame2.Enabled = False
    Frame1.Enabled = True
    Me.Bt_InfoAyuda.Enabled = False
End Sub

Private Sub Op_Manual_Click()
    Frame2.Enabled = True
    Frame1.Enabled = False
    Me.Bt_InfoAyuda.Enabled = True
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
   Dim i As Integer
   Dim x As Integer
   Dim v As Integer
   Dim ArrayAux() As String
   Dim Info() As String
   Dim Traza As String
   Dim g As GUID
   Dim s As String

    Call CoCreateGuid(g)

    s = Space$(255)

    Call StringFromGUID2(g, ByVal StrPtr(s), Len(s))

    If InStr(1, s, vbNullChar) Then
        s = Left$(s, InStr(1, s, vbNullChar) - 1)
    End If
    
    
'    If W.InDesign Then
'        Rut = "11108309-6"
'        Clave = "3tres"
'        lAno = "2024"
'        lMes = "05"
'    End If
'
   v = 0
   Params = "rut=" & vFmtCID(Rut) & "&dv=" & DV_Rut(vFmtCID(Rut)) & "&referencia=https%3A%2F%2Fmisiir.sii.cl%2Fcgi_misii%2Fsiihome.cgi&411=%20&rutcntr=" & vFmtCID(Rut) & "-" & DV_Rut(vFmtCID(Rut)) & "&clave=" & Clave
   Url = URL_SII_LOGIN
   
   'paso 1 633744
'   MsgBox1 "Paso 1: Login"
'   Call AddLog("Paso 1: Login")
   
   Resp = FwPostPageSII2(Url, Params, "application/x-www-form-urlencoded", SII_LOGIN)
   
   '633744
'    MsgBox1 "Paso 1.1: Respuesta login : " & Resp & " URL: " & Url & " Params : " & Params & " TOKEN_SII : " & TOKEN_SII & ""
'    Call AddLog("Paso 1.1: Respuesta login : " & Resp & " URL: " & Url & " Params : " & Params & " TOKEN_SII : " & TOKEN_SII)

    
   'La_Title = gLexContab
   If Val(InStr(1, Resp, "titulo", vbTextCompare)) > 0 Then
        'MsgBox1 Replace(Utf8Ansi(Trim(ReplaceStr(ReplaceStr(ReplaceStr(ReplaceStr(GetMensajeSII(FwGetXmlTag(Resp, "div", 1)), "<br>", ""), "<p>", ""), vbLf, ""), vbCr, ""))), "                ", vbLf), vbInformation
        MsgBox1 "Error con la informacion ingresada, Favor verificar su Clave"
        valida = False
        Exit Function
   End If

   Params = "{" & Chr(34) & "metaData" & Chr(34) & ":{" & Chr(34) & "namespace" & Chr(34) & ":" & Chr(34) & "cl.sii.sdi.lob.diii.consdcv.data.api.interfaces.FacadeService/getDetalleCompraExport" & Chr(34) & "," & Chr(34) & "conversationId" & Chr(34) & ":" & Chr(34) & TOKEN_SII & Chr(34) & "," & Chr(34) & "transactionId" & Chr(34) & ":" & Chr(34) & Replace(Replace(s, "{", ""), "}", "") & Chr(34) & "," & Chr(34) & "page" & Chr(34) & ":null}," & Chr(34) & "data" & Chr(34) & ":{" & Chr(34) & "rutEmisor" & Chr(34) & ":" & Chr(34) & vFmtCID(Rut) & Chr(34) & "," & Chr(34) & "dvEmisor" & Chr(34) & ":" & Chr(34) & DV_Rut(vFmtCID(Rut)) & Chr(34) & "," & Chr(34) & "ptributario" & Chr(34) & ":" & Chr(34) & lAno & Format(lMes, "00") & Chr(34) & "," & Chr(34) & "codTipoDoc" & Chr(34) & ":0," & Chr(34) & "operacion" & Chr(34) & ":" & Chr(34) & "COMPRA" & Chr(34) & "," & Chr(34) & "estadoContab" & Chr(34) & ":" & Chr(34) & "REGISTRO" & Chr(34) & "}}"
   Url = URL_SII_COMPRA
   'paso 2 633744
'   MsgBox1 "Paso 2: Obtener archivo"
'   Call AddLog("Paso 2: Obtener archivo")
    
   Resp = FwPostPage(Url, Params)
   '633744
'   MsgBox1 "Paso 2.1: Obtener Archivo : " & Resp & " URL: " & Url & " Params : " & Params & " TOKEN_SII : " & TOKEN_SII & ""
'   Call AddLog("Paso 2.1: Obtener Archivo : " & Resp & " URL: " & Url & " Params : " & Params & " TOKEN_SII : " & TOKEN_SII)

    
   ArrayAux = Split(Mid(Resp, InStr(1, Resp, "[", vbTextCompare), InStr(1, Resp, "]", vbTextCompare) - InStr(1, Resp, "[", vbTextCompare) + 1), Chr(34))
   If Val(InStr(1, Resp, "titulo", vbTextCompare)) > 0 Then
        MsgBox1 GetMensajeSII(FwGetXmlTag(Resp, "div", 1)), vbInformation
        valida = False
   End If
   x = 0
   For i = LBound(ArrayAux) To UBound(ArrayAux)
        If IsNumeric(Mid(ArrayAux(i), 1, 1)) Then
             ReDim Preserve Info(x)
             Info(x) = ArrayAux(i)
             x = x + 1
        'ffv  3424117
        ElseIf Mid(ArrayAux(i), 1, 1) = ";" Then
             ReDim Preserve Info(x)
             Info(x) = ArrayAux(i)
             x = x + 1
         
        End If
        '3424117 ffv
   Next
   
   
   If x > 0 Then
    InformacionSii = Info
   Else
    MsgBox1 "No tiene informacion a importar para el año " & lAno & " mes " & LCase(gNomMes(lMes)), vbInformation
    valida = False
   End If
   
   Url = URL_SII_LOGOUT
   Resp = FwPostPageSII2(Url, "", "application/x-www-form-urlencoded", SII_LOGOUT)
   If Val(InStr(1, Resp, "titulo", vbTextCompare)) > 0 Then
        MsgBox1 GetMensajeSII(FwGetXmlTag(Resp, "div", 1)), vbInformation
        valida = False
   End If
   
   
'********* FIN ***********
End Function
