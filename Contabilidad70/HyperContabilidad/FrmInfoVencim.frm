VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmInfoVencim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Vencimientos"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6075
      Left            =   0
      TabIndex        =   8
      Top             =   1740
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   10716
      _Version        =   393216
      Rows            =   20
      Cols            =   8
      FixedRows       =   2
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   0
      TabIndex        =   20
      Top             =   660
      Width           =   10395
      Begin VB.TextBox Tx_Rut 
         Height          =   315
         Left            =   6240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
      End
      Begin VB.ComboBox Cb_Nombre 
         Height          =   315
         Left            =   6240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   2775
      End
      Begin VB.ComboBox Cb_Entidad 
         Height          =   315
         Left            =   7560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   1455
      End
      Begin VB.CheckBox Ch_Rut 
         Caption         =   "RUT:"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   5520
         TabIndex        =   2
         Top             =   240
         Width           =   225
      End
      Begin VB.ComboBox Cb_TipoLib 
         Height          =   315
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   2175
      End
      Begin VB.CommandButton Bt_List 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   9180
         Picture         =   "FrmInfoVencim.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Left            =   960
         TabIndex        =   21
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Left            =   1980
         Picture         =   "FrmInfoVencim.frx":043E
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   230
      End
      Begin VB.ComboBox Cb_Cuentas 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   540
         Width           =   4395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   0
         Left            =   5520
         TabIndex        =   25
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Left            =   5760
         TabIndex        =   18
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label1 
         Caption         =   "Libro:"
         Height          =   255
         Index           =   14
         Left            =   2700
         TabIndex        =   24
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vencim. al:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   10395
      Begin VB.CommandButton Bt_Print 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Picture         =   "FrmInfoVencim.frx":0748
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   9180
         TabIndex        =   17
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Preview 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   660
         Picture         =   "FrmInfoVencim.frx":0C02
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_CopyExcel 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Picture         =   "FrmInfoVencim.frx":10A9
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Calendar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         Picture         =   "FrmInfoVencim.frx":14EE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Calendario"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_ConvMoneda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         Picture         =   "FrmInfoVencim.frx":1917
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Convertir moneda"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Calc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3060
         Picture         =   "FrmInfoVencim.frx":1CB5
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Calculadora"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Sum 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2100
         Picture         =   "FrmInfoVencim.frx":2016
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_VerDoc 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "FrmInfoVencim.frx":20BA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Detalle documento seleccionado"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   0
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7740
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   19
      FixedCols       =   2
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmInfoVencim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDDOC = 0
Const C_FVENC = 1
Const C_FEMISION = 2
Const C_NUMDOC = 3
Const C_RUT = 4
Const C_NOMBRE = 5
Const C_VALOR = 6
Const C_SALDO = 7

Const NCOLS = C_SALDO

Const M_IDENTIDAD = 1
Const M_RUT = 2
Const M_NOTVALIDRUT = 3

Dim lcbNombre As ClsCombo

Dim lOrientacion As Integer

Dim lDias As Integer

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub


Private Sub Bt_Fecha_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_Fecha)
   Bt_List.Enabled = True
   
   Set Frm = Nothing

End Sub

Private Sub Cb_Cuentas_Click()
   Bt_List.Enabled = True

End Sub
Private Sub Cb_TipoLib_Click()
   Bt_List.Enabled = True
End Sub
Private Sub Form_Load()

   lOrientacion = ORIENT_VER

   Call SetTxDate(Tx_Fecha, DateAdd("d", lDias, Now))
   Call BtFechaImg(Bt_Fecha)
      
   Ch_Rut = 1
   Set lcbNombre = New ClsCombo
   Call lcbNombre.SetControl(Cb_Nombre)
   
   Call FillCb

   Call SetUpGrid
   
   Call RecalcSaldos(gEmpresa.Id, gEmpresa.Ano)
   Call RecalcSaldosFulle(gEmpresa.Id, gEmpresa.Ano)
   Call LoadGrid
   
End Sub

Private Sub SetUpGrid()

   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_IDDOC) = 0
   Grid.ColWidth(C_FVENC) = FW_FECHA
   Grid.ColWidth(C_FEMISION) = FW_FECHA
   Grid.ColWidth(C_NUMDOC) = 1600
   Grid.ColWidth(C_RUT) = 1200
   Grid.ColWidth(C_NOMBRE) = 2500
   Grid.ColWidth(C_VALOR) = 1300
   Grid.ColWidth(C_SALDO) = 1300
      
   Grid.ColAlignment(C_FVENC) = flexAlignLeftCenter
   Grid.ColAlignment(C_FEMISION) = flexAlignLeftCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_NOMBRE) = flexAlignLeftCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDO) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_FVENC) = "Fecha"
   Grid.TextMatrix(1, C_FVENC) = "Vencim."
   Grid.TextMatrix(0, C_FEMISION) = "Fecha"
   Grid.TextMatrix(1, C_FEMISION) = "Emisión"
   Grid.TextMatrix(1, C_NUMDOC) = "Documento"
   Grid.TextMatrix(1, C_RUT) = "RUT"
   Grid.TextMatrix(1, C_NOMBRE) = "Nombre Entidad"
   Grid.TextMatrix(1, C_VALOR) = "Total"
   Grid.TextMatrix(1, C_SALDO) = "Saldo"
   
   Call FGrTotales(Grid, GridTot)

End Sub
Private Sub FillCb()
   Dim i As Integer

   Call FillCbCuentas(Cb_Cuentas, True)
      
   Cb_TipoLib.AddItem ReplaceStr(gTipoLib(LIB_COMPRAS), "Libro de ", "")
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = LIB_COMPRAS
   Cb_TipoLib.AddItem ReplaceStr(gTipoLib(LIB_VENTAS), "Libro de ", "")
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = LIB_VENTAS
   Cb_TipoLib.AddItem ReplaceStr(gTipoLib(LIB_RETEN), "Libro de ", "")
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = LIB_RETEN
   Cb_TipoLib.AddItem ReplaceStr(gTipoLib(LIB_OTROS), "Libro de ", "")
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = LIB_OTROS
   
   Call AddItem(Cb_Entidad, "", -1)
   For i = ENT_CLIENTE To ENT_OTRO
      Call AddItem(Cb_Entidad, gClasifEnt(i), i)
   Next i
   Cb_Entidad.ListIndex = 0     'para no seleccionar ninguno al partir

End Sub
Private Sub Bt_List_Click()
   
   If Trim(Tx_Rut) <> "" And Val(lcbNombre.Matrix(M_IDENTIDAD)) = 0 Then
      MsgBox1 "El RUT ingresado no es válido o no ha sido ingresado al sistema.", vbExclamation + vbOKOnly
      Tx_Rut.SetFocus
      Exit Sub
   End If
   
   Call LoadGrid
   
End Sub
Private Sub LoadGrid(Optional ByVal IdDoc As Long = 0)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim Row As Integer
   Dim Total As Double
   Dim TotSaldo As Double
   Dim NotValidRut As Boolean
         
   Grid.Redraw = False
      
   If Trim(Tx_Rut) <> "" Then
      IdEnt = GetIdEntidad(Trim(Tx_Rut), NombEnt, NotValidRut)
      If IdEnt > 0 Then
         Where = Where & " AND Documento.IdEntidad = " & IdEnt
      Else
         Tx_Rut = ""
         Cb_Entidad.ListIndex = 0
         Cb_Nombre.ListIndex = 0
      End If
   
   End If
      
   If ItemData(Cb_TipoLib) > 0 Then
      Where = Where & " AND Documento.TipoLib = " & ItemData(Cb_TipoLib)
   End If
   
   If ItemData(Cb_Cuentas) > 0 Then
      Where = Where & " AND MovDocumento.IdCuenta = " & ItemData(Cb_Cuentas)
   End If
      
   If Tx_Fecha <> "" Then
      Where = Where & " AND (FVenc > 0 AND Documento.FVenc <= " & GetTxDate(Tx_Fecha) & ")"
   End If

   Where = Where & " AND Documento.SaldoDoc <> 0 AND (EsTotalDoc <> 0 OR EsTotalDoc IS NULL)"
   
'   If ItemData(Cb_TipoLib) > 0 And ItemData(Cb_TipoLib) <> LIB_OTROS Then
'      Where = Where & " AND EsTotalDoc <> 0 "
'   End If
   
   
   If Where <> "" Then
      Where = " WHERE " & Mid(Where, 6)
   End If

   Q1 = "SELECT Documento.IdDoc, TipoLib, TipoDoc, NumDoc, NumDocHasta, Documento.IdEntidad, Entidades.Rut, "
   Q1 = Q1 & " Entidades.Nombre, Entidades.NotValidRut, FEmision, FVenc, Total, Descrip, Documento.Estado, SaldoDoc, "
   Q1 = Q1 & " MovDocumento.Debe, MovDocumento.Haber"
   Q1 = Q1 & " FROM ((Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & "  AND Documento.IdEmpresa = Entidades.IdEmpresa )"
   Q1 = Q1 & " LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento")
   Q1 = Q1 & Where
   'Q1 = Q1 & " GROUP BY Documento.IdDoc, TipoLib, TipoDoc, NumDoc, NumDocHasta, Documento.IdEntidad, Entidades.Rut, "
   'Q1 = Q1 & " Entidades.Nombre, Entidades.NotValidRut, FEmision, FVenc, Total, Descrip, Documento.Estado, SaldoDoc "
   
   If Where <> "" Then
      Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.Id & " AND Documento.Ano = " & gEmpresa.Ano
   Else
      Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.Id & " AND Ano = " & gEmpresa.Ano
   End If
   
   Q1 = Q1 & " ORDER BY FVenc, Entidades.Nombre, NumDoc"

   Set Rs = OpenRs(DbMain, Q1)

   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   Do While Rs.EOF = False
      
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
      
      If IdDoc > 0 And vFld(Rs("IdDoc")) = IdDoc Then
         Row = i
      End If
      
      Grid.TextMatrix(i, C_FVENC) = Format(vFld(Rs("FVenc")), SDATEFMT)
      Grid.TextMatrix(i, C_FEMISION) = Format(vFld(Rs("FEmision")), SDATEFMT)
      Grid.TextMatrix(i, C_NUMDOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc"))) & " " & vFld(Rs("NumDoc"))
      
      If vFld(Rs("IdEntidad")) <> 0 Then
         Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = False)
         Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"), True)
      End If
      
      
      Grid.TextMatrix(i, C_VALOR) = Format(vFld(Rs("Total")), NUMFMT)
      Total = Total + vFld(Rs("Total"))
      
      Grid.TextMatrix(i, C_SALDO) = Format(vFld(Rs("SaldoDoc")), NEGNUMFMT)
      TotSaldo = TotSaldo + vFld(Rs("SaldoDoc"))
                               
      Rs.MoveNext
      i = i + 1
   Loop

   Call CloseRs(Rs)
   
   GridTot.TextMatrix(0, C_FVENC) = "TOTAL"
   GridTot.TextMatrix(0, C_VALOR) = Format(Total, NUMFMT)
   GridTot.TextMatrix(0, C_SALDO) = Format(TotSaldo, NEGNUMFMT)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
      
   If Row = 0 Then
      Row = Grid.FixedRows
   End If
   
   Call FGrSelRow(Grid, Row)
      
   Grid.Redraw = True
   Bt_List.Enabled = False
   
End Sub


Private Sub Ch_Rut_Click()
   Bt_List.Enabled = True

End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - GridTot.Height - 500
   GridTot.Top = Grid.Top + Grid.Height + 30
   'Grid.Width = Me.Width - 230
   GridTot.Width = Grid.Width - 230
   
   Call FGrVRows(Grid)

End Sub

Private Sub Grid_DblClick()
   
   If Grid.MouseRow < Grid.FixedRows Then
      Exit Sub
   End If
   
   Call PostClick(Bt_VerDoc)
      
End Sub

Private Sub Tx_Rut_Change()
   Bt_List.Enabled = True
End Sub

Private Sub Tx_Rut_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = vbKeyReturn Then
      Call Tx_Rut_LostFocus
      KeyAscii = 0
   ElseIf Ch_Rut <> 0 Then
      Call KeyCID(KeyAscii)
   Else
      Call KeyName(KeyAscii)
      Call KeyUpper(KeyAscii)
   End If
   
End Sub
Private Sub Tx_RUT_Validate(Cancel As Boolean)
   
   If Tx_Rut = "" Then
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_Rut, Ch_Rut <> 0) Then
      Cancel = True
      Exit Sub
   End If
   
End Sub

Private Sub Tx_Rut_LostFocus()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdEnt As Long
   Dim i As Integer
   Dim AuxRut As String

   If Tx_Rut = "" Then
      Cb_Entidad.ListIndex = 0  'en blanco
      Exit Sub
   End If
         
   Q1 = "SELECT IdEntidad, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5 FROM Entidades WHERE Rut = '" & vFmtCID(Tx_Rut, Ch_Rut <> 0) & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.Id
   Set Rs = OpenRs(DbMain, Q1)
   
   IdEnt = 0
   
   If Rs.EOF = False Then   'existe
      IdEnt = vFld(Rs("IdEntidad"))
            
      'seleccionamos el tipo de entidad y esto llena la lista de nombres de entidades
      For i = 0 To MAX_ENTCLASIF
         If Cb_Entidad.ItemData(i) >= 0 Then
            If vFld(Rs("Clasif" & Cb_Entidad.ItemData(i))) <> 0 Then
               Cb_Entidad.ListIndex = i
               Exit For
            End If
         End If
      Next i
   
      'ahora seleccionamos la entidad
      For i = 0 To Cb_Nombre.ListCount - 1
         If lcbNombre.Matrix(M_IDENTIDAD, i) = IdEnt Then
            lcbNombre.ListIndex = i
            Exit For
         End If
      Next i
      
      Bt_List.Enabled = True

   Else
      MsgBox1 "Este RUT no ha sido ingresado al sistema.", vbExclamation + vbOKOnly
      Cb_Entidad.ListIndex = -1
      
   End If
      
   Call CloseRs(Rs)
   
   If Ch_Rut <> 0 Then
      AuxRut = FmtCID(vFmtCID(Tx_Rut))
      If AuxRut <> "0-0" Then
         Tx_Rut = AuxRut
      End If
   End If
   
End Sub
Private Sub cb_Nombre_Click()
   
   If lcbNombre.ListIndex >= 0 Then
      Tx_Rut = FmtCID(lcbNombre.Matrix(M_RUT), Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0)
      Ch_Rut = IIf(Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0, 1, 0)
   End If
   
   Bt_List.Enabled = True

End Sub
Private Sub Cb_Entidad_Click()
      
   Cb_Nombre.Clear
   If ItemData(Cb_Entidad) >= 0 Then
      Call SelCbEntidad(ItemData(Cb_Entidad))
   Else
      Tx_Rut = ""
   End If
   
   Bt_List.Enabled = True

End Sub

Private Sub SelCbEntidad(Clasif As Integer)
   Dim Q1 As String
   
   lcbNombre.Clear
   If Clasif >= 0 Then
      Q1 = "SELECT Nombre, idEntidad, Rut, NotValidRut FROM Entidades"
      Q1 = Q1 & " WHERE Clasif" & Clasif & "=" & CON_CLASIF
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.Id
      Q1 = Q1 & " ORDER BY Nombre "
      Call lcbNombre.FillCombo(DbMain, Q1, -1)
   End If
   
End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Tit As String
   
   If Bt_List.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de copiar.", vbExclamation
      Exit Sub
   End If
   
   If Tx_Rut <> "" Then
      Tit = "RUT: " & Tx_Rut & " " & Cb_Nombre
   End If
   
   If ItemData(Cb_Cuentas) > 0 Then
      Tit = Tit & " Cuenta: " & Cb_Cuentas
   End If
   
   Call FGr2Clip(Grid, Me.Caption & " al " & Tx_Fecha & " " & Tit)
End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   
   If Bt_List.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de seleccionar al vista previa.", vbExclamation
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientation As Integer
         
   If Bt_List.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   OldOrientation = Printer.Orientation
      
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
   
End Sub
Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid)
   
   Set Frm = Nothing

End Sub
Private Sub Bt_ConvMoneda_Click()
   Dim Frm As FrmConverMoneda
   Dim Valor As Double
      
   Set Frm = New FrmConverMoneda
   Frm.FView (Valor)
      
   Set Frm = Nothing
   
End Sub
Private Sub Bt_Calc_Click()
   Call Calculadora
End Sub
Private Sub Bt_Calendar_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
   
   Call Frm.SelDate(Fecha)
   
   Set Frm = Nothing

End Sub
Private Sub Bt_VerDoc_Click()
   Dim Row As Integer
   Dim Col As Integer
   Dim Frm As FrmDoc
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 Then
      Set Frm = New FrmDoc
      Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDDOC)))
      Set Frm = Nothing
   End If

End Sub

Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(2) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = lOrientacion
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Caption & " Al " & Tx_Fecha
   If ItemData(Cb_Cuentas) > 0 Then
      Titulos(1) = "Cuenta: " & Cb_Cuentas
   End If
   If Tx_Rut <> "" Then
      If Titulos(1) = "" Then
         Titulos(1) = Cb_Nombre
      Else
         Titulos(2) = Cb_Nombre
      End If
   End If
   
   gPrtReportes.Titulos = Titulos
      
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
               
   Total(C_NOMBRE) = "Total"
   Total(C_VALOR) = GridTot.TextMatrix(0, C_VALOR)
   Total(C_SALDO) = GridTot.TextMatrix(0, C_SALDO)
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_IDDOC
   gPrtReportes.NTotLines = 1
   

End Sub

Public Sub FView(Optional ByVal Dias As Integer = 30)
   lDias = Dias
   
   Me.Show vbModal
End Sub

