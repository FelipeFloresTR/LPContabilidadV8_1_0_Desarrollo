VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmImpCartola 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresar o Importar Cartola"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   Icon            =   "FrmImpCartola.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   1
      Left            =   60
      TabIndex        =   26
      Top             =   0
      Width           =   8955
      Begin VB.CommandButton bt_Help 
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
         Left            =   1680
         Picture         =   "FrmImpCartola.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Ver formato archivo para importar cartola"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton bt_DelCartola 
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
         Left            =   540
         Picture         =   "FrmImpCartola.frx":03EF
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Eliminar Cartola"
         Top             =   180
         Width           =   435
      End
      Begin VB.CommandButton Bt_Cancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   7740
         TabIndex        =   19
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   6600
         TabIndex        =   18
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton bt_Del 
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
         Left            =   1200
         Picture         =   "FrmImpCartola.frx":0873
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Eliminar línea de detalle"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton bt_Imp 
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
         Left            =   120
         Picture         =   "FrmImpCartola.frx":0C6F
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Seleccionar archivo para importar cartola"
         Top             =   180
         Width           =   375
      End
   End
   Begin FlexEdGrid2.FEd2Grid Grid1 
      Height          =   5295
      Left            =   360
      TabIndex        =   14
      Top             =   2040
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9340
      Cols            =   8
      Rows            =   24
      FixedCols       =   0
      FixedRows       =   1
      ScrollBars      =   3
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Index           =   0
      Left            =   360
      TabIndex        =   24
      Top             =   840
      Width           =   8295
      Begin VB.TextBox Tx_SaldoIni 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6780
         MaxLength       =   14
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cb_Banco 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
      Begin VB.CommandButton Bt_Buscar 
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
         Left            =   1620
         Picture         =   "FrmImpCartola.frx":0FF8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar número de cartola"
         Top             =   660
         Width           =   375
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   6540
         Picture         =   "FrmImpCartola.frx":1436
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   660
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   4200
         Picture         =   "FrmImpCartola.frx":1740
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   660
         Width           =   230
      End
      Begin VB.TextBox Tx_Fecha 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   3000
         MaxLength       =   11
         TabIndex        =   8
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox Tx_Fecha 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   5340
         MaxLength       =   11
         TabIndex        =   11
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox Tx_NCart 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         MaxLength       =   3
         TabIndex        =   5
         Top             =   660
         Width           =   555
      End
      Begin VB.CheckBox Ck_NoImp 
         Caption         =   "Sin detalle"
         Height          =   315
         Left            =   7080
         TabIndex        =   13
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Saldo inicial:"
         Height          =   195
         Index           =   5
         Left            =   5820
         TabIndex        =   2
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Banco:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° &Cartola:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Desde:"
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Hasta:"
         Height          =   195
         Index           =   3
         Left            =   4860
         TabIndex        =   10
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.TextBox Tx_Info 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7920
      Width           =   8295
   End
   Begin VB.TextBox Tx_Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   4
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   7500
      Width           =   1635
   End
   Begin VB.TextBox Tx_Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Index           =   3
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   7500
      Width           =   1635
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      DialogTitle     =   "Seleccionar Cartila"
      Filter          =   "Texto separado por tabulaciones (*.txt)|*.txt"
   End
   Begin VB.Label Label1 
      Caption         =   "Sin Detalle permite sólo ingresar número de cartola,  período y totales."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   25
      Top             =   8280
      Width           =   8295
   End
   Begin VB.Label La_Totales 
      AutoSize        =   -1  'True
      Caption         =   "Totales:"
      Height          =   195
      Left            =   2820
      TabIndex        =   23
      Top             =   7560
      Width           =   570
   End
End
Attribute VB_Name = "FrmImpCartola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_FECHA = 0
Private Const C_DETALLE = 1
Private Const C_NRODOC = 2
Private Const C_CARGO = 3
Private Const C_ABONO = 4
Private Const C_HFECHA = 5
Private Const C_ESTADO = 6
Private Const C_ID = 7

Dim lOper As Integer
Dim lidCartola As Long
Dim lTotAbono As Double
Dim lTotCargo As Double
Dim bMod As Boolean
Dim lHayMovConciliados As Boolean

Private Sub Bt_Buscar_Click()
   Dim Nro As Integer
   
   If cb_Banco.ListIndex < 0 Then
      MsgBox1 "Debe seleccionar un banco.", vbExclamation
      Exit Sub
   End If
   
   Nro = Val(Tx_NCart)
   If Nro <= 0 Then
      MsgBox1 "Ingrese el número de la cartola.", vbExclamation
      Exit Sub
   End If
   
   Call LoadAll
   Call EnabHab(False)
   Tx_NCart.SetFocus
   
End Sub

Private Sub Bt_Cancel_Click()

   If bMod Then
      gRc.Rc = vbOK
   End If

   Unload Me
   
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   
   Row = Grid1.Row
   
   If Trim(Grid1.TextMatrix(Row, C_FECHA)) = "" Then
      Exit Sub
   End If
   
   If MsgBox1("¿Esta seguro de eliminar este detalle " & Grid1.TextMatrix(Row, C_DETALLE), vbQuestion Or vbDefaultButton2 Or vbYesNo) <> vbYes Then
      Exit Sub
   End If
   
   Grid1.RowHeight(Row) = 0
   Grid1.TextMatrix(Row, C_ESTADO) = FGR_D
   Grid1.rows = Grid1.rows + 1
   Call Total
   
End Sub

Private Sub bt_DelCartola_Click()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim ExistMov As Boolean
   Dim Msg As String
   Dim Rc As Long

   If cb_Banco.ListIndex < 0 Then
      MsgBox1 "Debe seleccionar un banco.", vbExclamation
      Exit Sub
   End If
   
   If lidCartola = 0 And Trim(Grid1.TextMatrix(1, C_FECHA)) = "" Then
      MsgBox1 "La cartola número " & Tx_NCart & " no existe", vbExclamation
      Exit Sub
   End If
   
   If lidCartola <> 0 Then
      Q1 = "SELECT idCartola FROM MovComprobante WHERE idCartola=" & lidCartola
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         Msg = "ATENCIÓN" & vbCrLf & vbCrLf & "Algunos movimientos de esta cartola están conciliados. Si la elimina, se desconciliarán los movimientos correspondientes." & vbCrLf & vbCrLf & "Esta acción no podrá ser cancelada posteriormente." & vbNewLine & vbNewLine & "¿Desea continuar?"
         ExistMov = True
      Else
         Msg = "¿Está seguro de eliminar la cartola número " & Tx_NCart & "?"
      End If
      Call CloseRs(Rs)
      
      If MsgBox1(Msg, vbDefaultButton2 Or vbQuestion Or vbYesNo) <> vbYes Then
         Exit Sub
      End If
      
      'If ExistMov Then
         'Desmarco cartolas
         Q1 = "UPDATE MovComprobante SET idCartola=0 WHERE idCartola=" & lidCartola
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      'End If
      
      '3376884
      Call SeguimientoMovComprobante(0, gEmpresa.id, gEmpresa.Ano, "FrmImpCartola.bt_DelCartola_Click()", Q1, 1, "WHERE idCartola=" & lidCartola & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano, 1, 2)
      'fin 3376884
      
'      Q1 = "DELETE * FROM DetCartola WHERE idCartola=" & lidCartola
'      Rc = ExecSQL(DbMain, Q1)
      Q1 = " WHERE idCartola=" & lidCartola
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Rc = DeleteSQL(DbMain, "DetCartola", Q1)
      
'      Q1 = "DELETE * FROM Cartola WHERE idCartola=" & lidCartola
'      Rc = ExecSQL(DbMain, Q1)
      Q1 = " WHERE idCartola=" & lidCartola
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Rc = DeleteSQL(DbMain, "Cartola", Q1)
      
      bMod = True
   Else
      MsgBox1 "Esta cartola número " & Tx_NCart & " no esta grabada, pero se limpiarán los datos ingresados.", vbExclamation
   End If
   
   'Ahora igual se hace una limpieza en caso q haya ingresado varias filas y luego quiere eliminarlas sin haber grabado
   Tx_Info = ""
   Call Cb_Banco_Click
   Call LoadAll

End Sub
Private Sub Bt_Fecha_Click(Index As Integer)
   Dim Frm As FrmCalendar
   Dim Dt1 As Long, Dt2 As Long

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_Fecha(Index))
   Set Frm = Nothing

   Dt1 = GetTxDate(Tx_Fecha(0))
   Dt2 = GetTxDate(Tx_Fecha(1))
   
   If Dt1 = 0 Or Dt2 = 0 Or Dt1 < Dt2 Then
      Exit Sub
   End If
   
   If Index = 0 Then
      Dt2 = DateSerial(Year(Dt1), month(Dt1), Day(Dt2))
      Call SetTxDate(Tx_Fecha(1), Dt2)
   Else
      Dt1 = DateSerial(Year(Dt2), month(Dt2), Day(Dt1))
      Call SetTxDate(Tx_Fecha(0), Dt1)
   End If
   

End Sub

Private Sub bt_Help_Click()
   Dim Frm As FrmHelpImpCartola
   
   Set Frm = New FrmHelpImpCartola
   Frm.Show vbModal
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Imp_Click()
   Dim fname As String, Fd As Long, Buf As String, p As Long, l As Integer, r As Integer
   Dim Fecha As Long, Detalle As String, NroDoc As String, Cargo As Double, Abono As Double
   Dim TCargo As Double, TAbono As Double, Aux As String
   Dim Dt1 As Long, Dt2 As Long, nMsg1 As Byte

   If Validar(False) = False Then
      Exit Sub
   End If

   On Error Resume Next

   CmDialog1.InitDir = gImportPath
   CmDialog1.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
   CmDialog1.ShowOpen

   If ERR <> 0 Then
      Exit Sub
   End If

   fname = CmDialog1.Filename
   Fd = FreeFile
   
   Open fname For Input As #Fd
   If ERR Then
      MsgErr fname
      Exit Sub
   End If
   
   Dt1 = GetTxDate(Tx_Fecha(0))
   Dt2 = GetTxDate(Tx_Fecha(1))
     
   Grid1.rows = 1
   
   l = 0
   TAbono = 0
   TCargo = 0
   r = 0
   Do Until EOF(Fd)
      
      Line Input #Fd, Buf
      l = l + 1
      If l = 1 Then
         Line Input #Fd, Buf
         l = l + 1
      End If
         
      p = 1
      Buf = Trim(Buf)

      Aux = Trim(NextField2(Buf, p))
      Fecha = GetDate(Aux)
   
      If Fecha < Dt1 Or Fecha > Dt2 Then
         If nMsg1 < 5 Then
            MsgBox1 "El movimiento con fecha " & FmtFecha(Fecha) & " no pertenece al período de la cartola, será omitido.", vbExclamation
            nMsg1 = nMsg1 + 1
         End If
      Else
   
         Detalle = Trim(NextField2(Buf, p))
         NroDoc = Trim(NextField2(Buf, p))
         Cargo = vFmt(NextField2(Buf, p))
         Abono = vFmt(NextField2(Buf, p))
      
         r = r + 1
         Grid1.rows = r + 1
         Grid1.TextMatrix(r, C_HFECHA) = Fecha
         Grid1.TextMatrix(r, C_FECHA) = Format(Fecha, EDATEFMT) 'FmtFecha(Fecha)
         Grid1.TextMatrix(r, C_DETALLE) = Detalle
         Grid1.TextMatrix(r, C_NRODOC) = NroDoc
         Grid1.TextMatrix(r, C_CARGO) = Format(Cargo, NEGBL_NUMFMT)
         Grid1.TextMatrix(r, C_ABONO) = Format(Abono, NEGBL_NUMFMT)
         Grid1.TextMatrix(r, C_ESTADO) = FGR_I
         
         TCargo = TCargo + Cargo
         TAbono = TAbono + Abono
         
      End If
   Loop
   
   Close #Fd

   Call EnableImp(Grid1.rows <= 1)

   Call FGrVRows(Grid1)
   Grid1.rows = Grid1.rows + 1

   Tx_Total(C_CARGO) = Format(TCargo, NEGBL_NUMFMT)
   Tx_Total(C_ABONO) = Format(TAbono, NEGBL_NUMFMT)

End Sub

Private Sub Bt_OK_Click()
   Dim TotAbono As Double, TotCargo As Double
   Dim Rc As Long

   If Validar(True) = False Then
      Exit Sub
   End If

   TotCargo = vFmt(Tx_Total(C_CARGO))
   TotAbono = vFmt(Tx_Total(C_ABONO))
   
   If TotCargo = 0 And TotAbono = 0 Then
   
      If Ck_NoImp.Value Then
         MsgBox1 "Debe ingresar Total Abonos y Total Cargos.", vbExclamation
         Tx_Total(C_CARGO).SetFocus
         Exit Sub
      ElseIf MsgBox1("¡ATENCION!" & vbNewLine & vbNewLine & "No existen totales registrados. ¿Desea continuar?", vbDefaultButton2 Or vbYesNo Or vbQuestion) <> vbYes Then
         Exit Sub
        ' bt_Imp.SetFocus
      End If
      

   End If

   MousePointer = vbHourglass
   DoEvents

   Rc = SaveAll
   
   gRc.Rc = vbOK

   MousePointer = vbDefault

   If Rc Then
      Unload Me
   End If

End Sub

Private Sub Cb_Banco_Click()
   Dim Q1 As String, Dt1 As Long, Dt2 As Long, Rs As Recordset, NCart As Integer
   
   Q1 = "SELECT Max(Cartola) as NCart, Max(FDesde) as Dt1, Max(FHasta) as Dt2"
   Q1 = Q1 & " FROM Cartola"
   Q1 = Q1 & " WHERE idCuentaBco=" & ItemData(cb_Banco)
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      NCart = vFld(Rs(NCart))
      Dt1 = vFld(Rs("Dt1"))
      Dt2 = vFld(Rs("Dt2"))
   End If
   Call CloseRs(Rs)
   
   If NCart = 0 Or Dt1 = 0 Or Dt2 = 0 Then
      NCart = 1
      Dt1 = DateSerial(gEmpresa.Ano, 1, 1)
      Dt2 = DateSerial(gEmpresa.Ano, 2, 1) - 1
   Else
      NCart = NCart + 1
      Dt1 = Dt2 + 1
      Dt2 = DateSerial(Year(Dt2), month(Dt2) + 1, Day(Dt2)) - 1
   End If

   Tx_NCart = NCart
   Call SetTxDate(Tx_Fecha(0), Dt1)
   Call SetTxDate(Tx_Fecha(1), Dt2)
   

   Call EnableImp(True)
   

End Sub

Private Sub Ck_NoImp_Click()
   Dim bDetalle As Boolean
   
   If cb_Banco.ListIndex < 0 And Ck_NoImp.Value Then
      MsgBox1 "Debe seleccionar un banco.", vbExclamation
      Ck_NoImp.Value = 0
      Exit Sub
   End If
   
   If Ck_NoImp.Value And lidCartola <> 0 And Grid1.TextMatrix(Grid1.FixedRows, C_FECHA) <> "" Then
      MsgBox1 "Esta cartola ya fue grabada con detalles, no puede eliminar sus detalles.", vbExclamation
      Ck_NoImp.Value = 0
      Exit Sub
   End If
   
   bt_Imp.Enabled = (Ck_NoImp.Value = 0 And cb_Banco.ListIndex >= 0)
  
   If lidCartola = 0 Or (lidCartola <> 0 And Grid1.TextMatrix(Grid1.FixedRows, C_FECHA) = "") Then
      If Grid1.TextMatrix(Grid1.FixedRows, C_FECHA) <> "" Then
         If MsgBox1("¡ATENCION!" & vbNewLine & "Existen datos ingresados, al seleccionar sin detalle estos datos se perderán." & vbNewLine & "¿Desea continua?", vbDefaultButton2 Or vbYesNo Or vbQuestion) <> vbYes Then
            Ck_NoImp.Value = 0
            Exit Sub
         End If
      End If
      
      Call FGrClear(Grid1)
      If lOper = O_NEW Then
         Tx_Total(C_CARGO) = ""
         Tx_Total(C_ABONO) = ""
      End If
      
   End If
   
   Call SetTxRO(Tx_Total(C_CARGO), Ck_NoImp.Value = 0)
   Call SetTxRO(Tx_Total(C_ABONO), Ck_NoImp.Value = 0)
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   
   gRc.Rc = vbCancel
   lOper = O_NEW
   
   Call SetupForm
   
   Q1 = "SELECT Descripcion, idCuenta FROM Cuentas WHERE Atrib" & ATRIB_CONCILIACION & "<>0"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call FillCombo(cb_Banco, DbMain, Q1, -2)
   
   If cb_Banco.ListCount = 1 Then
      cb_Banco.ListIndex = 0
   End If

End Sub

Private Sub Grid1_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim LDate As Long
   
   'Franca 11/07/05
   Action = vbOK

   Select Case Col
   
      Case C_FECHA

         If Trim(Value) <> "" Then
            LDate = GetDate(Value, "dmy")
            If LDate < GetTxDate(Tx_Fecha(0)) Or LDate > GetTxDate(Tx_Fecha(1)) Then
               MsgBox1 "Fecha fuera del periodo de la cartola.", vbExclamation + vbOKOnly
               If Grid1.Row = Grid1.rows - 1 Then
                  Grid1.rows = Grid1.rows + 1
               End If
               Action = vbCancel
               
            Else
               Value = Format(LDate, EDATEFMT)
               
               If Grid1.Row = Grid1.rows - 1 Then
                  Grid1.rows = Grid1.rows + 1
               End If

            End If
            
         ElseIf vFmt(Grid1.TextMatrix(Row, C_CARGO)) <> 0 Or vFmt(Grid1.TextMatrix(Row, C_ABONO)) <> 0 Then
               MsgBox1 "Debe ingresar una fecha o borrar el contenido del resto de las columnas.", vbExclamation + vbOKOnly
               Action = vbCancel
         End If
   
      Case C_CARGO
         If Value <> "" And vFmt(Grid1.TextMatrix(Row, C_ABONO)) <> 0 Then
            Grid1.TextMatrix(Row, C_ABONO) = ""
         End If
         Value = Format(vFmt(Value), NUMFMT)
         Grid1.TextMatrix(Row, Col) = Value
         Call Total
      
      Case C_ABONO
         If Value <> "" And vFmt(Grid1.TextMatrix(Row, C_CARGO)) <> 0 Then
            Grid1.TextMatrix(Row, C_CARGO) = ""
         End If
         Value = Format(vFmt(Value), NUMFMT)
         Grid1.TextMatrix(Row, Col) = Value
         Call Total
   
   End Select
   
   If Action = vbOK Then
      Call FGrModRow(Grid1, Row, FGR_U, C_ID, C_ESTADO)
   End If
      
End Sub

Private Sub Grid1_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   
   If cb_Banco.ListIndex = -1 Then
      MsgBox1 "Debe seleccionar banco.", vbExclamation
      cb_Banco.SetFocus
      Exit Sub
   End If
   
   If Ck_NoImp.Value Then
      Exit Sub
   End If
   
   'franca 11/7/2005
'   If Trim(Grid1.TextMatrix(Row - 1, C_FECHA)) = "" Or Trim(Grid1.TextMatrix(Row - 1, C_DETALLE)) = "" Or Trim(Grid1.TextMatrix(Row - 1, C_NRODOC)) = "" Then
'      Exit Sub
'   End If
   
   'Ahora debe permitir ingresar datos si no tiene nº de doc
   If Trim(Grid1.TextMatrix(Row - 1, C_FECHA)) = "" Or Trim(Grid1.TextMatrix(Row - 1, C_DETALLE)) = "" And Trim(Grid1.TextMatrix(Row - 1, C_ESTADO)) <> FGR_D Then
      Exit Sub
   End If
   
   'If Col <> C_FECHA And Col = C_CARGO And Col = C_ABONO Then
   If Col <> C_FECHA Then
      If Col = C_DETALLE And Trim(Grid1.TextMatrix(Row, Col - 1)) = "" Then
         Exit Sub
      ElseIf (Col = C_NRODOC Or Col = C_CARGO Or Col = C_ABONO) And Trim(Grid1.TextMatrix(Row, C_DETALLE)) = "" Then
         Exit Sub
      End If
   End If
   
   'If Grid1.TextMatrix(Row - 1, C_FECHA) = "" And (Col = C_FECHA) Then
   '   Exit Sub
   'End If
   
   EdType = FEG_Edit
   
   If Col = C_FECHA And Trim(Grid1.TextMatrix(Row, C_FECHA)) = "" Then
      Grid1.TextMatrix(Row, Col) = Format(GetTxDate(Tx_Fecha(0)), EDATEFMT)
   End If
   
End Sub

Private Sub Grid1_EditKeyPress(KeyAscii As Integer)
   Dim Col As Integer
   
   Col = Grid1.Col
   
   If Col = C_NRODOC Or Col = C_CARGO Or Col = C_ABONO Then
      Call KeyNum(KeyAscii)
   End If
      
End Sub

Private Sub Grid1_SelChange()
   Tx_Info = Grid1.Text
End Sub

Private Sub Tx_Fecha_GotFocus(Index As Integer)
   
   Call DtGotFocus(Tx_Fecha(Index))

End Sub

Private Sub Tx_Fecha_LostFocus(Index As Integer)
   
   Call DtLostFocus(Tx_Fecha(Index))
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   Dim NroCart As Integer
      
   NroCart = Val(Tx_NCart)
   
   Q1 = "SELECT * FROM Cartola "
   Q1 = Q1 & " WHERE Cartola.Ano=" & gEmpresa.Ano & " AND Cartola.Cartola=" & NroCart
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND idCuentaBco=" & ItemData(cb_Banco)
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid1.rows = 1
   Row = 1
   
   Tx_Total(C_CARGO) = ""
   Tx_Total(C_ABONO) = ""
   lidCartola = 0
   
   If Rs.EOF = False Then
      Call SetTxDate(Tx_Fecha(0), vFld(Rs("FDesde")))
      Call SetTxDate(Tx_Fecha(1), vFld(Rs("FHasta")))
      
      lidCartola = vFld(Rs("idCartola"))
      lTotAbono = vFmt(vFld(Rs("TotAbono")))
      lTotCargo = vFmt(vFld(Rs("TotCargo")))
      
      Tx_SaldoIni = Format(vFld(Rs("SaldoIni")), NUMFMT)
      
      Call CloseRs(Rs)
      
      Q1 = "SELECT *"
      Q1 = Q1 & " FROM DetCartola "
      Q1 = Q1 & " WHERE idCartola=" & lidCartola
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      'Q1 = Q1 & " ORDER BY DetCartola.idMov DESC"
      Q1 = Q1 & " ORDER BY Fecha, DetCartola.idDetCartola"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         lTotAbono = 0
         lTotCargo = 0
      End If
      
      lHayMovConciliados = False
      
      Do While Rs.EOF = False
         Grid1.rows = Row + 1
         
         Grid1.TextMatrix(Row, C_FECHA) = Format(vFld(Rs("Fecha")), EDATEFMT)
         Grid1.TextMatrix(Row, C_DETALLE) = vFld(Rs("Detalle"), True)
         Grid1.TextMatrix(Row, C_NRODOC) = IIf(Val(vFld(Rs("NumDoc"))) = 0, "", vFld(Rs("NumDoc")))
         Grid1.TextMatrix(Row, C_CARGO) = Format(vFld(Rs("Cargo")), BL_NUMFMT)
         Grid1.TextMatrix(Row, C_ABONO) = Format(vFld(Rs("Abono")), BL_NUMFMT)
         Grid1.TextMatrix(Row, C_ID) = vFld(Rs("idDetCartola"))
         
         lTotAbono = lTotAbono + vFld(Rs("Abono"))
         lTotCargo = lTotCargo + vFld(Rs("Cargo"))
         
         If vFld(Rs("IdMov")) <> 0 Then
            lHayMovConciliados = True
         End If
         
         Row = Row + 1
         
         Rs.MoveNext
      Loop
      Call CloseRs(Rs)
      lOper = O_EDIT
            
      Tx_Total(C_CARGO) = Format(lTotCargo, NUMFMT)
      Tx_Total(C_ABONO) = Format(lTotAbono, NUMFMT)
            
   Else
      lOper = O_NEW
      Call CloseRs(Rs)
   End If
   
   Call FGrVRows(Grid1)
   Grid1.rows = Grid1.rows + 1

   Ck_NoImp.Value = Abs(Row <= 1 And lidCartola <> 0)
   
   If lHayMovConciliados Then
      MsgBox1 "Si realiza modificaciones en esta cartola, los movimientos conciliados quedarán sin conciliar.", vbInformation
   End If
   
End Sub

Private Sub SetupForm()

   Grid1.TextMatrix(0, C_FECHA) = "Fecha"
   Grid1.TextMatrix(0, C_DETALLE) = "Detalle"
   Grid1.TextMatrix(0, C_NRODOC) = "Nro. Doc."
   Grid1.TextMatrix(0, C_CARGO) = "Cargo"
   Grid1.TextMatrix(0, C_ABONO) = "Abono"

   Grid1.ColWidth(C_FECHA) = FW_FECHA
   Grid1.ColWidth(C_HFECHA) = 0
   Grid1.ColWidth(C_DETALLE) = 3500
   Grid1.ColWidth(C_NRODOC) = 1000
   Grid1.ColWidth(C_CARGO) = 1200
   Grid1.ColWidth(C_ABONO) = 1200
   Grid1.ColWidth(C_ID) = 0
   Grid1.ColWidth(C_ESTADO) = 0

   Call FGrSetup(Grid1)
   Grid1.rows = Grid1.rows + 1
   
   Grid1.ColAlignment(C_DETALLE) = flexAlignLeftCenter

   Call FGrLocateCntrl(Grid1, La_Totales, C_NRODOC)
   Call FGrLocateCntrl(Grid1, Tx_Total(C_CARGO), C_CARGO)
   Call FGrLocateCntrl(Grid1, Tx_Total(C_ABONO), C_ABONO)

   Call BtFechaImg(Bt_Fecha(0))
   Call BtFechaImg(Bt_Fecha(1))

End Sub

Private Function Validar(bool As Boolean) As Boolean
   Dim Dt1 As Long, Dt2 As Long, Nro As Integer
   Dim Row As Integer
   Dim SinDoc As Boolean, SinMonto As Boolean
   Dim SinDetalle As Boolean
   
   Validar = False

   If cb_Banco.ListIndex < 0 Then
      MsgBox1 "Debe seleccionar un banco.", vbExclamation
      Exit Function
   End If

   Dt1 = GetTxDate(Tx_Fecha(0))
   Dt2 = GetTxDate(Tx_Fecha(1))

   If Year(Dt1) <> gEmpresa.Ano Or Year(Dt2) <> gEmpresa.Ano Then
      MsgBox1 "El período de la cartola no pertenece al año actual " & gEmpresa.Ano & ".", vbExclamation
      Exit Function
   End If

   If Dt1 > Dt2 Then
      MsgBox1 "El orden de las fechas es inválido.", vbExclamation
      Exit Function
   End If

   Nro = Val(Tx_NCart)
   If Nro <= 0 Then
      MsgBox1 "Ingrese el número de la cartola.", vbExclamation
      Exit Function
   End If
   
   'Esta validación es cuando no viene del importador
   If bool Then
      SinDetalle = True
      For Row = 1 To Grid1.rows - 1
         If Grid1.TextMatrix(Row, C_FECHA) <> "" And Grid1.TextMatrix(Row, C_DETALLE) = "" And Trim(Grid1.TextMatrix(Row, C_ESTADO)) <> FGR_D Then
            MsgBox1 "En la línea " & Row - Grid1.FixedRows + 1 & " debe ingresar detalle.", vbExclamation
            Exit Function
         End If
         
         If Grid1.TextMatrix(Row, C_FECHA) <> "" And vFmt(Grid1.TextMatrix(Row, C_NRODOC)) = 0 And Trim(Grid1.TextMatrix(Row, C_ESTADO)) <> FGR_D Then
            SinDoc = True
            
         End If
         
         If Grid1.TextMatrix(Row, C_FECHA) <> "" And vFmt(Grid1.TextMatrix(Row, C_CARGO)) = 0 And vFmt(Grid1.TextMatrix(Row, C_ABONO)) = 0 And Trim(Grid1.TextMatrix(Row, C_ESTADO)) <> FGR_D Then
            SinMonto = True
         End If
         
         If Grid1.TextMatrix(Row, C_FECHA) <> "" And Grid1.TextMatrix(Row, C_ESTADO) <> FGR_D Then
            SinDetalle = False
         End If
         
      Next Row

      If Ck_NoImp.Value = False And SinDetalle Then
         MsgBox1 "No ha ingresado detalle a la cartola " & Nro, vbExclamation
         Exit Function
      End If
      
      If SinDoc Or SinMonto Then
         MsgBox1 "¡ATENCION!" & vbNewLine & vbNewLine & "Existen registros sin número de documento o con valor cero", vbExclamation
      End If
      
   Else
      If Grid1.TextMatrix(1, C_FECHA) <> "" Then
         If MsgBox1("¡Atención!" & vbNewLine & "Existen datos ingresados que se perderán." & vbNewLine & "¿Desea continuar?", vbYesNo Or vbDefaultButton2 Or vbQuestion) <> vbYes Then
            Exit Function
         End If
      End If
   End If
   
   Validar = True
   
End Function

Private Function SaveAll() As Boolean
   Dim Q1 As String, Rc As Long, Rs As Recordset, Q2 As String
   Dim r As Integer, Dt1 As Long, Dt2 As Long, Ano As Integer, NroCart As Integer
   Dim idCartola As Long, TotCargo As Double, TotAbono As Double, idMov As Long
   Dim ModCartola As Boolean, idCtaBco As Long, SaldoIni As Double
   Dim LDate As Long
   Dim FldArray(3) As AdvTbAddNew_t
   
   SaveAll = False
   ModCartola = False
   
   Dt1 = GetTxDate(Tx_Fecha(0))
   Dt2 = GetTxDate(Tx_Fecha(1))
   Ano = Year(Dt1)
   idCtaBco = ItemData(cb_Banco)
   NroCart = Val(Tx_NCart)
   TotCargo = vFmt(Tx_Total(C_CARGO))
   TotAbono = vFmt(Tx_Total(C_ABONO))
   SaldoIni = vFmt(Tx_SaldoIni)
   
   idCartola = 0
   idMov = 0
   
   If gDbType = SQL_SERVER Then
        Q1 = "SELECT IIF(TYPE_NAME(c.user_type_id) = 'int', 1,0) AS type_name  "
        Q1 = Q1 & " FROM sys.objects AS o"
        Q1 = Q1 & " JOIN sys.columns AS c  ON o.object_id = c.object_id"
        Q1 = Q1 & " WHERE O.name = 'DetCartola'"
        Q1 = Q1 & " AND c.name = 'NumDoc'"
        Set Rs = OpenRs(DbMain, Q1)
        
        If Rs.EOF = False Then
          If vFld(Rs("type_name")) = 1 Then
                Q1 = "ALTER TABLE DetCartola DROP CONSTRAINT DF_DetCartola_NumDoc;"
                Q1 = Q1 & " DROP INDEX [NumDoc] ON [dbo].[DetCartola];"
                Q1 = Q1 & " ALTER TABLE DetCartola ALTER COLUMN NumDoc NUMERIC;"
                Q1 = Q1 & " ALTER TABLE [dbo].[DetCartola] ADD  CONSTRAINT [DF_DetCartola_NumDoc]  DEFAULT ((0)) FOR [NumDoc];"
                Q1 = Q1 & " CREATE NONCLUSTERED INDEX [NumDoc] ON [dbo].[DetCartola]"
                Q1 = Q1 & " ("
                Q1 = Q1 & "     [IdEmpresa] ASC,"
                Q1 = Q1 & "     [Ano] ASC,"
                Q1 = Q1 & "     [NumDoc] Asc"
                '3401961
                'Q1 = Q1 & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]"
                 Q1 = Q1 & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
                '3401961
                Call ExecSQL(DbMain, Q1)
          End If
        End If
   Else
      'Call TipoDatoNumDocAccess
   End If
   
   
   If lOper = O_NEW Then
   
      Q1 = "SELECT Cartola.idCartola, DetCartola.idMov"
      Q1 = Q1 & " FROM Cartola LEFT JOIN DetCartola ON Cartola.IdCartola = DetCartola.IdCartola "
      Q1 = Q1 & " AND Cartola.IdEmpresa = DetCartola.IdEmpresa AND Cartola.Ano = DetCartola.Ano "
      Q1 = Q1 & JoinEmpAno(gDbType, "Cartola", "DetCartola")
      Q1 = Q1 & " WHERE Cartola.Ano=" & Ano & " AND Cartola.Cartola=" & NroCart
      Q1 = Q1 & " AND Cartola.idCuentaBco=" & idCtaBco
      Q1 = Q1 & " AND Cartola.IdEmpresa = " & gEmpresa.id
      Q1 = Q1 & " ORDER BY DetCartola.idMov DESC"
      
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         idCartola = vFld(Rs("idCartola"))
         idMov = vFld(Rs("idMov"))
      End If
      
      Call CloseRs(Rs)
      
      If idCartola Then
      
         If idMov = 0 Then
            If MsgBox1("Ya existe la cartola " & NroCart & " del año " & Ano & "." & vbLf & "¿Desea reemplazarla?", vbYesNo Or vbQuestion Or vbDefaultButton2) <> vbYes Then
               Exit Function
            End If
            
            'por si las moscas
            Q1 = "UPDATE MovComprobante SET idCartola=0 WHERE idCartola=" & idCartola
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Call ExecSQL(DbMain, Q1)
            
            '3376884
            Call SeguimientoMovComprobante(0, gEmpresa.id, gEmpresa.Ano, "FrmImpCartola.SaveAll", Q1, 1, "WHERE idCartola=" & lidCartola & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano, 1, 2)
            'fin 3376884
            
'            Q1 = "DELETE * FROM DetCartola WHERE idCartola=" & idCartola
'            Rc = ExecSQL(DbMain, Q1)
            Q1 = " WHERE idCartola=" & idCartola
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Rc = DeleteSQL(DbMain, "DetCartola", Q1)
            
'            Q1 = "DELETE * FROM Cartola WHERE idCartola=" & idCartola
'            Rc = ExecSQL(DbMain, Q1)
            Q1 = " WHERE idCartola=" & idCartola
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Rc = DeleteSQL(DbMain, "Cartola", Q1)
            
            bMod = True

         Else
            MsgBox1 "La cartola " & NroCart & " del año " & Ano & " tiene movimientos conciliados, no puede reemplazarse.", vbExclamation
            Exit Function
         End If
         
      End If
      
   End If
   
   If lOper = O_NEW Then
      FldArray(0).FldName = "IdCuentaBco"
      FldArray(0).FldValue = idCtaBco
      FldArray(0).FldIsNum = True
      
      FldArray(1).FldName = "Ano"
      FldArray(1).FldValue = Ano
      FldArray(1).FldIsNum = True
      
      FldArray(2).FldName = "Cartola"
      FldArray(2).FldValue = NroCart
      FldArray(2).FldIsNum = True
      
      FldArray(3).FldName = "IdEmpresa"
      FldArray(3).FldValue = gEmpresa.id
      FldArray(3).FldIsNum = True
      
      lidCartola = AdvTbAddNewMult(DbMain, "Cartola", "idCartola", FldArray)
      
   End If
      
   Q1 = "UPDATE Cartola SET"
   Q1 = Q1 & " IdCuentaBco=" & idCtaBco
   Q1 = Q1 & ", Ano=" & Ano
   Q1 = Q1 & ", Cartola=" & NroCart
   Q1 = Q1 & ", FDesde=" & Dt1
   Q1 = Q1 & ", FHasta=" & Dt2
   Q1 = Q1 & ", TotCargo=" & Str0(TotCargo)
   Q1 = Q1 & ", TotAbono=" & Str0(TotAbono)
   Q1 = Q1 & ", SaldoIni=" & Str0(SaldoIni)
   Q1 = Q1 & " WHERE idCartola=" & lidCartola
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Rc = ExecSQL(DbMain, Q1)
   bMod = True
   
   If Ck_NoImp.Value = 0 Then
      For r = 1 To Grid1.rows - 1
         If Grid1.TextMatrix(r, C_FECHA) = "" Then
            Exit For
         End If
         
         'PS 27/09/2005 agregué variable LDate, porque se mareaba con el mes al hacer VFmtDate
         LDate = GetDate(Grid1.TextMatrix(r, C_FECHA))
         If Grid1.TextMatrix(r, C_ESTADO) = FGR_I Then
            Q1 = "INSERT INTO DetCartola (IdCartola, Fecha, Detalle, NumDoc, Cargo, Abono, IdEmpresa, Ano) VALUES (" & lidCartola
            Q1 = Q1 & "," & LDate 'VFmtDate(Grid1.TextMatrix(r, C_FECHA))
            Q1 = Q1 & ",'" & ParaSQL(Left(Grid1.TextMatrix(r, C_DETALLE), 30)) & "'"
            Q1 = Q1 & "," & Int(Val(Grid1.TextMatrix(r, C_NRODOC)))
            Q1 = Q1 & "," & Str0(vFmt(Grid1.TextMatrix(r, C_CARGO)))
            Q1 = Q1 & "," & Str0(vFmt(Grid1.TextMatrix(r, C_ABONO)))
            Q1 = Q1 & "," & gEmpresa.id
            Q1 = Q1 & "," & gEmpresa.Ano
            Q1 = Q1 & ")"
            Rc = ExecSQL(DbMain, Q1)
            bMod = True
            
         ElseIf Grid1.TextMatrix(r, C_ESTADO) = FGR_U Then
            Q1 = "UPDATE DetCartola SET"
            Q1 = Q1 & "  Fecha=" & LDate 'VFmtDate(Grid1.TextMatrix(r, C_FECHA))
            Q1 = Q1 & ", Detalle='" & ParaSQL(Left(Grid1.TextMatrix(r, C_DETALLE), 30)) & "'"
            Q1 = Q1 & ", NumDoc=" & Int(Val(Grid1.TextMatrix(r, C_NRODOC)))
            Q1 = Q1 & ", Cargo=" & Str0(vFmt(Grid1.TextMatrix(r, C_CARGO)))
            Q1 = Q1 & ", Abono=" & Str0(vFmt(Grid1.TextMatrix(r, C_ABONO)))
            Q1 = Q1 & " WHERE idDetCartola=" & Grid1.TextMatrix(r, C_ID)
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Rc = ExecSQL(DbMain, Q1)
            bMod = True
            
         ElseIf Grid1.TextMatrix(r, C_ESTADO) = FGR_D And Val(Grid1.TextMatrix(r, C_ID)) <> 0 Then
'            Q1 = "DELETE * FROM DetCartola "
'            Q1 = Q1 & " WHERE idDetCartola=" & Grid1.TextMatrix(r, C_ID)
'            Rc = ExecSQL(DbMain, Q1)
            Q1 = " WHERE idDetCartola=" & Grid1.TextMatrix(r, C_ID)
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Rc = DeleteSQL(DbMain, "DetCartola", Q1)
            bMod = True
            
         End If
      Next r
   End If
   
   'Por si hizo alguna conciliación con esta cartola
   If (lTotAbono <> vFmt(Tx_Total(C_ABONO)) Or lTotCargo <> vFmt(Tx_Total(C_CARGO))) And lidCartola <> 0 Then
      Q1 = "UPDATE Cartola SET"
      Q1 = Q1 & " TotCargo=" & Str0(TotCargo)
      Q1 = Q1 & ", TotAbono=" & Str0(TotAbono)
      Q1 = Q1 & " WHERE idCartola=" & lidCartola
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Rc = ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE MovComprobante SET idCartola=0 WHERE idCartola=" & lidCartola
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Rc = ExecSQL(DbMain, Q1)
      
      '3376884
      Call SeguimientoMovComprobante(0, gEmpresa.id, gEmpresa.Ano, "FrmImpCartola.SaveAll", Q1, 1, "WHERE idCartola=" & lidCartola & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano, 1, 2)
      'fin 3376884
      
      Q1 = "UPDATE DetCartola SET IdMov=0 WHERE idCartola=" & lidCartola
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Rc = ExecSQL(DbMain, Q1)
      
      If lHayMovConciliados Then
         MsgBox1 "Debido a los cambios realizados en esta cartola, todos los movimientos conciliados quedan sin conciliar.", vbInformation
      End If
      
      bMod = True

   End If
   
   SaveAll = True

End Function

Public Sub EnableImp(ByVal bEnab As Boolean)

   Tx_NCart.Enabled = bEnab
   Tx_Fecha(0).Enabled = bEnab
   Tx_Fecha(1).Enabled = bEnab
   Bt_Fecha(0).Enabled = bEnab
   Bt_Fecha(1).Enabled = bEnab
   Tx_SaldoIni.Enabled = bEnab
   
   bt_Imp.Enabled = (bEnab And (Ck_NoImp.Value = 0))
   Ck_NoImp.Enabled = bEnab
   Call EnabHab(False)
End Sub

Private Sub Tx_NCart_Click()
     Call EnabHab(True)
End Sub

Private Sub Tx_SaldoIni_GotFocus()
   Call NumGotFocus(Tx_SaldoIni)
End Sub

Private Sub Tx_SaldoIni_LostFocus()
   Call NumLostFocus(Tx_SaldoIni)

End Sub

Private Sub Tx_Total_GotFocus(Index As Integer)

   If Tx_Total(Index).Locked Then
      Exit Sub
   End If
   
   Tx_Total(Index) = vFmt(Tx_Total(Index))

End Sub

Private Sub Tx_Total_LostFocus(Index As Integer)

   If Tx_Total(Index).Locked Then
      Exit Sub
   End If
   
   Tx_Total(Index) = Format(vFmt(Tx_Total(Index)), NEGBL_NUMFMT)

End Sub
Private Sub Total()
   Dim Row As Integer
   
   Tx_Total(C_CARGO) = 0
   Tx_Total(C_ABONO) = 0
   For Row = 1 To Grid1.rows - 1
      If Grid1.RowHeight(Row) > 0 Then
         Tx_Total(C_CARGO) = Format(vFmt(Grid1.TextMatrix(Row, C_CARGO)) + vFmt(Tx_Total(C_CARGO)), NUMFMT)
         Tx_Total(C_ABONO) = Format(vFmt(Grid1.TextMatrix(Row, C_ABONO)) + vFmt(Tx_Total(C_ABONO)), NUMFMT)
      End If
   Next Row
End Sub
Private Sub EnabHab(bool As Boolean)

   Bt_Buscar.Enabled = bool
   Bt_OK.Enabled = Not bool
   Grid1.Enabled = Not bool
   Bt_Del.Enabled = Not bool
   bt_DelCartola.Enabled = Not bool
   
End Sub
Private Sub TipoDatoNumDocAccess()
Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

On Error Resume Next
ERR.Clear
      
      'Agregamos campo idCCosto a MovActivoFijo
      Set Tbl = DbMain.TableDefs("DetCartola")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NumDoc1", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.idCCosto", vbExclamation
         'lUpdOK = False
      End If
      
        Q1 = "UPDATE DETCARTOLA SET NUMDOC1 = NUMDOC  WHERE IdDetCartola = IdDetCartola"
        Rc = ExecSQL(DbMain, Q1)
        
        Q1 = "ALTER TABLE DETCARTOLA DROP COLUMN NUMDOC"
        Rc = ExecSQL(DbMain, Q1)
        
        'Agregamos campo idCCosto a MovActivoFijo
      Set Tbl = DbMain.TableDefs("DetCartola")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NumDoc", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.idCCosto", vbExclamation
         'lUpdOK = False
      End If
      
      Q1 = "UPDATE DETCARTOLA SET NUMDOC = NUMDOC1  WHERE IdDetCartola = IdDetCartola"
        Rc = ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE DETCARTOLA DROP COLUMN NUMDOC1"
        Rc = ExecSQL(DbMain, Q1)



End Sub
