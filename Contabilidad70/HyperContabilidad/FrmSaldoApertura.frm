VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmSaldoApertura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresar/Listar Saldos de Apertura"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   Icon            =   "FrmSaldoApertura.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11115
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   9900
         TabIndex        =   13
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Ok 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   8700
         TabIndex        =   12
         Top             =   180
         Width           =   1095
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
         Left            =   1560
         Picture         =   "FrmSaldoApertura.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Sumar movimientos seleccionados"
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
         Left            =   2520
         Picture         =   "FrmSaldoApertura.frx":00B0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Calculadora"
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
         Left            =   2100
         Picture         =   "FrmSaldoApertura.frx":0411
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Convertir moneda"
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
         Left            =   2940
         Picture         =   "FrmSaldoApertura.frx":07AF
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Calendario"
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
         Left            =   1020
         Picture         =   "FrmSaldoApertura.frx":0BD8
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
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
         Left            =   120
         Picture         =   "FrmSaldoApertura.frx":101D
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
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
         Left            =   540
         Picture         =   "FrmSaldoApertura.frx":14C4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.CommandButton Bt_Detalle 
      Caption         =   "&Detalle"
      Height          =   800
      Left            =   9900
      Picture         =   "FrmSaldoApertura.frx":197E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ingresar detalle saldo de apertura por RUT"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Tx_TotHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   8340
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox Tx_TotDebe 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   7140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   5895
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10398
      Cols            =   5
      Rows            =   2
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
End
Attribute VB_Name = "FrmSaldoApertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CUENTA = 0
Const C_DESC = 1
Const C_DEBE = 2
Const C_HABER = 3
Const C_ID = 4

Const NCOLS = C_ID

Dim lOrientacion As Integer
Dim lEditEnable As Boolean
Dim lMsgDet As Boolean


Private Sub SetUpGrid()
   Dim Col As Integer
   
   Grid.ColWidth(C_CUENTA) = 2300
   Grid.ColWidth(C_DESC) = 4500
   Grid.ColWidth(C_DEBE) = 1300
   Grid.ColWidth(C_HABER) = 1300
   Grid.ColWidth(C_ID) = 0
   
   For Col = 0 To Grid.Cols - 1
      Grid.FixedAlignment(Col) = flexAlignCenterCenter
   Next Col
   
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_DESC) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.TextMatrix(0, C_DESC) = "Descripción"
   Grid.TextMatrix(0, C_DEBE) = "Debe $"
   Grid.TextMatrix(0, C_HABER) = "Haber $"
      
   Call FGrLocateCntrl(Grid, Tx_TotDebe, C_DEBE)
   Call FGrLocateCntrl(Grid, Tx_TotHaber, C_HABER)

   
End Sub

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_Detalle_Click()
   Dim Frm As FrmDetSaldoAp
   Dim IdCuenta As Long
   Dim Saldo As Double
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim Rs As Recordset
   Dim Q1 As String
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   IdCuenta = Grid.TextMatrix(Grid.Row, C_ID)
   If IdCuenta = 0 Then
      Exit Sub
   End If
   
   If Not lMsgDet Then
      MsgBox1 "Los cambios que realice en el Detalle del Saldo de Apertura, no podrán ser posteriormente cancelados desde esta ventana.", vbInformation + vbOKOnly
      lMsgDet = True
   End If
   
   Set Frm = New FrmDetSaldoAp
   Call Frm.FEdit(IdCuenta, Bt_Ok.visible)
   Set Frm = Nothing
   
   Q1 = "SELECT Debe, Haber FROM Cuentas WHERE IdCuenta = " & IdCuenta
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Grid.TextMatrix(Grid.Row, C_DEBE) = Format(vFld(Rs("Debe")), NUMFMT)
      Grid.TextMatrix(Grid.Row, C_HABER) = Format(vFld(Rs("Haber")), NUMFMT)
      Call CalcTot
      
      'puede ser que, con la edición del detalle,
      'los totales de debe y haber hayan quedado distintos.
      'Entonces, deshabilitamos Cancelar para obligar a que los iguale antes de salir.
      If vFmt(Tx_TotDebe) <> vFmt(Tx_TotHaber) Then
         Bt_Cancel.Enabled = False
      End If
   End If
   
   Call CloseRs(Rs)

End Sub

Private Sub bt_OK_Click()
   Dim Q1 As String
   Dim Row As Integer
   
   If vFmt(Tx_TotDebe) <> vFmt(Tx_TotHaber) Then
      MsgBox1 "Los totales de Debe y Haber son distintos.", vbExclamation + vbOKOnly
      Exit Sub
   End If

   'se copian los saldos de apertura financieros a los tributarios (Solicitado por Víctor Morales 9 jul 2020)
   For Row = 1 To Grid.rows - 1
      If Trim(Grid.TextMatrix(Row, C_CUENTA)) <> "" Then
         Q1 = "UPDATE Cuentas SET Debe =" & vFmt(Grid.TextMatrix(Row, C_DEBE))
         Q1 = Q1 & ", DebeTrib =" & vFmt(Grid.TextMatrix(Row, C_DEBE))
         Q1 = Q1 & ", Haber =" & vFmt(Grid.TextMatrix(Row, C_HABER))
         Q1 = Q1 & ", HaberTrib =" & vFmt(Grid.TextMatrix(Row, C_HABER))
         Q1 = Q1 & " WHERE idCuenta=" & Grid.TextMatrix(Row, C_ID)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      End If
      
   Next Row
   
   
   Unload Me
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   lOrientacion = ORIENT_VER
   
   Call SetUpGrid
   
   Call LoadAll
   
   Bt_Detalle.visible = gFunciones.DetSaldoApertura
   
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   Call SetupPriv
   
   Call SetTxRO(Tx_TotDebe, True)
   Call SetTxRO(Tx_TotHaber, True)
   
   'veamos si la empresa tiene historia (año anterior a partir del cual se generó este año)
   If HayAnoAnterior() Then
      MsgBox1 "No es posible modificar los saldos de apertura. Estos se generan automáticamente a partir del año anterior.", vbInformation
      Grid.Locked = True
      Bt_Ok.Enabled = False
      Bt_Ok.visible = False
      Bt_Cancel.Caption = "Cerrar"
      Exit Sub
   End If
   
   'verificamos que no hayan comprobantes de apertura ingresados, que tengan movimientos
   'si hay comprobante de apertura vacío, que se genera automáticamente para guardar número, no damos mensaje
   Q1 = "SELECT Comprobante.IdComp FROM Comprobante INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Tipo=" & TC_APERTURA
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then    'hay al menos un comprobante de tipo Apertura
   
      MsgBox1 "¡Atención!" & vbNewLine & vbNewLine & "Ya existe un comprobante de apertura. Si cambia los saldos, debe generar nuevamente el comprobante de apertura.", vbExclamation + vbOKOnly
      Call CloseRs(Rs)
      Exit Sub
   
   End If
   
   Call CloseRs(Rs)
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   Dim Nivel As Integer
   
   Grid.FlxGrid.Redraw = False
   
   Nivel = gLastNivel
   
   Q1 = "SELECT DISTINCT Cuentas.idCuenta, Cuentas.Codigo, Cuentas.Nivel,  "
   Q1 = Q1 & " Cuentas.Descripcion, Debe, Haber FROM Cuentas  "
   Q1 = Q1 & " WHERE Cuentas.Nivel = " & Nivel
   Q1 = Q1 & " AND Clasificacion IN (" & CLASCTA_ACTIVO & "," & CLASCTA_PASIVO & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Cuentas.Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   
   Row = 1
   Grid.rows = 1
   Do While Rs.EOF = False
      Grid.rows = Row + 1
      
      Grid.TextMatrix(Row, C_CUENTA) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
      Grid.TextMatrix(Row, C_DESC) = FCase(vFld(Rs("Descripcion"), True))
      Grid.TextMatrix(Row, C_DEBE) = Format(vFld(Rs("Debe")), NUMFMT)
      Grid.TextMatrix(Row, C_HABER) = Format(vFld(Rs("Haber")), NUMFMT)
      Grid.TextMatrix(Row, C_ID) = vFld(Rs("idCuenta"))
   
      Row = Row + 1
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   Grid.FlxGrid.Redraw = True
   
   Call CalcTot

End Sub
Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - Tx_TotDebe.Height - 500
   'Grid.Width = Me.Width - 230
   Tx_TotDebe.Top = Grid.Top + Grid.Height + 60
   Tx_TotHaber.Top = Tx_TotDebe.Top
   Call FGrLocateCntrl(Grid, Tx_TotDebe, C_DEBE)
   Call FGrLocateCntrl(Grid, Tx_TotHaber, C_HABER)
   
   Call FGrVRows(Grid)

End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   
   If Col = C_HABER And vFmt(Grid.TextMatrix(Row, C_DEBE)) <> 0 And vFmt(Value) <> 0 Then
      MsgBox1 "Ya existe valor en la columna Debe", vbExclamation
      Value = 0
      
   ElseIf Col = C_DEBE And vFmt(Grid.TextMatrix(Row, C_HABER)) <> 0 Then
      MsgBox1 "Ya existe valor en la columna Haber", vbExclamation
      Value = 0
      
   Else
      
      Value = Format(vFmt(Value), NUMFMT)
      Grid.TextMatrix(Row, Col) = Format(vFmt(Value), NUMFMT)
      Call CalcTot
      
   End If
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FEG2_EdType)
   Dim Q1 As String
   Dim Rs As Recordset
   
   If Col < C_DEBE Then
      Exit Sub
   End If
   
   If Trim(Grid.TextMatrix(Row, C_CUENTA)) = "" Then
      Exit Sub
   End If
   
   'vemos si tiene líneas de detalle
   If gFunciones.DetSaldoApertura Then
      Q1 = "SELECT Id FROM DetSaldosAp WHERE IdCuenta=" & vFmt(Grid.TextMatrix(Row, C_ID))
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         'hay líneas de detalle
         Call PostClick(Bt_Detalle)
         EdType = FEG_None
      
      Else
         EdType = FEG_Edit
      
      End If
   Else
      EdType = FEG_Edit
      
   End If
   
   Call CloseRs(Rs)
   
End Sub

Private Sub Grid_DblClick()
   Dim Col As Integer
   Dim Row As Integer
   
   Col = Grid.MouseCol
   Row = Grid.MouseRow
   
   If Row > Grid.FixedRows Then
      Exit Sub
   End If
   
   If (Col = C_CUENTA Or Col = C_DESC) And Grid.TextMatrix(Row, C_CUENTA) <> "" Then
      Call PostClick(Bt_Detalle)
   End If

End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
   
End Sub

Private Sub CalcTot()
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim i As Integer
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_CUENTA) = "" Then
         Exit For
      End If
      
      TotDebe = TotDebe + Grid.TextMatrix(i, C_DEBE)
      TotHaber = TotHaber + Grid.TextMatrix(i, C_HABER)
      
   Next i
   
   Tx_TotDebe = Format(TotDebe, NUMFMT)
   Tx_TotHaber = Format(TotHaber, NUMFMT)
      
End Sub

Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Me.Caption)
End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
      
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
         
   OldOrientation = Printer.Orientation
      
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
   
End Sub
Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumMov
   
   Set Frm = New FrmSumMov
   
   Call Frm.FViewSum(Grid, C_DEBE, C_HABER)
   
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
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(0) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = lOrientacion
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Caption
   
   gPrtReportes.Titulos = Titulos
      
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
               
   Total(C_DESC) = "Total"
   Total(C_DEBE) = Tx_TotDebe
   Total(C_HABER) = Tx_TotHaber
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_CUENTA
   gPrtReportes.NTotLines = 1
   

End Sub
Private Sub SetupPriv()

   If Not ChkPriv(PRV_CFG_EMP) Then
      Grid.Locked = True
   End If
   
End Sub


