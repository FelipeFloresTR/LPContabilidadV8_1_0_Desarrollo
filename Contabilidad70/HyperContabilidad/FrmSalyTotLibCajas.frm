VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmSalyTotLibCajas 
   Caption         =   "Saldos y Totales Libro de Cajas"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9135
      Left            =   -360
      TabIndex        =   0
      Top             =   -360
      Width           =   12615
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   375
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "SALDOS Y TOTALES LIBRO DE CAJA"
         Top             =   480
         Width           =   11535
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   375
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "MONTOS QUE AFECTAN LA BASE IMPONIBLE"
         Top             =   840
         Width           =   5775
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   375
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "FLUJO DE INGRESOS Y EGRESOS"
         Top             =   840
         Width           =   5775
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
         Left            =   480
         Picture         =   "FrmSalyTotLibCajas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   5760
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   1335
         Left            =   480
         TabIndex        =   1
         Top             =   1200
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   3
         Cols            =   5
         FixedRows       =   2
         FixedCols       =   0
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "FrmSalyTotLibCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fecha As Long
Public GetMonth As Boolean

Private Sub Form_Load()

Call LoadGrid(False)

End Sub
Private Sub LoadGrid(impr As Boolean)
Dim TotFluIng As Long, TotFluEgr As Long, TotBaseIng As Long, TotBaseEgr As Long
Dim grilla As MSFlexGrid

'If Not impr Then
'Call SetUpGrid2
'Else
Call SetUpGrid2
'End If
Call Totales(Fecha, TotFluIng, TotFluEgr, TotBaseIng, TotBaseEgr)


'With Grid
'
'.Row = 4
'.Col = 1
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 1) = TotFluIng
'
'.Row = 4
'.Col = 2
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 2) = TotFluEgr
'
'.Row = 4
'.Col = 3
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 3) = TotFluIng - TotFluEgr
'
'.Row = 4
'.Col = 4
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 4) = TotBaseIng
'
'.Row = 4
'.Col = 5
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 5) = TotBaseEgr
'
'.Row = 4
'.Col = 6
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 6) = TotBaseIng - TotBaseEgr

'End With

With Grid 'IIf(impr, Grid, GridImp)
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 0
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotFluIng, NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
'End With
'
'With grilla
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 1
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotFluEgr, NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
'End With
'
'
'
'With grilla
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 2
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotFluIng - TotFluEgr, NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
'End With
'
'
'
'With grilla
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 3
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotBaseIng, NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
'End With
'
'
'
'With grilla
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 4
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotBaseEgr, NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
'End With
'
'
'
'
'With grilla
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 5
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotBaseIng - Abs(TotBaseEgr), NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
End With




End Sub

Private Sub LoadGrid2(impr As Boolean)
Dim TotFluIng As Long, TotFluEgr As Long, TotBaseIng As Long, TotBaseEgr As Long
Dim grilla As MSFlexGrid

'If Not impr Then
'Call SetUpGrid2
'Else
Call SetUpGrid3
'End If
Call Totales(Fecha, TotFluIng, TotFluEgr, TotBaseIng, TotBaseEgr)


'With Grid
'
'.Row = 4
'.Col = 1
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 1) = TotFluIng
'
'.Row = 4
'.Col = 2
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 2) = TotFluEgr
'
'.Row = 4
'.Col = 3
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 3) = TotFluIng - TotFluEgr
'
'.Row = 4
'.Col = 4
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 4) = TotBaseIng
'
'.Row = 4
'.Col = 5
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 5) = TotBaseEgr
'
'.Row = 4
'.Col = 6
'.RowHeight(3) = 960
'.CellFontBold = False
'.CellFontSize = 7
'.CellAlignment = flexAlignCenterCenter
'.BorderStyle = flexBorderNone
'.FillStyle = flexFillSingle
'.WordWrap = True
'.TextMatrix(4, 6) = TotBaseIng - TotBaseEgr

'End With

With Grid 'IIf(impr, Grid, GridImp)
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 2
.Col = 0
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotFluIng, NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 2
.Col = 1
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotFluEgr, NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
End With



With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 2
.Col = 2
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotFluIng - TotFluEgr, NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
End With



With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 2
.Col = 3
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotBaseIng, NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
End With



With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 2
.Col = 4
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotBaseEgr, NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
End With




With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 2
.Col = 5
.WordWrap = True
.CellAlignment = flexAlignCenterCenter
.Text = Format(TotBaseIng - Abs(TotBaseEgr), NUMFMT)
.FillStyle = flexFillSingle
.CellFontSize = 7
.Redraw = True
End With




End Sub



Private Sub SetUpGrid()

With Grid
.Cols = 7
.rows = 5
.GridLinesFixed = flexGridFlat

.GridLines = flexGridInset

.ColWidth(0) = 0
.RowHeight(0) = 0
.ColWidth(1) = 2500
.ColWidth(2) = 2500
.ColWidth(3) = 2500
.ColWidth(4) = 2500
.ColWidth(5) = 2500
.ColWidth(6) = 2500


.Row = 1
.CellFontBold = True
.CellFontSize = 11
.ColAlignment(1) = flexAlignCenterCenter
.MergeCells = flexMergeRestrictRows
.TextMatrix(1, 1) = "SALDOS Y TOTALES LIBRO DE CAJA"
.TextMatrix(1, 2) = "SALDOS Y TOTALES LIBRO DE CAJA"
.TextMatrix(1, 3) = "SALDOS Y TOTALES LIBRO DE CAJA"
.TextMatrix(1, 4) = "SALDOS Y TOTALES LIBRO DE CAJA"
.TextMatrix(1, 5) = "SALDOS Y TOTALES LIBRO DE CAJA"
.TextMatrix(1, 6) = "SALDOS Y TOTALES LIBRO DE CAJA"
.MergeRow(1) = True


.Row = 2
.CellFontBold = True
.CellFontSize = 9
.ColAlignment(2) = flexAlignCenterCenter
.MergeCells = flexMergeRestrictRows
.TextMatrix(2, 1) = "FLUJO DE INGRESOS Y EGRESOS"
.TextMatrix(2, 2) = "FLUJO DE INGRESOS Y EGRESOS"
.TextMatrix(2, 3) = "FLUJO DE INGRESOS Y EGRESOS"
.MergeRow(2) = True

.Row = 2
.Col = 4
.CellFontBold = True
.CellFontSize = 9
.ColAlignment(4) = flexAlignCenterCenter
.MergeCells = flexMergeRestrictRows
.TextMatrix(2, 4) = "MONTOS QUE AFECTAN LA BASE IMPONIBLE"
.TextMatrix(2, 5) = "MONTOS QUE AFECTAN LA BASE IMPONIBLE"
.TextMatrix(2, 6) = "MONTOS QUE AFECTAN LA BASE IMPONIBLE"
.MergeRow(2) = True

.Row = 3
.Col = 1
.RowHeight(3) = 960
.CellFontBold = False
.CellFontSize = 7
.CellAlignment = flexAlignCenterCenter
.BorderStyle = flexBorderNone
.FillStyle = flexFillSingle
.WordWrap = True
.TextMatrix(3, 1) = "TOTAL MONTO FLUJO DE INGRESOS"

.Row = 3
.Col = 2
.RowHeight(3) = 960
.CellFontBold = False
.CellFontSize = 7
.CellAlignment = flexAlignCenterCenter
.BorderStyle = flexBorderNone
.FillStyle = flexFillSingle
.WordWrap = True
.TextMatrix(3, 2) = "TOTAL MONTO FLUJO DE EGRESOS"

.Row = 3
.Col = 3
.RowHeight(3) = 960
.CellFontBold = False
.CellFontSize = 7
.CellAlignment = flexAlignCenterCenter
.BorderStyle = flexBorderNone
.FillStyle = flexFillSingle
.WordWrap = True
.TextMatrix(3, 3) = "SALDO FLUJO DE CAJA"

.Row = 3
.Col = 4
.RowHeight(3) = 960
.CellFontBold = False
.CellFontSize = 7
.CellAlignment = flexAlignCenterCenter
.BorderStyle = flexBorderNone
.FillStyle = flexFillSingle
.WordWrap = True
.TextMatrix(3, 4) = "INGRESOS"

.Row = 3
.Col = 5
.RowHeight(3) = 960
.CellFontBold = False
.CellFontSize = 7
.CellAlignment = flexAlignCenterCenter
.BorderStyle = flexBorderNone
.FillStyle = flexFillSingle
.WordWrap = True
.TextMatrix(3, 5) = "EGRESOS"

.Row = 3
.Col = 6
.RowHeight(3) = 960
.CellFontBold = False
.CellFontSize = 7
.CellAlignment = flexAlignCenterCenter
.BorderStyle = flexBorderNone
.FillStyle = flexFillSingle
.WordWrap = True
.TextMatrix(3, 6) = "RESULTADO NETO"



End With

End Sub

Private Sub SetUpGrid2()

With Grid
.Cols = 6
.rows = 2
.GridLinesFixed = flexGridFlat

.GridLines = flexGridInset

'.ColWidth(0) = 0
'.RowHeight(0) = 0
.ColWidth(0) = 2010
.ColWidth(1) = 2010
.ColWidth(2) = 1710
.ColWidth(3) = 1900
.ColWidth(4) = 1900
.ColWidth(5) = 1900
.RowHeight(0) = 1000
.CellFontSize = 8


'.Redraw = False
'.FillStyle = flexFillRepeat
'.MergeCells = flexMergeFree
'.MergeRow(0) = True
'.Row = 0
'.Col = 0
'.ColSel = 5
'.CellAlignment = flexAlignCenterCenter
'.Text = "SALDOS Y TOTALES LIBRO DE CAJA"
'.FillStyle = flexFillSingle
'.Redraw = True
'End With

'With Grid
'.Redraw = False
'.FillStyle = flexFillRepeat
'.MergeCells = flexMergeFree
'.MergeRow(1) = True
'.Row = 1
'.Col = 0
'.ColSel = 3
'.CellAlignment = flexAlignCenterCenter
'.Text = "FLUJO DE INGRESOS Y EGRESOS"
'.FillStyle = flexFillSingle
'.Redraw = True
'End With
'
'With Grid
'.Redraw = False
'.FillStyle = flexFillRepeat
'.MergeCells = flexMergeFree
'.MergeRow(1) = True
'.Row = 1
'.Col = 3
'.ColSel = 5
'.CellAlignment = flexAlignCenterCenter
'.Text = "MONTOS QUE AFECTAN LA BASE IMPONIBLE"
'.FillStyle = flexFillSingle
'.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
'.MergeCol(1) = True
.Row = 0
.Col = 0
'.RowSel = 2
.CellAlignment = flexAlignCenterCenter
.Text = "TOTAL MONTO FLUJO DE INGRESOS"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.MergeRow(0) = True
.Row = 0
.Col = 1
.CellAlignment = flexAlignCenterCenter
.Text = "TOTAL MONTO FLUJO DE EGRESOS"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.MergeRow(0) = True
.Row = 0
.Col = 2
.CellAlignment = flexAlignCenterCenter
.Text = "SALDO FLUJO DE CAJA"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.MergeRow(0) = True
.Row = 0
.Col = 3
.CellAlignment = flexAlignCenterCenter
.Text = "INGRESOS"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.MergeRow(0) = True
.Row = 0
.Col = 4
.CellAlignment = flexAlignCenterCenter
.Text = "EGRESOS"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.MergeRow(0) = True
.Row = 0
.Col = 5
.CellAlignment = flexAlignCenterCenter
.Text = "RESULTADO NETO"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With


'End With

End Sub

Private Sub SetUpGrid3()

With Grid
.Cols = 6
.rows = 3
.GridLinesFixed = flexGridFlat

.GridLines = flexGridInset

'.ColWidth(0) = 0
'.RowHeight(0) = 0
.ColWidth(0) = 2010
.ColWidth(1) = 2010
.ColWidth(2) = 1710
.ColWidth(3) = 1900
.ColWidth(4) = 1900
.ColWidth(5) = 1900
.RowHeight(0) = 500
.CellFontSize = 8


'.Redraw = False
'.FillStyle = flexFillRepeat
'.MergeCells = flexMergeFree
'.MergeRow(0) = True
'.Row = 0
'.Col = 0
'.ColSel = 5
'.CellAlignment = flexAlignCenterCenter
'.Text = "SALDOS Y TOTALES LIBRO DE CAJA"
'.FillStyle = flexFillSingle
'.Redraw = True
'End With

'With Grid
'.Redraw = False
'.FillStyle = flexFillRepeat
'.MergeCells = flexMergeFree
'.MergeRow(1) = True
'.Row = 1
'.Col = 0
'.ColSel = 3
'.CellAlignment = flexAlignCenterCenter
'.Text = "FLUJO DE INGRESOS Y EGRESOS"
'.FillStyle = flexFillSingle
'.Redraw = True
'End With
'
'With Grid
'.Redraw = False
'.FillStyle = flexFillRepeat
'.MergeCells = flexMergeFree
'.MergeRow(1) = True
'.Row = 1
'.Col = 3
'.ColSel = 5
'.CellAlignment = flexAlignCenterCenter
'.Text = "MONTOS QUE AFECTAN LA BASE IMPONIBLE"
'.FillStyle = flexFillSingle
'.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
'.MergeCells = flexMergeFree
.Row = 0
.Col = 0
.CellAlignment = flexAlignCenterCenter
.Text = "TOTAL MONTO"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
'.MergeCells = flexMergeFree
.Row = 0
.Col = 1
.CellAlignment = flexAlignCenterCenter
.Text = "TOTAL MONTO"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 0
.Col = 2
.CellAlignment = flexAlignCenterCenter
.Text = "SALDO FLUJO"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 0
.Col = 3
.CellAlignment = flexAlignCenterCenter
.Text = "INGRESOS"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 0
.Col = 4
.CellAlignment = flexAlignCenterCenter
.Text = "EGRESOS"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 0
.Col = 5
.CellAlignment = flexAlignCenterCenter
.Text = "RESULTADO NETO"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With



With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 0
'.RowSel = 2
.CellAlignment = flexAlignCenterCenter
.Text = "FLUJO DE INGRESOS"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 1
.CellAlignment = flexAlignCenterCenter
.Text = "FLUJO DE EGRESOS"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 2
.CellAlignment = flexAlignCenterCenter
.Text = "DE CAJA"
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 3
.CellAlignment = flexAlignCenterCenter
.Text = ""
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 4
.CellAlignment = flexAlignCenterCenter
.Text = ""
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With

With Grid
.Redraw = False
.FillStyle = flexFillRepeat
.MergeCells = flexMergeFree
.Row = 1
.Col = 5
.CellAlignment = flexAlignCenterCenter
.Text = ""
.FillStyle = flexFillSingle
.WordWrap = True
.Redraw = True
End With


'End With

End Sub



Public Sub FView()

Me.Show vbModal


End Sub

'Public Function FPrint(Frm As FrmPrintPreview)
'
'Load Me
'Call MyFrmPrint(Frm)
'Unload Me
'
'End Function
'Private Function MyFrmPrint(frm3 As FrmPrintPreview)
'    Dim Frm As FrmPrintPreview
'    Dim FrmPrt As FrmPrtSetup
'    Call LoadGrid
'    Call SetupPritGrid
'    Set Frm = New FrmPrintPreview
'    Set FrmPrt = Nothing
'    'gPrtLibro.CallEndDoc = True
'    'gPrtLibros.PrtFlexGrid (frm3)
'
'End Function

Public Function GetGrid() As MSFlexGrid

Call LoadGrid2(True)
Set GetGrid = Grid

End Function


'Private Sub SetupPritGrid()
'   Dim i As Integer
'   Dim ColWi(NCOLS) As Integer
'   Dim Total(NCOLS) As String
'   Dim Titulos(2) As String
'   Dim Encabezados(3) As String
'   Dim FontTit(2) As FontDef_t
'   Dim OldOrient As Integer
'   Dim Mes As String
'   Dim Idx As Integer
'
'   Set gPrtLibros.Grid = Grid
'
'   'Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
'
'   Titulos(0) = Me.Caption & " " & gTipoOperCaja(lTipoOper)
'
'   FontTit(0).FontBold = True
'
''   If lOper = O_EDIT Then
''      Titulos(1) = Titulos(1) & gNomMes(CbItemData(Cb_Mes)) & " " & lAno
''   Else
''      Titulos(1) = Titulos(1) & Cb_Mes & " " & Val(Cb_Ano)
''   End If
'
'   If lInfoPreliminar Then
'      Titulos(2) = INFO_PRELIMINAR
'      FontTit(2).FontBold = True
'   End If
'
'   gPrtLibros.Titulos = Titulos
'   Call gPrtLibros.FntTitulos(FontTit())
'
'   gPrtLibros.Encabezados = Encabezados
'
'   gPrtLibros.GrFontName = Grid.Font.Name
'   gPrtLibros.GrFontSize = Grid.Font.Size
'
'   For i = 0 To Grid.Cols - 1
'      ColWi(i) = Grid.ColWidth(i)
'      'Total(i) = GridTot.TextMatrix(0, i)
'   Next i
'
'   ColWi(C_TIPODOC) = 0
'   ColWi(C_DTE) = 0
'   ColWi(C_NOMBRE) = 0
'
'
'   gPrtLibros.ColWi = ColWi
'   gPrtLibros.Total = Total
'   gPrtLibros.NTotLines = 1
'   gPrtLibros.ColObligatoria = C_SALDO
'
'   gPrtLibros.Obs = ""   'para que no ponga las notas
'
''   If Ch_LibOficial <> 0 Then
''      gPrtLibros.PrintFecha = False
''   End If
'
'End Sub

Private Sub Totales(Fecha As Long, ByRef TotFluIng As Long, ByRef TotFluEgr As Long, ByRef TotBaseIng As Long, ByRef TotBaseEgr As Long)
Dim Rs As Recordset
Dim Q1 As String
Dim Tm1 As Long, Tm2 As Long

    Call FirstLastMonthDay(Fecha, Tm1, Tm2)


'    Q1 = " SELECT (SELECT SUM(TOTAL) FROM LibroCaja L1 WHERE L1.FechaIngresoLibro = LC.FechaIngresoLibro AND TIPOOPER = 1) AS Expr1, "
'    Q1 = Q1 & " (SELECT SUM(TOTAL) FROM LibroCaja L1 WHERE L1.FechaIngresoLibro = LC.FechaIngresoLibro AND TIPOOPER = 2) AS Expr2, "
'    Q1 = Q1 & " (SELECT SUM(MontoAfectaBaseImp) FROM LibroCaja L1 WHERE L1.FechaIngresoLibro = LC.FechaIngresoLibro AND TIPOOPER = 1) AS Expr3, "
'    Q1 = Q1 & " (SELECT SUM(MontoAfectaBaseImp) FROM LibroCaja L1 WHERE L1.FechaIngresoLibro = LC.FechaIngresoLibro AND TIPOOPER = 2) AS Expr4 "
'    Q1 = Q1 & " FROM LibroCaja AS LC "
'    If GetMonth Then
'        Q1 = Q1 & " Where (((lc.FechaIngresoLibro) > " & Tm1 & " And (lc.FechaIngresoLibro) < " & Tm2 & ")) "
'    End If
'    Q1 = Q1 & " GROUP BY FechaIngresoLibro; "

    '3289932
    'Q1 = " SELECT SUM(TOTAL) as Expr1,SUM(MontoAfectaBaseImp) as Expr2,TIPOOPER as Expr3 "
    Q1 = " SELECT SUM(pagado) as Expr1,SUM(MontoAfectaBaseImp) as Expr2,TIPOOPER as Expr3  "
    '3289932
    Q1 = Q1 & " FROM LibroCaja AS LC "
    If GetMonth Then
    '3289932
        'Q1 = Q1 & " Where (((lc.FechaIngresoLibro) > " & Tm1 & " And (lc.FechaIngresoLibro) < " & Tm2 & ")) "
        Q1 = Q1 & " Where (((lc.FechaIngresoLibro) >= " & Tm1 & " And (lc.FechaIngresoLibro) <= " & Tm2 & ")) "
    '3289932
    End If
    Q1 = Q1 & " group by TIPOOPER; "


   Set Rs = OpenRs(DbMain, Q1)
   
  
   
   Do Until Rs.EOF
    If vFld(Rs("Expr3")) = 1 Then
      TotFluIng = vFld(Rs("Expr1"))
      '3289932
       'TotBaseIng = vFld(Rs("Expr2"))
       TotBaseIng = vMontoBaseImpoIngreso
      '3289932
    Else
     TotFluEgr = vFld(Rs("Expr1"))
    '3289932
     'TotBaseEgr = vFld(Rs("Expr2"))
     TotBaseEgr = vMontoBaseImpoEgreso
      '3289932
    End If
   Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)


End Sub


'Private Sub Bt_Preview_Click()
'   Dim Frm As FrmPrintPreview
'   Dim PrtOrient As Integer
'   Dim Pag As Integer
'
'   PrtOrient = Printer.Orientation
'
''   If Bt_Search.Enabled = True Then
''      MsgBox1 "Presione el botón Listar antes de seleccionar la vista previa.", vbExclamation
''      Exit Sub
''   End If
'
'   Call SetUpPrtGrid
'
'   Set Frm = Nothing
'   Set Frm = New FrmPrintPreview
'
'   Me.MousePointer = vbHourglass
'   Pag = gPrtReportes.PrtFlexGrid(Frm)
'   Set Frm.PrtControl = Bt_Print
'   Me.MousePointer = vbDefault
'
'   If Pag >= 0 Then
'      Call Frm.FView(Caption)
'   End If
'
'   Set Frm = Nothing
'
'   Printer.Orientation = PrtOrient
'   gPrtReportes.FmtCol = -1
'   Call ResetPrtBas(gPrtReportes)
'
'End Sub

'Private Sub SetUpPrtGrid()
'   Dim i As Integer, j As Integer
'   Dim ColWi(NCOLS) As Integer
'   Dim Total(NCOLS * 3)
'   Dim Titulos(0) As String
'
'   Printer.Orientation = ORIENT_HOR
'   Set gPrtReportes.Grid = Grid
'
'   Titulos(0) = Me.Caption
'   gPrtReportes.Titulos = Titulos
'
'   For i = 0 To Grid.Cols - 1
'      ColWi(i) = Grid.ColWidth(i)
'   Next i
'
'   ColWi(C_TIPOLIB) = ColWi(C_TIPOLIB) - 100 '- 400
'   ColWi(C_TIPODOC) = ColWi(C_TIPODOC) - 100 '- 800
'   ColWi(C_COUNT) = ColWi(C_COUNT) - 100 '- 200
'   ColWi(C_COUNTDTE) = ColWi(C_COUNTDTE) - 100 '- 230
'   ColWi(C_OIMPDTE) = ColWi(C_OIMPDTE) - 100 '- 70
'
'   For i = 0 To Grid.Cols - 1
'      If ColWi(i) > 0 Then
'         ColWi(i) = ColWi(i) * 0.95
'      End If
'   Next i
'
'   gPrtReportes.ColWi = ColWi
'   gPrtReportes.ColObligatoria = C_OBLIGATORIA
'   gPrtReportes.FmtCol = C_FMT
'
'   gPrtReportes.GrFontSize = 7
'   gPrtReportes.GrFontName = "Arial"
'   'gPrtReportes.TotFntBold = False
'
'
''   j = 0
''   For i = 0 To Grid.Cols - 1
''      Total(j) = GridTot(0).TextMatrix(0, i)
''      j = j + 1
''   Next i
''
''   For i = 0 To Grid.Cols - 1
''      Total(j) = GridTot(1).TextMatrix(0, i)
''      j = j + 1
''   Next i
'
'   gPrtReportes.Total = Total
'   gPrtReportes.NTotLines = 2
'
'
'End Sub
Private Sub Text1_Change()

End Sub
