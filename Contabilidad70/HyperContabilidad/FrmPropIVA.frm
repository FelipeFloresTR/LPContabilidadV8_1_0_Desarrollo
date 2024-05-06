VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPropIVA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calcular Proporcionalidad de IVA Crédito Fiscal"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_RecalcTot 
      Caption         =   "Recalcular Totales y  Prop. IVA CF"
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   4500
      Width           =   2835
   End
   Begin VB.CommandButton Bt_AplicarPropIVA 
      Caption         =   "Aplicar  Propocionalidad IVA CF"
      Height          =   315
      Left            =   7140
      TabIndex        =   12
      ToolTipText     =   "Modificar Documentos de Compra de acuerdo a la Proporción de IVA calculada"
      Top             =   4500
      Width           =   2835
   End
   Begin VB.Frame Frame 
      Height          =   555
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10155
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
         Left            =   3120
         Picture         =   "FrmPropIVA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Calculadora"
         Top             =   120
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
         Left            =   2580
         Picture         =   "FrmPropIVA.frx":0361
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Convertir moneda"
         Top             =   120
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
         Left            =   3540
         Picture         =   "FrmPropIVA.frx":06FF
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Calendario"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   8760
         TabIndex        =   11
         Top             =   180
         Width           =   1275
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
         Left            =   1200
         Picture         =   "FrmPropIVA.frx":0B28
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Print 
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
         Picture         =   "FrmPropIVA.frx":0FCF
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprimir listado en pantalla"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_DetLibro 
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
         Picture         =   "FrmPropIVA.frx":1489
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Detalle comprobante seleccionado"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton bt_CopyExcel 
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
         Left            =   2040
         Picture         =   "FrmPropIVA.frx":18EE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Copiar Excel"
         Top             =   120
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
         Left            =   660
         Picture         =   "FrmPropIVA.frx":1D33
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   120
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3495
      Left            =   180
      TabIndex        =   1
      Top             =   720
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   6165
      _Version        =   393216
      BackColorBkg    =   16777215
   End
   Begin MSComctlLib.ProgressBar Pb_Proceso 
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   5220
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmPropIVA.frx":1DD7
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   4
      Left            =   180
      TabIndex        =   18
      Top             =   7440
      Width           =   9795
   End
   Begin VB.Label Label1 
      Caption         =   " - El cálculo de la proporcionalidad se hace acumulado hasta Diciembre de cada año partiendo de $ 0 desde Enero del año siguiente"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   3
      Left            =   180
      TabIndex        =   17
      Top             =   7020
      Width           =   9795
   End
   Begin VB.Label Label1 
      Caption         =   " -  % Prop IVA CF = (Total Afecto / (Total Afecto + Exento)) * 100"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   180
      TabIndex        =   16
      Top             =   6600
      Width           =   9795
   End
   Begin VB.Label Label1 
      Caption         =   " -  Estos datos se calculan a partir de los documentos ingresados en el Libro de Ventas de cada mes."
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   15
      Top             =   6180
      Width           =   9795
   End
   Begin VB.Label Label1 
      Caption         =   "NOTAS:"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   14
      Top             =   5760
      Width           =   615
   End
End
Attribute VB_Name = "FrmPropIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_MES = 0
Const C_STRMES = 1
Const C_TOTAFECTO = 2
Const C_TOTEXENTO = 3
Const C_TOTAL = 4
Const C_CALCPROP = 5
Const C_ACUMAFECTO = 6
Const C_ACUMTOTAL = 7
Const C_PROPORCION = 8

Const NCOLS = C_PROPORCION

Private Sub Bt_AplicarPropIVA_Click()

   Pb_Proceso.Value = 0

   If MsgBox1("¿Desea aplicar la Proporcionalidad del IVA Crédito Fiscal, de acuerdo a los valores aquí calculados?" & vbCrLf & vbCrLf & "Recuerde que este proceso sólo se realiza para los documentos en estado PENDIENTE" & vbCrLf & "y que tienen seleccionada alguna opción en la columna Prop. IVA.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   If gCtasBas.IdCtaIVAIrrec = 0 Then
      MsgBox "ATENCIÖN:  Falta definir la cuenta de IVA Irrecuperable." & vbCrLf & vbCrLf & "Utilice la opción: " & vbCrLf & vbCrLf & "Configuración >> Configuración Inicial >> Definir Cuentas Básicas >> Otros", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
      
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio)
   End If
   
   Call PropIVA_UpdateMovDoc(0, Pb_Proceso)
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")
   End If
   
   Me.MousePointer = vbDefault
   
   'MsgBox1 "El proceso de aplicación de Proporcionalidad de IVA Crédito Fiscal ha finalizado.", vbInformation

End Sub

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, "Proporcionalidad de IVA Crédito Fiscal" & vbTab & "Año " & gEmpresa.Ano)

End Sub

Private Sub Bt_DetLibro_Click()
   Dim Frm As FrmCompraVenta
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Set Frm = New FrmCompraVenta
   Call Frm.FView(LIB_VENTAS, Val(Grid.TextMatrix(Grid.Row, C_MES)))
   Set Frm = Nothing
   
   
End Sub


Private Sub Bt_RecalcTot_Click()
   
   Me.MousePointer = vbHourglass
   
   Pb_Proceso.Value = 50
   
   Call PropIVA_UpdateTblTotMensual(0, True)
   
   Pb_Proceso.Value = 100
   
   Call LoadAll
   
   Me.MousePointer = vbDefault

End Sub

Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid, Grid.Row, Grid.RowSel, Grid.Col, Grid.ColSel)
   
   Set Frm = Nothing

End Sub

Private Sub Form_Activate()
   MsgBox1 "Recuerde Recalcular Totales y Pro. IVA CF en el botón habilitado para este efecto,de modo de ver los impactos en el Libro de Compras.", vbOKOnly + vbInformation
End Sub

Private Sub Form_Load()
   Call SetUpGrid
   
   Call LoadAll

End Sub


Private Sub SetUpGrid()
   Dim i As Integer

   Grid.Cols = NCOLS + 1
   Grid.FixedCols = 2
   Grid.rows = 2 + 12 '12 meses
   Grid.FixedRows = 2
   
   Call FGrSetup(Grid, True)
   
   Grid.ColWidth(C_MES) = 0
   Grid.ColWidth(C_STRMES) = 500
   
   For i = C_TOTAFECTO To Grid.Cols - 1
      Grid.ColWidth(i) = 1300
      Grid.ColAlignment(i) = flexAlignRightCenter
   Next i
   
   Grid.ColAlignment(C_CALCPROP) = flexAlignCenterCenter
   
   
   Grid.TextMatrix(1, C_STRMES) = "Mes"
   Grid.TextMatrix(0, C_TOTAFECTO) = "Total"
   Grid.TextMatrix(1, C_TOTAFECTO) = "Afecto"
   Grid.TextMatrix(0, C_TOTEXENTO) = "Total"
   Grid.TextMatrix(1, C_TOTEXENTO) = "Exento"
   Grid.TextMatrix(0, C_TOTAL) = "Total"
   Grid.TextMatrix(1, C_TOTAL) = "Afecto+Exento"
   Grid.TextMatrix(0, C_CALCPROP) = "Corresp."
   Grid.TextMatrix(1, C_CALCPROP) = "Prop. IVA"
   Grid.TextMatrix(0, C_ACUMAFECTO) = "Acumulado"
   Grid.TextMatrix(1, C_ACUMAFECTO) = "Afecto"
   Grid.TextMatrix(0, C_ACUMTOTAL) = "Acumulado"
   Grid.TextMatrix(1, C_ACUMTOTAL) = "Total"
   Grid.TextMatrix(0, C_PROPORCION) = "% Prop."
   Grid.TextMatrix(1, C_PROPORCION) = "IVA CF"
   
End Sub

Private Sub LoadAll()
   Dim i As Integer, Row As Integer
   Dim UltMes As Integer
   
   UltMes = GetUltimoMesConMovs(True)

   Grid.Redraw = False
   Grid.rows = Grid.FixedRows
   Grid.rows = Grid.FixedRows + 12 '12 meses
   
   For i = 1 To 12
      Row = i + Grid.FixedRows - 1
      Grid.TextMatrix(Row, C_MES) = i
      Grid.TextMatrix(Row, C_STRMES) = Left(gNomMes(i), 3)
      Grid.TextMatrix(Row, C_TOTAFECTO) = Format(gValPropIVA(i).TotAfecto, NEGNUMFMT)
      Grid.TextMatrix(Row, C_TOTEXENTO) = Format(gValPropIVA(i).TotExento, NEGNUMFMT)
      Grid.TextMatrix(Row, C_TOTAL) = Format(gValPropIVA(i).Total, NEGNUMFMT)
      Grid.TextMatrix(Row, C_CALCPROP) = FmtSiNo(gValPropIVA(i).CalcProp, False)
      Grid.TextMatrix(Row, C_ACUMAFECTO) = Format(gValPropIVA(i).AcumAfecto, NEGNUMFMT)
      Grid.TextMatrix(Row, C_ACUMTOTAL) = Format(gValPropIVA(i).AcumTotal, NEGNUMFMT)
      
      If i > UltMes Then
         Grid.TextMatrix(Row, C_PROPORCION) = Format(0, DBLFMT4)
      Else
         Grid.TextMatrix(Row, C_PROPORCION) = Format(gValPropIVA(i).Proporcion * 100, DBLFMT4)
      End If
      
   Next i
   
   Grid.Redraw = True
   
End Sub

Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(1) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = "Proporcionalidad de IVA Crédito Fiscal"
   Titulos(1) = "Año " & gEmpresa.Ano
   gPrtReportes.Titulos = Titulos
         
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
                  
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_MES
   gPrtReportes.NTotLines = 0

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

Private Sub Grid_DblClick()
   Call Bt_DetLibro_Click
End Sub
