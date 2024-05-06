VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmResLibAux 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Libros Auxiliares"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   Icon            =   "FrmResLibAux.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10980
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid GridCod 
      Height          =   5775
      Left            =   8880
      TabIndex        =   20
      Top             =   1620
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   10186
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5835
      Left            =   60
      TabIndex        =   14
      Top             =   1560
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10292
      _Version        =   393216
      Rows            =   20
      Cols            =   20
      FixedRows       =   2
      FixedCols       =   3
      AllowUserResizing=   1
   End
   Begin VB.Frame Fr_Botones 
      Height          =   555
      Left            =   60
      TabIndex        =   19
      Top             =   60
      Width           =   10815
      Begin VB.CommandButton Bt_Resumen 
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
         Picture         =   "FrmResLibAux.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Resumen IVA Compras - Ventas y Otros Impuestos"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_ListCod 
         Caption         =   "Cód."
         Height          =   315
         Left            =   8700
         TabIndex        =   21
         Top             =   180
         Visible         =   0   'False
         Width           =   555
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
         Left            =   1140
         Picture         =   "FrmResLibAux.frx":0484
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Orden 
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
         Left            =   1740
         Picture         =   "FrmResLibAux.frx":0528
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ordenar listado por columna seleccionada"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_CopyExcel 
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
         Picture         =   "FrmResLibAux.frx":0918
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Copiar Excel"
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
         Left            =   2700
         Picture         =   "FrmResLibAux.frx":0D5D
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir listado"
         Top             =   120
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
         Left            =   2280
         Picture         =   "FrmResLibAux.frx":1217
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Vista previa de la impresión"
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
         Left            =   4500
         Picture         =   "FrmResLibAux.frx":16BE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Calendario"
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
         Left            =   3660
         Picture         =   "FrmResLibAux.frx":1AE7
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Convertir moneda"
         Top             =   120
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
         Left            =   4080
         Picture         =   "FrmResLibAux.frx":1E85
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Calculadora"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_DetDoc 
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
         Left            =   720
         Picture         =   "FrmResLibAux.frx":21E6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Detalle libro seleccionado"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Close 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   9540
         TabIndex        =   13
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   60
      TabIndex        =   15
      Top             =   660
      Width           =   10815
      Begin VB.CommandButton Bt_Search 
         Height          =   435
         Left            =   9540
         Picture         =   "FrmResLibAux.frx":264B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   1335
      End
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
      Begin VB.ComboBox Cb_TipoLib 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   2460
         TabIndex        =   18
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   4380
         TabIndex        =   17
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Libro:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   390
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7440
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   503
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7680
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   503
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmResLibAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_MES = 0
Const C_LNGMES = 1
Const C_TIPOLIB = 2
Const C_IDTIPOLIB = 3
Const C_TIPODOCDIM = 4
Const C_TIPODOC = 5
Const C_CLASIF = 6
Const C_COUNT = 7
Const C_AFECTO = 8
Const C_EXENTO = 9
Const C_IVADEB = 10
Const C_IVACRED = 11
Const C_OTROIMP = 12
Const C_IVAIRREC = 13
Const C_IVARET = 14
Const C_COUNTDTE = 15
Const C_IVADTE = 16
Const C_OIMPDTE = 17
Const C_OBLIGATORIA = 18
Const C_FMT = 19

Const NCOLS = C_FMT

Const CC_CODF29 = 0
Const CC_VALOR = 1

Const T_COMPRASVENTAS = 200

Dim lMes As Integer
Dim lHayActFijo As Boolean

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_DetDoc_Click()
   Dim Frm As Form
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_LNGMES)) = 0 Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDTIPOLIB)) = LIB_COMPRAS Or Val(Grid.TextMatrix(Row, C_IDTIPOLIB)) = LIB_VENTAS Then
      Set Frm = New FrmCompraVenta
      Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDTIPOLIB)), Val(Grid.TextMatrix(Row, C_LNGMES)))
      Set Frm = Nothing
   ElseIf Val(Grid.TextMatrix(Row, C_IDTIPOLIB)) = LIB_RETEN Then
      Set Frm = New FrmLibRetenciones
      Call Frm.FView(Val(Grid.TextMatrix(Row, C_LNGMES)))
      Set Frm = Nothing
   End If
   
End Sub

Private Sub Bt_ListCod_Click()
   
   If GridCod.visible = True Then
      GridCod.visible = False
   Else
      GridCod.visible = True
      Call LoadAllCod
      Call FGr2Clip(GridCod, Me.Caption)
   End If
      
End Sub

Private Sub Bt_Search_Click()
   Me.MousePointer = vbHourglass
   GridCod.visible = False
   Call LoadAll
   Me.MousePointer = vbDefault
End Sub

Private Sub Cb_Ano_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_Mes_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_TipoLib_Click()
   Call EnableFrm(True)

End Sub

Private Sub Form_Activate()

   If Grid.visible And lHayActFijo Then
      MsgBox1 "Atención: Una factura puede aparecer duplicada en el libro de compras cuando tiene una parte de Activo Fijo y otra no.", vbInformation
   End If

End Sub

Private Sub Form_Load()

   Bt_ListCod.visible = W.InDesign
   
   Call FillCb
   Call SetUpGrid
   Call SetUpGridCod
   Call LoadAll

End Sub

Private Sub FillCb()
   Dim i As Integer
   Dim MesActual As Integer
   
   Cb_TipoLib.AddItem ""
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = 0
   
   Cb_TipoLib.AddItem ReplaceStr(gTipoLib(LIB_COMPRAS), "Libro de ", "")
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = LIB_COMPRAS
   Cb_TipoLib.AddItem ReplaceStr(gTipoLib(LIB_VENTAS), "Libro de ", "")
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = LIB_VENTAS
'   Cb_TipoLib.AddItem "Compras y Ventas"
'   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = T_COMPRASVENTAS
   Cb_TipoLib.AddItem ReplaceStr(gTipoLib(LIB_RETEN), "Libro de ", "")
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = LIB_RETEN
   
   MesActual = GetMesActual()
      
   Cb_Mes.AddItem " "
   Cb_Mes.ItemData(Cb_Mes.NewIndex) = 0
   
   Call FillMes(Cb_Mes)
               
   If lMes > 0 Then
      Cb_Mes.ListIndex = lMes
   Else
      If MesActual > 0 Then
         Cb_Mes.ListIndex = MesActual
      Else
         Cb_Mes.ListIndex = GetUltimoMesConMovs()
      End If
   End If

'   Cb_Ano.AddItem gEmpresa.Ano - 2
'   Cb_Ano.AddItem gEmpresa.Ano - 1
   Cb_Ano.AddItem gEmpresa.Ano
   Cb_Ano.ListIndex = Cb_Ano.NewIndex

End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.ColWidth(C_MES) = 400
   Grid.ColWidth(C_LNGMES) = 0
   Grid.ColWidth(C_TIPOLIB) = 1100
   Grid.ColWidth(C_IDTIPOLIB) = 0
   Grid.ColWidth(C_TIPODOCDIM) = 0
   Grid.ColWidth(C_TIPODOC) = 2350
   Grid.ColWidth(C_CLASIF) = 1100
   Grid.ColWidth(C_COUNT) = 700
   Grid.ColWidth(C_EXENTO) = 1100
   Grid.ColWidth(C_AFECTO) = 1100
   Grid.ColWidth(C_IVADEB) = 1100
   Grid.ColWidth(C_IVACRED) = 1100
   Grid.ColWidth(C_OTROIMP) = 1100
   Grid.ColWidth(C_IVAIRREC) = 1100
   Grid.ColWidth(C_IVARET) = 1100
   Grid.ColWidth(C_COUNTDTE) = 700
   Grid.ColWidth(C_IVADTE) = 1100
   Grid.ColWidth(C_OIMPDTE) = 1100
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_OBLIGATORIA) = 0
         
   Grid.ColAlignment(C_TIPOLIB) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPODOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPODOCDIM) = flexAlignLeftCenter
   Grid.ColAlignment(C_CLASIF) = flexAlignLeftCenter
   Grid.ColAlignment(C_MES) = flexAlignLeftCenter
   Grid.ColAlignment(C_COUNT) = flexAlignRightCenter
   Grid.ColAlignment(C_EXENTO) = flexAlignRightCenter
   Grid.ColAlignment(C_AFECTO) = flexAlignRightCenter
   Grid.ColAlignment(C_IVADEB) = flexAlignRightCenter
   Grid.ColAlignment(C_IVACRED) = flexAlignRightCenter
   Grid.ColAlignment(C_OTROIMP) = flexAlignRightCenter
   Grid.ColAlignment(C_IVAIRREC) = flexAlignRightCenter
   Grid.ColAlignment(C_IVARET) = flexAlignRightCenter
   Grid.ColAlignment(C_COUNTDTE) = flexAlignRightCenter
   Grid.ColAlignment(C_IVADTE) = flexAlignRightCenter
   Grid.ColAlignment(C_OIMPDTE) = flexAlignRightCenter
   
   Grid.TextMatrix(1, C_MES) = "Mes"
   Grid.TextMatrix(1, C_TIPOLIB) = "Libro"
   Grid.TextMatrix(1, C_TIPODOCDIM) = ""
   Grid.TextMatrix(1, C_TIPODOC) = "Tipo Documento"
   Grid.TextMatrix(1, C_CLASIF) = "Clasific."
   Grid.TextMatrix(0, C_COUNT) = "Cant."
   Grid.TextMatrix(1, C_COUNT) = "Total"
   Grid.TextMatrix(1, C_EXENTO) = "Exento"
   Grid.TextMatrix(1, C_AFECTO) = "Afecto"
   Grid.TextMatrix(0, C_IVADEB) = "IVA"
   Grid.TextMatrix(1, C_IVADEB) = "Débito"
   Grid.TextMatrix(0, C_IVACRED) = "IVA"
   Grid.TextMatrix(1, C_IVACRED) = "Crédito"
   Grid.TextMatrix(0, C_OTROIMP) = "Otros"
   Grid.TextMatrix(1, C_OTROIMP) = "Impuestos"
   Grid.TextMatrix(0, C_IVAIRREC) = "IVA"
   Grid.TextMatrix(1, C_IVAIRREC) = "Irrecuperable"
   Grid.TextMatrix(0, C_IVARET) = "IVA"
   Grid.TextMatrix(1, C_IVARET) = "Retenido"
   Grid.TextMatrix(0, C_COUNTDTE) = "Cant."
   Grid.TextMatrix(1, C_COUNTDTE) = "DTE"
   Grid.TextMatrix(1, C_IVADTE) = "IVA DTE"
   Grid.TextMatrix(0, C_OIMPDTE) = "Otros Imp."
   Grid.TextMatrix(1, C_OIMPDTE) = "DTE"
   Grid.TextMatrix(0, C_FMT) = ""
        
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot(0))
   Call FGrTotales(Grid, GridTot(1))
   
'   GridTot(0).RowHeight(1) = 270
'   GridTot(0).Height = GridTot(0).RowHeight(0) * 2 + 30
'   GridTot(1).RowHeight(1) = 270
'   GridTot(1).Height = GridTot(1).RowHeight(0) * 2 + 30

   
   Call FGrVRows(Grid)
   Grid.rows = Grid.rows + 1
   Grid.TopRow = Grid.FixedRows
   
End Sub
Private Sub SetUpGridCod()
   Dim i As Integer
   
   GridCod.ColWidth(CC_CODF29) = 500
   GridCod.ColWidth(CC_VALOR) = 1000

   GridCod.ColAlignment(CC_CODF29) = flexAlignRightCenter
   GridCod.ColAlignment(CC_VALOR) = flexAlignRightCenter
   
   GridCod.TextMatrix(0, CC_CODF29) = "Cód."
   GridCod.TextMatrix(0, CC_VALOR) = "Valor"

   For i = 0 To GridCod.Cols - 1
      GridCod.FixedAlignment(i) = flexAlignCenterCenter
   Next i
   
   Call FGrVRows(GridCod)
   GridCod.rows = Grid.rows + 1
   GridCod.TopRow = Grid.FixedRows

End Sub
Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - GridTot(0).Height - GridTot(1).Height - 500
   GridTot(0).Top = Grid.Top + Grid.Height + 30
   GridTot(1).Top = GridTot(0).Top + GridTot(0).Height + 30
   Grid.Width = Me.Width - 200
   GridTot(0).Width = Grid.Width - 230
   GridTot(1).Width = Grid.Width - 230
   GridCod.Left = Grid.Width - GridCod.Width
   GridCod.Height = Grid.Height
   
   Call FGrVRows(Grid)

End Sub
Private Sub LoadAll()
   Dim Where As String
   Dim i As Integer
   Dim CurReg As String
   Dim PrevReg As String
   Dim Total(NCOLS) As Double
   Dim ResLib() As ResLib_t
   Dim ResOImp() As ResOImp_t
   Dim j As Integer
   Dim Col As Integer
   Dim Rc As Integer
   Dim SubTotLib(NCOLS) As Double
   Dim SubTotMes(NCOLS) As Double
   Dim TipoLib As Integer
   Dim Mes As Integer
   Dim nDocs As Long
   
   lHayActFijo = False
   Grid.Redraw = False
      
   Where = SqlYearLng("FEmision") & " = " & Val(Cb_Ano)
   
   If ItemData(Cb_TipoLib) > 0 Then
      
      If ItemData(Cb_TipoLib) = T_COMPRASVENTAS Then
         Where = Where & " AND Documento.TipoLib IN (" & LIB_COMPRAS & ", " & LIB_VENTAS & ")"
      Else
         Where = Where & " AND Documento.TipoLib = " & ItemData(Cb_TipoLib)
      End If
      
   End If
   
   If ItemData(Cb_Mes) > 0 Then
      Where = Where & " AND " & SqlMonthLng("FEmision") & " = " & ItemData(Cb_Mes)
   End If

         
   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows - 1
   CurReg = ""
      
   Rc = GenResLibros(Where, ResLib, ResOImp, False)     'ResOImp no se usa en esta función
   Mes = 0
   TipoLib = 0
   
   If Rc = True Then
   
      For j = 0 To UBound(ResLib)
      
         If i >= Grid.FixedRows Then
            For Col = C_COUNT To C_OIMPDTE
               Total(Col) = Total(Col) + vFmt(Grid.TextMatrix(i, Col))
               SubTotLib(Col) = SubTotLib(Col) + vFmt(Grid.TextMatrix(i, Col))
               SubTotMes(Col) = SubTotMes(Col) + vFmt(Grid.TextMatrix(i, Col))
            Next Col
         End If
                 
         If TipoLib <> ResLib(j).TipoLib Or Mes <> ResLib(j).Mes Then
         
            If TipoLib > 0 Then
               'ponemos subtotal libro
               Grid.rows = Grid.rows + 1
               i = i + 1
               Grid.TextMatrix(i, C_TIPODOC) = "SubTotal " & ReplaceStr(gTipoLib(TipoLib), "Libro de ", "")
               Call FGrSetRowStyle(Grid, i, "B")
               Grid.TextMatrix(i, C_FMT) = "B"
               Grid.TextMatrix(i, C_OBLIGATORIA) = "1"
               For Col = C_COUNT To C_OIMPDTE
                  Grid.TextMatrix(i, Col) = Format(SubTotLib(Col), NEGNUMFMT)
                  SubTotLib(Col) = 0
               Next Col
               
               'se cuentan los documentos porque si se suman en la lista hay docs que aparecen duplicados cuando tinen activo fijo y sin activo fijo
               If TipoLib = LIB_COMPRAS Then
                  nDocs = GetNDocs(TipoLib, Mes)
                  Grid.TextMatrix(i, C_COUNT) = Format(nDocs, NUMFMT)
               End If
               
            End If
            
            TipoLib = ResLib(j).TipoLib
            
         End If
              
         If Mes <> ResLib(j).Mes Then
         
            If Mes > 0 Then
               'ponemos subtotal Mes
               Grid.rows = Grid.rows + 1
               i = i + 1
               Call FGrSetRowStyle(Grid, i, "B")
               Call FGrSetRowStyle(Grid, i, "BC", vbButtonFace)
               Grid.TextMatrix(i, C_FMT) = "B"
               Grid.TextMatrix(i, C_OBLIGATORIA) = "1"
               If SubTotMes(C_IVADEB) > SubTotMes(C_IVACRED) Then
                  Grid.TextMatrix(i, C_IVADEB) = Format(SubTotMes(C_IVADEB) - SubTotMes(C_IVACRED), NEGNUMFMT)
                  Grid.TextMatrix(i, C_TIPODOC) = "IVA Débito de " & gNomMes(Mes)
               ElseIf SubTotMes(C_IVADEB) < SubTotMes(C_IVACRED) Then
                  Grid.TextMatrix(i, C_IVACRED) = Format(SubTotMes(C_IVACRED) - SubTotMes(C_IVADEB), NEGNUMFMT)
                  Grid.TextMatrix(i, C_TIPODOC) = "IVA Crédito de " & gNomMes(Mes)
               Else   'iguales
                  Grid.TextMatrix(i, C_IVADEB) = "0"
                  Grid.TextMatrix(i, C_IVACRED) = "0"
                  Grid.TextMatrix(i, C_TIPODOC) = "IVA de " & gNomMes(Mes)
               End If
               For Col = C_COUNT To C_OIMPDTE
                  SubTotMes(Col) = 0
               Next Col
               Grid.rows = Grid.rows + 1
               i = i + 1
               Grid.TextMatrix(i, C_OBLIGATORIA) = "1"
               Grid.TextMatrix(i, C_FMT) = "L"
            End If
            
            Mes = ResLib(j).Mes
            
         End If
              
              
         Grid.rows = Grid.rows + 1
         i = i + 1
         
         Grid.TextMatrix(i, C_OBLIGATORIA) = "1"
         Grid.TextMatrix(i, C_MES) = Left(gNomMes(ResLib(j).Mes), 3)
         Grid.TextMatrix(i, C_LNGMES) = ResLib(j).Mes
         Grid.TextMatrix(i, C_IDTIPOLIB) = ResLib(j).TipoLib
         Grid.TextMatrix(i, C_TIPOLIB) = ReplaceStr(gTipoLib(ResLib(j).TipoLib), "Libro de ", "")
         Grid.TextMatrix(i, C_TIPODOC) = GetNombreTipoDoc(ResLib(j).TipoLib, ResLib(j).TipoDoc)
         Grid.TextMatrix(i, C_TIPODOCDIM) = GetDiminutivoDoc(ResLib(j).TipoLib, ResLib(j).TipoDoc)
         If ResLib(j).EsSupermercado Then
            Grid.TextMatrix(i, C_TIPODOC) = "Factura Sup. o Com. Sim."
         End If

         If ResLib(j).Giro = 0 Then
            Grid.TextMatrix(i, C_CLASIF) = "No Giro"
         End If
         
         If ResLib(j).FacCompraRetParcial <> 0 Then
            If Grid.TextMatrix(i, C_CLASIF) <> "" Then
               Grid.TextMatrix(i, C_CLASIF) = Grid.TextMatrix(i, C_CLASIF) & "-" & "Ret. Parc."
            Else
               Grid.TextMatrix(i, C_CLASIF) = "Ret. Parc."
            End If
         End If
         
         If ResLib(j).IVAIrrec = IVAIRREC_PARCIAL Then
            If Grid.TextMatrix(i, C_CLASIF) <> "" Then
               Grid.TextMatrix(i, C_CLASIF) = Grid.TextMatrix(i, C_CLASIF) & "-" & "IVA Irrec."
            Else
               Grid.TextMatrix(i, C_CLASIF) = "IVA Irrec."
            End If
            
         ElseIf ResLib(j).IVAIrrec = IVAIRREC_TOTAL Then
            If Grid.TextMatrix(i, C_CLASIF) <> "" Then
               Grid.TextMatrix(i, C_CLASIF) = Grid.TextMatrix(i, C_CLASIF) & "-" & "IVA Irrec. Tot."
            Else
               Grid.TextMatrix(i, C_CLASIF) = "IVA Irrec. Tot."
            End If
            
         End If
         
         If ResLib(j).ActFijo <> 0 Then
            If Grid.TextMatrix(i, C_CLASIF) <> "" Then
               Grid.TextMatrix(i, C_CLASIF) = Grid.TextMatrix(i, C_CLASIF) & " - " & "Act. Fijo"
            Else
               Grid.TextMatrix(i, C_CLASIF) = "Act. Fijo"
               lHayActFijo = True
            End If
            
         ElseIf ResLib(j).TipoReten = TR_HONORARIOS Then
            Grid.TextMatrix(i, C_CLASIF) = "Honorarios"
            
         ElseIf ResLib(j).TipoReten = TR_DIETA Then
            Grid.TextMatrix(i, C_CLASIF) = "Dieta"
            
         ElseIf ResLib(j).TipoReten = TR_OTRO Then
            Grid.TextMatrix(i, C_CLASIF) = "Otro"
            
         End If
                               
         If ResLib(j).TipoLib <> LIB_VENTAS Or ResLib(j).Giro <> 0 Then   'compras, ventas del giro y retenciones
            
            If ResLib(j).IVAIrrec = IVAIRREC_TOTAL Then
               Grid.TextMatrix(i, C_COUNT) = Format(vFmt(Grid.TextMatrix(i, C_COUNT)) + ResLib(j).CountIVAIrrec, BL_NUMFMT)
            ElseIf ResLib(j).CountRetParcial > 0 Then
               Grid.TextMatrix(i, C_COUNT) = Format(vFmt(Grid.TextMatrix(i, C_COUNT)) + ResLib(j).CountRetParcial, BL_NUMFMT)
            Else
               Grid.TextMatrix(i, C_COUNT) = Format(vFmt(Grid.TextMatrix(i, C_COUNT)) + ResLib(j).CountTot, BL_NUMFMT)
            End If
            
            Grid.TextMatrix(i, C_EXENTO) = Format(vFmt(Grid.TextMatrix(i, C_EXENTO)) + ResLib(j).Exento, NEGBL_NUMFMT)
            If ResLib(j).CountRetParcial > 0 Then
               Grid.TextMatrix(i, C_AFECTO) = Format(vFmt(Grid.TextMatrix(i, C_AFECTO)) + ResLib(j).NetoRetParcial, NEGBL_NUMFMT)
            Else
               Grid.TextMatrix(i, C_AFECTO) = Format(vFmt(Grid.TextMatrix(i, C_AFECTO)) + ResLib(j).Afecto, NEGBL_NUMFMT)
            End If
         
         
         Else    'ventas no del giro
            Grid.TextMatrix(i, C_COUNT) = Format(vFmt(Grid.TextMatrix(i, C_COUNT)) + ResLib(j).CountTotNoGiro, BL_NUMFMT)
            
            Grid.TextMatrix(i, C_EXENTO) = Format(vFmt(Grid.TextMatrix(i, C_EXENTO)) + ResLib(j).ExentoNoGiro, NEGBL_NUMFMT)
            Grid.TextMatrix(i, C_AFECTO) = Format(vFmt(Grid.TextMatrix(i, C_AFECTO)) + ResLib(j).AfectoNoGiro, NEGBL_NUMFMT)
         End If
         
         If ResLib(j).TipoLib = LIB_VENTAS Then
            If ResLib(j).Giro <> 0 Then
               Grid.TextMatrix(i, C_IVADEB) = Format(vFmt(Grid.TextMatrix(i, C_IVADEB)) + ResLib(j).IVA, NEGBL_NUMFMT)
            Else
               Grid.TextMatrix(i, C_IVADEB) = Format(vFmt(Grid.TextMatrix(i, C_IVADEB)) + ResLib(j).IVANoGiro, NEGBL_NUMFMT)
            End If
            
         ElseIf ResLib(j).TipoLib = LIB_COMPRAS Then
            Grid.TextMatrix(i, C_IVACRED) = Format(vFmt(Grid.TextMatrix(i, C_IVACRED)) + ResLib(j).IVA, NEGBL_NUMFMT)
         End If
         
         Grid.TextMatrix(i, C_OTROIMP) = Format(vFmt(Grid.TextMatrix(i, C_OTROIMP)) + ResLib(j).OtroImp - Round(ResLib(j).NetoIVAIrrec * gIVA) - ResLib(j).IVARetenido, NEGBL_NUMFMT)
         Grid.TextMatrix(i, C_IVAIRREC) = Format(vFmt(Grid.TextMatrix(i, C_IVAIRREC)) + Round(ResLib(j).NetoIVAIrrec * gIVA), NEGBL_NUMFMT)
         Grid.TextMatrix(i, C_IVARET) = Format(vFmt(Grid.TextMatrix(i, C_IVARET)) + ResLib(j).IVARetenido, NEGBL_NUMFMT)
         
         Grid.TextMatrix(i, C_COUNTDTE) = Format(vFmt(Grid.TextMatrix(i, C_COUNTDTE)) + ResLib(j).CountDTE, NEGBL_NUMFMT)
         Grid.TextMatrix(i, C_IVADTE) = Format(vFmt(Grid.TextMatrix(i, C_IVADTE)) + ResLib(j).IVADTE, NEGBL_NUMFMT)
         Grid.TextMatrix(i, C_OIMPDTE) = Format(vFmt(Grid.TextMatrix(i, C_OIMPDTE)) + ResLib(j).OImpDTE, NEGBL_NUMFMT)
         
         If gTipoDoc(GetTipoDoc(ResLib(j).TipoLib, ResLib(j).TipoDoc)).EsRebaja Then
            
            Grid.TextMatrix(i, C_EXENTO) = Format(vFmt(Grid.TextMatrix(i, C_EXENTO)) * -1, NEGBL_NUMFMT)
            Grid.TextMatrix(i, C_AFECTO) = Format(vFmt(Grid.TextMatrix(i, C_AFECTO)) * -1, NEGBL_NUMFMT)
            If ResLib(j).TipoLib = LIB_VENTAS Then
               If ResLib(j).Giro <> 0 Then
                  Grid.TextMatrix(i, C_IVADEB) = Format(vFmt(Grid.TextMatrix(i, C_IVADEB)) * -1, NEGBL_NUMFMT)    'Para IVANoGiro ya viene en negativo si es rebaja, por la exportación a IVA Estándar
               End If
            ElseIf ResLib(j).TipoLib = LIB_COMPRAS Then
               Grid.TextMatrix(i, C_IVACRED) = Format(vFmt(Grid.TextMatrix(i, C_IVACRED)) * -1, NEGBL_NUMFMT)
            End If
            
            'éstos ya vienen negativos, para la exportación a IVA Estándar
            'Grid.TextMatrix(i, C_OTROIMP) = Format(vFmt(Grid.TextMatrix(i, C_OTROIMP)) * -1, NEGBL_NUMFMT)
            'Grid.TextMatrix(i, C_IVADTE) = Format(vFmt(Grid.TextMatrix(i, C_IVADTE)) * -1, NEGBL_NUMFMT)
            'Grid.TextMatrix(i, C_OIMPDTE) = Format(vFmt(Grid.TextMatrix(i, C_OIMPDTE)) * -1, NEGBL_NUMFMT)
         
         End If
        
      Next j
      
   End If
   
   'sumamos el último total
   If i >= Grid.FixedRows Then
      For Col = C_COUNT To C_OIMPDTE
         Total(Col) = Total(Col) + vFmt(Grid.TextMatrix(i, Col))
         SubTotLib(Col) = SubTotLib(Col) + vFmt(Grid.TextMatrix(i, Col))
         SubTotMes(Col) = SubTotMes(Col) + vFmt(Grid.TextMatrix(i, Col))
      Next Col
   End If
   
   If TipoLib > 0 Then
      'ponemos subtotal libro
      Grid.rows = Grid.rows + 1
      i = i + 1
      Grid.TextMatrix(i, C_OBLIGATORIA) = "1"
      Grid.TextMatrix(i, C_TIPODOC) = "SubTotal " & ReplaceStr(gTipoLib(TipoLib), "Libro de ", "")
      Call FGrSetRowStyle(Grid, i, "B")
      Grid.TextMatrix(i, C_FMT) = "B"
      For Col = C_COUNT To C_OIMPDTE
         Grid.TextMatrix(i, Col) = Format(SubTotLib(Col), NEGNUMFMT)
         SubTotLib(Col) = 0
      Next Col
      
      'se cuentan los documentos porque si se suman en la lista hay docs que aparecen duplicados cuando tinen activo fijo y sin activo fijo
      If TipoLib = LIB_COMPRAS Then
         nDocs = GetNDocs(TipoLib, Mes)
         Grid.TextMatrix(i, C_COUNT) = Format(nDocs, NUMFMT)
      End If
      
   End If
           
   If Mes > 0 Then
      'ponemos subtotal Mes
      Grid.rows = Grid.rows + 1
      i = i + 1
      Grid.TextMatrix(i, C_OBLIGATORIA) = "1"
      Call FGrSetRowStyle(Grid, i, "B")
      Call FGrSetRowStyle(Grid, i, "BC", vbButtonFace)
      Grid.TextMatrix(i, C_FMT) = "B"
      If SubTotMes(C_IVADEB) > SubTotMes(C_IVACRED) Then
         Grid.TextMatrix(i, C_IVADEB) = Format(SubTotMes(C_IVADEB) - SubTotMes(C_IVACRED), NEGNUMFMT)
         Grid.TextMatrix(i, C_TIPODOC) = "IVA Débito de " & gNomMes(Mes)
      ElseIf SubTotMes(C_IVADEB) < SubTotMes(C_IVACRED) Then
         Grid.TextMatrix(i, C_IVACRED) = Format(SubTotMes(C_IVACRED) - SubTotMes(C_IVADEB), NEGNUMFMT)
         Grid.TextMatrix(i, C_TIPODOC) = "IVA Crédito de " & gNomMes(Mes)
      Else   'iguales
         Grid.TextMatrix(i, C_IVADEB) = "0"
         Grid.TextMatrix(i, C_IVACRED) = "0"
         Grid.TextMatrix(i, C_TIPODOC) = "IVA de " & gNomMes(Mes)
      End If
   End If
                    
   GridTot(0).TextMatrix(0, C_TIPODOC) = "TOTAL"
   If ItemData(Cb_TipoLib) > 0 Then
      For Col = C_COUNT To C_OIMPDTE
         GridTot(0).TextMatrix(0, Col) = Format(Total(Col), NEGNUMFMT)
      Next Col
      
   Else
'      GridTot.TextMatrix(0, C_TIPOLIB) = ""
      For Col = C_COUNT To C_OIMPDTE
         GridTot(0).TextMatrix(0, Col) = ""
      Next Col
      For Col = C_IVADEB To C_IVACRED
         GridTot(0).TextMatrix(0, Col) = Format(Total(Col), NEGNUMFMT)
      Next Col
      
   End If
   
   If Total(C_IVADEB) > Total(C_IVACRED) Then
      GridTot(1).TextMatrix(0, C_IVADEB) = Format(Total(C_IVADEB) - Total(C_IVACRED), NEGNUMFMT)
      GridTot(1).TextMatrix(0, C_IVACRED) = ""
      GridTot(1).TextMatrix(0, C_TIPODOC) = "IVA Débito del Año"
   ElseIf Total(C_IVADEB) < Total(C_IVACRED) Then
      GridTot(1).TextMatrix(0, C_IVACRED) = Format(Total(C_IVACRED) - Total(C_IVADEB), NEGNUMFMT)
      GridTot(1).TextMatrix(0, C_IVADEB) = ""
      GridTot(1).TextMatrix(0, C_TIPODOC) = "IVA Crédito del Año"
   Else   'iguales
      GridTot(1).TextMatrix(0, C_IVADEB) = "0"
      GridTot(1).TextMatrix(0, C_IVACRED) = "0"
      GridTot(1).TextMatrix(0, C_TIPODOC) = "IVA del Año"
   End If
   
   
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
   'Marco la columna Ordenada
   'Grid.Row = 0
   'Grid.Col = lOrdenSel
   'Set Grid.CellPicture = FrmMain.Pc_Flecha

   Call FGrSelRow(Grid, Grid.FixedRows)
      
   Grid.Redraw = True
   Call EnableFrm(False)
   
   If Grid.visible And lHayActFijo Then
      MsgBox1 "Atención: Una factura puede aparecer duplicada en el libro de compras cuando tiene una parte de Activo Fijo y otra no.", vbInformation
   End If

End Sub
Private Sub LoadAllCod()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Where As String
   Dim i As Integer
   Dim CurReg As String
   Dim NewReg As String
   Dim Total(NCOLS) As Double
   Dim ResLibCod() As ResLibCod_t
   Dim j As Integer
   Dim AnoMes As Long
   
   GridCod.Redraw = False
      
   Where = SqlYearLng("FEmision") & " = " & Val(Cb_Ano)
   
   If ItemData(Cb_TipoLib) > 0 Then
      Where = Where & " AND Documento.TipoLib = " & ItemData(Cb_TipoLib)
   End If
   
   If ItemData(Cb_Mes) > 0 Then
      Where = Where & " AND " & SqlMonthLng("FEmision") & " = " & ItemData(Cb_Mes)
   End If

   AnoMes = DateSerial(gEmpresa.Ano, CbItemData(Cb_Mes), 1)
         
   GridCod.rows = GridCod.FixedRows
   i = GridCod.FixedRows - 1
   
   If GenExportF29(Where, ResLibCod, AnoMes) = True Then
   
      For j = 0 To UBound(ResLibCod)
      
         If ResLibCod(j).CodF29 = 0 Then
            Exit For
         End If
      
         GridCod.rows = GridCod.rows + 1
         i = i + 1
         
         GridCod.TextMatrix(i, CC_CODF29) = ResLibCod(j).CodF29
         GridCod.TextMatrix(i, CC_VALOR) = Format(ResLibCod(j).Valor, NUMFMT)
         
      Next j
   End If
   
   Call FGrVRows(GridCod)
   GridCod.TopRow = GridCod.FixedRows
   GridCod.Redraw = True
   
End Sub


Public Sub FView(Optional ByVal Mes As Integer = 0)
   lMes = Mes
   
   If lMes = 0 Then
      Me.Show vbModeless
   Else
      Me.Show vbModal
   End If
   
End Sub
Private Sub EnableFrm(bool As Boolean)

   Bt_Search.Enabled = bool
   Bt_DetDoc.Enabled = Not bool
'   bt_Print.Enabled = Not bool
'   Bt_Preview.Enabled = Not bool
'   Bt_CopyExcel.Enabled = Not bool
   
End Sub

Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid)
   
   Set Frm = Nothing

End Sub
Private Sub Bt_Print_Click()
   Dim PrtOrient As Integer
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   PrtOrient = Printer.Orientation
   
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   
   Call gPrtReportes.PrtFlexGrid(Printer)
   
   Me.MousePointer = vbDefault
   
   Printer.Orientation = PrtOrient
   gPrtReportes.FmtCol = -1
   Call ResetPrtBas(gPrtReportes)

End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim PrtOrient As Integer
   Dim Pag As Integer
      
   PrtOrient = Printer.Orientation
      
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de seleccionar la vista previa.", vbExclamation
      Exit Sub
   End If
   
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Pag = gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   If Pag >= 0 Then
      Call Frm.FView(Caption)
   End If
   
   Set Frm = Nothing
   
   Printer.Orientation = PrtOrient
   gPrtReportes.FmtCol = -1
   Call ResetPrtBas(gPrtReportes)
   
End Sub

Private Sub Bt_CopyExcel_Click()
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de copiar.", vbExclamation
      Exit Sub
   End If
   
   Call FGr2Clip(Grid, Me.Caption)
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
   Dim i As Integer, j As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS * 3)
   Dim Titulos(0) As String
   
   Printer.Orientation = ORIENT_HOR
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Me.Caption
   gPrtReportes.Titulos = Titulos
            
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   ColWi(C_TIPOLIB) = ColWi(C_TIPOLIB) - 100 '- 400
   ColWi(C_TIPODOC) = ColWi(C_TIPODOC) - 100 '- 800
   ColWi(C_COUNT) = ColWi(C_COUNT) - 100 '- 200
   ColWi(C_COUNTDTE) = ColWi(C_COUNTDTE) - 100 '- 230
   ColWi(C_OIMPDTE) = ColWi(C_OIMPDTE) - 100 '- 70
   
   For i = 0 To Grid.Cols - 1
      If ColWi(i) > 0 Then
         ColWi(i) = ColWi(i) * 0.95
      End If
   Next i
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_OBLIGATORIA
   gPrtReportes.FmtCol = C_FMT
   
   gPrtReportes.GrFontSize = 7
   gPrtReportes.GrFontName = "Arial"
   'gPrtReportes.TotFntBold = False
   
   
   j = 0
   For i = 0 To Grid.Cols - 1
      Total(j) = GridTot(0).TextMatrix(0, i)
      j = j + 1
   Next i
   
   For i = 0 To Grid.Cols - 1
      Total(j) = GridTot(1).TextMatrix(0, i)
      j = j + 1
   Next i
   
   gPrtReportes.Total = Total
   gPrtReportes.NTotLines = 2
   
   
End Sub

Private Sub Grid_DblClick()
   Call PostClick(Bt_DetDoc)
End Sub

Private Sub Grid_Scroll()
   GridTot(0).LeftCol = Grid.LeftCol
   GridTot(1).LeftCol = Grid.LeftCol
End Sub

Private Sub GridCod_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(GridCod, "Códigos F29 - " & gEmpresa.NombreCorto & " - " & Cb_Mes & " " & Cb_Ano)
   End If
   
End Sub
Private Sub Bt_Resumen_Click()
   Dim Frm As FrmResIVA
   
   Set Frm = New FrmResIVA
   
   Me.MousePointer = vbHourglass
   
   If ItemData(Cb_Mes) > 0 Then
      Call Frm.FView(ItemData(Cb_Mes), Val(Cb_Ano), ItemData(Cb_TipoLib), False)
   Else
      Call Frm.FView(0, Val(Cb_Ano), ItemData(Cb_TipoLib), False)
   End If
   
   Me.MousePointer = vbDefault
   
   Set Frm = Nothing
   
End Sub
Private Function GetNDocs(ByVal TipoLib As Integer, Mes As Integer)
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Count(*) FROM Documento WHERE TipoLib = " & TipoLib & " AND " & SqlYearLng("FEmision") & " = " & Val(Cb_Ano) & " AND " & SqlMonthLng("FEmision") & " = " & Mes & " AND Estado <> " & ED_ANULADO
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetNDocs = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)

End Function
