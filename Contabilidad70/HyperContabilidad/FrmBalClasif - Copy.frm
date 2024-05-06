VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmBalClasif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Clasificado"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "FrmBalClasif.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fr_MultiSelect 
      BorderStyle     =   0  'None
      Height          =   3030
      Left            =   5400
      TabIndex        =   32
      Top             =   1980
      Width           =   3735
      Begin TabDlg.SSTab Tab_MultiSelect 
         Height          =   2895
         Left            =   0
         TabIndex        =   33
         Top             =   120
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   5106
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Área de Negocio"
         TabPicture(0)   =   "FrmBalClasif.frx":000C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Bt_ClearAreaNeg"
         Tab(0).Control(1)=   "Ls_AreaNeg"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Centro Gestión"
         TabPicture(1)   =   "FrmBalClasif.frx":0028
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Ls_CCosto"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Bt_ClearCCosto"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         Begin VB.CommandButton Bt_ClearAreaNeg 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   -71820
            Picture         =   "FrmBalClasif.frx":0044
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "De-seleccionar todo"
            Top             =   420
            Width           =   400
         End
         Begin VB.CommandButton Bt_ClearCCosto 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3180
            Picture         =   "FrmBalClasif.frx":0498
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "De-seleccionar todo"
            Top             =   420
            Width           =   400
         End
         Begin VB.ListBox Ls_CCosto 
            Height          =   2310
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   35
            Top             =   420
            Width           =   3015
         End
         Begin VB.ListBox Ls_AreaNeg 
            Height          =   2310
            Left            =   -74880
            Style           =   1  'Checkbox
            TabIndex        =   34
            Top             =   420
            Width           =   3015
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6075
      Left            =   60
      TabIndex        =   12
      Top             =   2100
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   10716
      _Version        =   393216
      Rows            =   25
      Cols            =   9
      FixedCols       =   0
      AllowUserResizing=   2
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   60
      TabIndex        =   24
      Top             =   660
      Width           =   10095
      Begin VB.CommandButton Bt_MultiSelect 
         Caption         =   "Selección múltiple"
         Height          =   315
         Left            =   6900
         TabIndex        =   10
         Top             =   900
         Width           =   1695
      End
      Begin VB.ComboBox Cb_TipoAjuste 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   540
         Width           =   1275
      End
      Begin VB.ComboBox Cb_AreaNeg 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   2955
      End
      Begin VB.ComboBox Cb_CCosto 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   540
         Width           =   2955
      End
      Begin VB.ComboBox Cb_Nivel 
         Height          =   315
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   2100
         Picture         =   "FrmBalClasif.frx":08EC
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   4080
         Picture         =   "FrmBalClasif.frx":0BF6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   230
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   3060
         TabIndex        =   2
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Buscar 
         Caption         =   "&Listar"
         Height          =   615
         Left            =   8760
         Picture         =   "FrmBalClasif.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   1155
      End
      Begin VB.CheckBox Ch_LibOficial 
         Caption         =   "Libro Oficial "
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Sólo comprobantes Aprobados (no Pendientes)"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox Ch_VerSubTot 
         Caption         =   "Ver Sub-totales"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Ajuste:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área Negocio:"
         Height          =   195
         Index           =   2
         Left            =   4500
         TabIndex        =   30
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro Gestión:"
         Height          =   195
         Index           =   1
         Left            =   4500
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel cuentas"
         Height          =   195
         Index           =   0
         Left            =   2460
         TabIndex        =   27
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   7
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   23
      Top             =   0
      Width           =   10095
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
         Left            =   2940
         Picture         =   "FrmBalClasif.frx":133E
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Calculadora"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_VerLibMayor 
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
         Picture         =   "FrmBalClasif.frx":169F
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ir a Libro Mayor para cuenta seleccionada"
         Top             =   180
         Width           =   375
      End
      Begin VB.CheckBox Ch_VerCodCuenta 
         Caption         =   "Ver Código Cuenta"
         Height          =   195
         Left            =   5040
         TabIndex        =   21
         Top             =   240
         Width           =   1815
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
         Left            =   1980
         Picture         =   "FrmBalClasif.frx":1A0D
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Sumar movimientos seleccionados"
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
         Left            =   2520
         Picture         =   "FrmBalClasif.frx":1AB1
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   3360
         Picture         =   "FrmBalClasif.frx":1E4F
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   1440
         Picture         =   "FrmBalClasif.frx":2278
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Left            =   600
         Picture         =   "FrmBalClasif.frx":26BD
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   8760
         TabIndex        =   22
         Top             =   180
         Width           =   1155
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
         Left            =   1020
         Picture         =   "FrmBalClasif.frx":2B64
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   60
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   8220
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmBalClasif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CODIGO = 0
Const C_CUENTA = 1
Const C_VALOR = 2
Const C_NIVEL = 3
Const C_IDCUENTA = 4
Const C_CLASCTA = 5
Const C_DEBITOS = 6
Const C_CREDITOS = 7
Const C_FMT = 8

Const C_INI_MES = C_FMT + 1
Const C_SALDOFIN = C_INI_MES + 12
Const C_INI_COMP = C_SALDOFIN + 1

Const TOT_CUENTA = -1

Const COLWI_CUENTA = 6230 + 460

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Dim lMes As Integer
Dim lCaption As String

Dim lClasCta As String
Dim lMensual As Boolean
Dim lComparativo As Boolean
Dim lBalClasif As Boolean
Dim lResEje(MAX_NIVELES) As Cuenta_t
Dim lWCodCta As Integer
Dim lWVal As Integer

Dim lLibOf As Integer

Private Sub Bt_Buscar_Click()
   Dim F1 As Long, F2 As Long
   
   Fr_MultiSelect.Visible = False
   
   F1 = GetTxDate(Tx_Desde)
   F2 = GetTxDate(Tx_Hasta)
      
   If F1 > F2 Then
      MsgBox1 "Fecha de inicio es posterior a la fecha de término del reporte.", vbExclamation
      Tx_Hasta.SetFocus
      Exit Sub
      
   End If
   
   If Year(F1) <> gEmpresa.Ano Then
      MsgBox1 "La fecha de inicio no corresponde al periodo actual.", vbExclamation
      Exit Sub
   End If
   If Year(F2) <> gEmpresa.Ano Then
      MsgBox1 "La fecha de término no corresponde al periodo actual.", vbExclamation
      Exit Sub
   End If
   
   MousePointer = vbHourglass
   Call LoadAll
   
   MousePointer = vbDefault
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_ClearAreaNeg_Click()
   Call cbClearSel(Ls_AreaNeg)
   Cb_AreaNeg.Enabled = True
End Sub

Private Sub Bt_ClearCCosto_Click()
   Call cbClearSel(Ls_CCosto)
   Cb_CCosto.Enabled = True
End Sub

Private Sub Bt_CopyExcel_Click()
   
   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de copiar.", vbExclamation
      Exit Sub
   End If
   
   Call LP_FGr2Clip(Grid, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta)
End Sub

Private Sub Bt_MultiSelect_Click()
   Fr_MultiSelect.Visible = Not Fr_MultiSelect.Visible

End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim Pag As Integer
   
   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de seleccionar la vista previa.", vbExclamation
      Exit Sub
   End If
   
'   If lMensual And Not lComparativo Then
'      If MsgBox1("Es muy probable que al imprimir este informe no quepan todas las columnas." & vbLf & vbLf & "Se sugiere copiarlo a Excel e imprimirlo desde ahí.", vbInformation + vbOKCancel) = vbCancel Then
'         Exit Sub
'      End If
'   End If
      
   
   lPapelFoliado = False
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   gPrtLibros.CallEndDoc = False
   
   If lMensual Then
      gPrtLibros.PermitirMasDe1Franja = True
   End If
      
   Pag = gPrtLibros.PrtFlexGrid(Frm)
   Call PrtPieBalance(Frm, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight)
   
   gPrtLibros.CallEndDoc = True
   gPrtLibros.PermitirMasDe1Franja = False
         
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   Call ResetPrtBas(gPrtLibros)
   
End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientacion As Integer
   Dim Frm As FrmPrtSetup
   Dim Fecha As Long
   Dim Usuario As String
   Dim FDesde As Long
   Dim FHasta As Long
   Dim Pag As Integer
   
   lPapelFoliado = False
   
   If Bt_Buscar.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   If lMensual And Not lComparativo Then
      If MsgBox1("Es muy probable que al imprimir este informe no quepan todas las columnas." & vbLf & vbLf & "Se sugiere copiarlo a Excel e imprimirlo desde ahí.", vbInformation + vbOKCancel) = vbCancel Then
         Exit Sub
      End If
   End If
   
   If Ch_LibOficial.Visible = True And Ch_LibOficial <> 0 Then
   
      If QryLogImpreso(lLibOf, 0, FDesde, FHasta, Fecha, Usuario) = True Then
         If MsgBox1("El " & gLibroOficial(lLibOf) & " Oficial ya ha sido impreso en papel foliado el día " & Format(Fecha, DATEFMT) & " por el usuario " & Usuario & ", para el período comprendido entre el " & Format(FDesde, DATEFMT) & " y el " & Format(FHasta, DATEFMT) & "." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
         End If
      End If
         
      lPapelFoliado = True
   End If
      
   Set Frm = New FrmPrtSetup
   If Frm.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar) = vbOK Then
      OldOrientacion = Printer.Orientation
      
      Me.MousePointer = vbHourglass
      
      Call SetUpPrtGrid
      
      gPrtLibros.CallEndDoc = False
      
      If lMensual Then
         gPrtLibros.PermitirMasDe1Franja = True
      End If
      
      Pag = gPrtLibros.PrtFlexGrid(Printer)
      
      Call PrtPieBalance(Printer, Pag, gPrtLibros.GrLeft, gPrtLibros.GrRight)
      
      gPrtLibros.CallEndDoc = True
      gPrtLibros.PermitirMasDe1Franja = False
      
      Me.MousePointer = vbDefault
      
      If lPapelFoliado And Ch_LibOficial.Visible = True And Ch_LibOficial <> 0 Then
         Call AppendLogImpreso(lLibOf, 0, GetTxDate(Tx_Desde), GetTxDate(Tx_Hasta))
      End If
      
      'Chequeo si debo actualizar folio ultimo usado
      Call UpdateUltUsado(lPapelFoliado, Pag)
      
      Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
      
   End If
   
   Call ResetPrtBas(gPrtLibros)
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   Dim Mes As Integer
   
   lWCodCta = 1300
   lWVal = G_DVALWIDTH + 200
   
   If Ch_VerCodCuenta <> 0 Then
      Grid.ColWidth(C_CODIGO) = lWCodCta
   Else
      Grid.ColWidth(C_CODIGO) = 0
   End If
   
   Grid.ColWidth(C_CUENTA) = COLWI_CUENTA
   Grid.ColWidth(C_VALOR) = lWVal
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_NIVEL) = 0
   Grid.ColWidth(C_CREDITOS) = 0
   Grid.ColWidth(C_DEBITOS) = 0
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_CLASCTA) = 0

   If lMensual Then
      Grid.Cols = Grid.Cols + 12 + 1    'meses más saldo
      Grid.ColWidth(C_CODIGO) = 1200
      Grid.ColWidth(C_CUENTA) = 3500
      Grid.ColWidth(C_VALOR) = 0
         
      If lComparativo Then
         Grid.Cols = Grid.Cols + 3
         GridTot.Cols = Grid.Cols
         Grid.FixedRows = 2
         
         Grid.ColWidth(C_CUENTA) = 3650
         
         For i = C_INI_COMP To C_INI_COMP + 2
            Grid.ColWidth(i) = G_VALWIDTH + 200
         Next i
         
         Grid.TextMatrix(0, C_INI_COMP) = "Saldo"
         Grid.TextMatrix(0, C_INI_COMP + 1) = "Saldo"
         Grid.TextMatrix(1, C_INI_COMP + 1) = "Período"
         Grid.TextMatrix(0, C_INI_COMP + 2) = "Saldo"
         
         For i = C_INI_MES To C_INI_COMP - 1
            Grid.ColWidth(i) = 0
            GridTot.ColWidth(i) = 0
         Next i
      
      Else
         For i = C_INI_MES To Grid.Cols - 1
            Grid.ColWidth(i) = G_VALWIDTH + 100
            Mes = i - C_INI_MES + 1
            If Mes <= 12 Then
               Grid.TextMatrix(0, i) = gNomMes(i - C_INI_MES + 1)
            End If
         Next i
      
         Grid.TextMatrix(0, C_SALDOFIN) = "Saldo"
         
         Grid.FixedCols = 2
      End If
      
      For i = C_INI_MES To Grid.Cols - 1
         Grid.ColAlignment(i) = flexAlignRightCenter
      Next i
      
   End If
   
   Grid.TextMatrix(0, C_CODIGO) = "Cód. Cuenta"
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.TextMatrix(0, C_VALOR) = "Saldo"

   Grid.ColAlignment(C_CODIGO) = flexAlignLeftCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter

   If lBalClasif Then
      GridTot.Visible = False
   End If
   
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot)
         
   Grid.TextMatrix(0, C_FMT) = "             .FMT"
   
   Call FGrVRows(Grid, 1)
   
End Sub

Private Sub Bt_VerLibMayor_Click()
   Dim Frm As FrmLibMayor
   Dim IdCuenta As Long
   
   IdCuenta = vFmt(Grid.TextMatrix(Grid.Row, C_IDCUENTA))
   
   If IdCuenta > 0 Then
   
      Set Frm = New FrmLibMayor
      Call Frm.FViewChain(GetTxDate(Tx_Desde), GetTxDate(Tx_Hasta), IdCuenta, CbItemData(Cb_TipoAjuste))
      Set Frm = Nothing
   
   End If

End Sub

Private Sub Cb_Nivel_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_TipoAjuste_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_LibOficial_Click()
   
   If Ch_LibOficial <> 0 Then
      Cb_AreaNeg.ListIndex = 0
      Cb_AreaNeg.Enabled = False
      Cb_CCosto.ListIndex = 0
      Cb_CCosto.Enabled = False
      Call cbClearSel(Ls_AreaNeg)
      Call cbClearSel(Ls_CCosto)
      Ls_AreaNeg.Enabled = False
      Ls_CCosto.Enabled = False
      Bt_MultiSelect.Enabled = False
      Fr_MultiSelect.Visible = False
   Else
      Cb_AreaNeg.Enabled = True
      Cb_CCosto.Enabled = True
      Ls_AreaNeg.Enabled = True
      Bt_MultiSelect.Enabled = True
      Ls_CCosto.Enabled = True
   End If
   
   Call EnableFrm(True)

End Sub

Private Sub Ch_VerCodCuenta_Click()
   Dim i As Integer
      
   If Ch_VerCodCuenta <> 0 Then
      Grid.ColWidth(C_CODIGO) = lWCodCta
      
      If lMensual Or lComparativo Then
         Grid.ColWidth(C_CUENTA) = 3300
         For i = C_INI_MES To Grid.Cols - 1
            If Grid.ColWidth(i) > 0 Then
               Grid.ColWidth(i) = lWVal - 100
            End If
         Next i
         
      Else
         Grid.ColWidth(C_CUENTA) = COLWI_CUENTA
         Grid.ColWidth(C_VALOR) = lWVal
      End If
      
   Else
      Grid.ColWidth(C_CODIGO) = 0
      
      If lMensual Or lComparativo Then
         Grid.ColWidth(C_CUENTA) = 3500 + 200
         For i = C_INI_MES To Grid.Cols - 1
            If Grid.ColWidth(i) > 0 Then
               Grid.ColWidth(i) = lWVal + 200
            End If
         Next i
         
      Else
         Grid.ColWidth(C_CUENTA) = COLWI_CUENTA + 900
         Grid.ColWidth(C_VALOR) = lWVal + 400
      End If
   End If
   
   For i = 0 To Grid.Cols - 1
      GridTot.ColWidth(i) = Grid.ColWidth(i)
   Next i
   
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim D1 As Long, D2 As Long
   Dim ActDate As Long
   Dim MesActual As Integer
      
   MesActual = GetMesActual()
   
   If lMes = 0 Then
      If MesActual > 0 Then
         lMes = MesActual
      Else
         lMes = GetUltimoMesConComps()
      End If
   End If
   
   ActDate = DateSerial(gEmpresa.Ano, lMes, 1)
   
   Call FirstLastMonthDay(ActDate, D1, D2)
   Call SetTxDate(Tx_Desde, DateSerial(gEmpresa.Ano, 1, 1))
   Call SetTxDate(Tx_Hasta, D2)
   
   Me.Caption = lCaption
   
   Call SetUpGrid
   
   If gNiveles.nNiveles >= 3 Then
      Call FillNivel(Cb_Nivel, 3)
   Else
      Call FillNivel(Cb_Nivel, gNiveles.nNiveles)
   End If
   
   
   Cb_Nivel.RemoveItem (0)   'eliminamos el nivel 1 porque no tiene sentido
   
   Call FillCbAreaNeg(Cb_AreaNeg)
   Call FillCbAreaNeg(Ls_AreaNeg, False)
   Call FillCbCCosto(Cb_CCosto)
   Call FillCbCCosto(Ls_CCosto, False)
   
   Fr_MultiSelect.Visible = False
   Tab_MultiSelect.Tab = 0
   
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_FINANCIERO), TAJUSTE_FINANCIERO)
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_TRIBUTARIO), TAJUSTE_TRIBUTARIO)
   Call CbSelItem(Cb_TipoAjuste, TAJUSTE_FINANCIERO)

   
   If lMensual Then
      If lComparativo Then
         lOrientacion = ORIENT_VER
      Else
         lOrientacion = ORIENT_HOR
      End If
   Else
      lOrientacion = ORIENT_VER
   End If
   
   Ch_VerCodCuenta = 1
   
   Ch_VerSubTot.Visible = False
   
   If lBalClasif = False Then
      Ch_LibOficial.Visible = False
      Ch_VerSubTot.Visible = True
      Ch_VerSubTot = 1
   End If
   
   Call ReadResEje
   Call LoadAll
   Call SetupPriv
      
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer, j As Integer
   Dim ColWi() As Integer
   Dim Total() As String
   Dim Titulos(1) As String
   Dim Encabezados(3) As String
   Dim FontTit(1) As FontDef_t
   Dim FontEnc(0) As FontDef_t
   Dim Nombres(5) As String
   Dim OldOrient As Integer
   
   Set gPrtLibros.Grid = Grid
   
   Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Titulos(0) = Me.Caption & " " & gEmpresa.Ano
   FontTit(0).FontBold = True
   
   If lInfoPreliminar Then
      Titulos(1) = INFO_PRELIMINAR
      FontTit(1).FontBold = True
   End If
   
   gPrtLibros.Titulos = Titulos
   Call gPrtLibros.FntTitulos(FontTit())
   
   If GetTxDate(Tx_Desde) <> DateSerial(gEmpresa.Ano, 1, 1) Then
      Encabezados(0) = Format(GetTxDate(Tx_Desde), DATEFMT) & " a "
   Else
      Encabezados(0) = "Al "
   End If
   Encabezados(0) = Encabezados(0) & Format(GetTxDate(Tx_Hasta), DATEFMT)
   
   i = 1
   If Cb_AreaNeg.ListIndex > 0 Then
      Encabezados(i) = "Area de Negocio   : " & Cb_AreaNeg
      i = i + 1
   
   ElseIf Ls_AreaNeg.SelCount >= 1 Then
      
      Encabezados(i) = "Area de Negocio   : "
      For j = 0 To Ls_AreaNeg.ListCount - 1
         If Ls_AreaNeg.Selected(i) Then
            Encabezados(i) = Encabezados(i) & Ls_AreaNeg.List(i) & ", "
         End If
      Next j
      
      Encabezados(i) = Left(Encabezados(i), Len(Encabezados(i)) - 2)   'secamos la última coma
      i = i + 1
       
   End If
   
   If Cb_CCosto.ListIndex > 0 Then
      Encabezados(i) = "Centro de Gestión : " & Cb_CCosto
   
   ElseIf Ls_CCosto.SelCount >= 1 Then
      
      Encabezados(i) = "Centro de Gestión : "
      For j = 0 To Ls_CCosto.ListCount - 1
         If Ls_CCosto.Selected(i) Then
            Encabezados(i) = Encabezados(i) & Ls_CCosto.List(i) & ", "
         End If
      Next j
      
      Encabezados(i) = Left(Encabezados(i), Len(Encabezados(i)) - 2)   'secamos la última coma
       
   End If
   
   gPrtLibros.Encabezados = Encabezados
   
   FontEnc(0).FontBold = True
   FontEnc(0).FontName = "Arial"
   FontEnc(0).FontSize = 10
   Call gPrtLibros.FntEncabezados(FontEnc())
   
   gPrtLibros.GrFontName = Grid.FontName
   gPrtLibros.GrFontSize = Grid.FontSize
    
   ReDim ColWi(Grid.Cols - 1)
   ReDim Total(Grid.Cols - 1)
    
   For i = 0 To Grid.Cols - 1
      If Grid.ColWidth(i) > 0 Then
         ColWi(i) = Grid.ColWidth(i) - 100
      End If
   Next i
 
   For i = 0 To Grid.Cols - 1
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
   
   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
   gPrtLibros.ColObligatoria = C_IDCUENTA
   If GridTot.Visible = True Then
      gPrtLibros.NTotLines = 1
   Else
      gPrtLibros.NTotLines = 0
   End If
   gPrtLibros.FmtCol = C_FMT
   
   If Ch_LibOficial <> 0 Then
      gPrtLibros.PrintFecha = False
   End If
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Nivel As Integer
   Dim Rs As Recordset
   Dim Total(MAX_NIVELES) As RepNiv_t
   Dim TotalMes(MAX_NIVELES, 12) As RepNiv_t
   Dim CurNiv As Integer
   Dim CurCta As String
   Dim i As Integer, j As Integer, k As Integer
   Dim Row As Integer
   Dim Diff As Double
   Dim FirstDiaMes As Long, LastDiaMes As Long
   Dim CodPadre As String, NomPadre As String
   Dim WhFecha As String
   Dim TotalFinal As Double
   Dim TotalFinalMes(12) As Double
   Dim Wh As String
   Dim Mes As Integer
   Dim FDesde As Long
   Dim FHasta As Long
   Dim TotMesHasta As Double
   Dim TotMesMenos1 As Double
   Dim TotFinHasta As Double
   Dim TotFinMenos1 As Double
   Dim LinPatrimonio As Integer
   Dim LinResEjercicio(MAX_NIVELES) As Integer
   Dim ClasifPadre As Integer
   Dim TotClasif(MAX_CLASCTA) As Double
   Dim LinTotClasif(MAX_CLASCTA) As Integer
   Dim ResEjercicio As Double
   Dim RowVisible As Boolean
   Dim Col As Integer
   Dim NombreCtaTotal As String
   
   Nivel = Val(Cb_Nivel)
   
   Grid.Redraw = False
   
   FDesde = GetTxDate(Tx_Desde)
   FHasta = GetTxDate(Tx_Hasta)
   WhFecha = "Comprobante.Fecha BETWEEN " & FDesde & " AND " & FHasta
   
   If ItemData(Cb_AreaNeg) > 0 Then
      Wh = Wh & " AND MovComprobante.IdAreaNeg = " & CbItemData(Cb_AreaNeg)
   
   ElseIf Ls_AreaNeg.SelCount > 0 And Ls_AreaNeg.SelCount <> Ls_AreaNeg.ListCount Then
      
      Wh = Wh & " AND MovComprobante.IdAreaNeg IN ( "
      For j = 0 To Ls_AreaNeg.ListCount - 1
         If Ls_AreaNeg.Selected(j) Then
            Wh = Wh & Ls_AreaNeg.ItemData(j) & ", "
         End If
      Next j
      
      Wh = Left(Wh, Len(Wh) - 2) & ")"  'secamos la última coma
          
   End If
   
   If ItemData(Cb_CCosto) > 0 Then
      Wh = Wh & " AND MovComprobante.IdCCosto = " & ItemData(Cb_CCosto)
   
   ElseIf Ls_CCosto.SelCount > 0 And Ls_CCosto.SelCount <> Ls_CCosto.ListCount Then
      
      Wh = Wh & " AND MovComprobante.IdCCosto IN ( "
      For j = 0 To Ls_CCosto.ListCount - 1
         If Ls_CCosto.Selected(j) Then
            Wh = Wh & Ls_CCosto.ItemData(j) & ", "
         End If
      Next j
      
      Wh = Left(Wh, Len(Wh) - 2) & ")"  'secamos la última coma
       
   End If
   
   If ItemData(Cb_TipoAjuste) > 0 Then
      If ItemData(Cb_TipoAjuste) = TAJUSTE_FINANCIERO Then
         Wh = Wh & " AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
      Else
         Wh = Wh & " AND Comprobante.TipoAjuste IN (" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & ")"
      End If
   End If


   Q1 = GenQueryPorNiveles(Nivel, WhFecha & Wh, Ch_LibOficial <> 0, lClasCta, lMensual)

   Set Rs = OpenRs(DbMain, Q1)
   
   For j = 0 To MAX_NIVELES
      Total(j).Debe = 0
      Total(j).Haber = 0
      Total(j).Linea = 0
      For k = 0 To 12
         TotalMes(j, k).Debe = 0
         TotalMes(j, k).Haber = 0
         TotalMes(j, k).Linea = 0
      Next k
   Next j
   
   
   If lMensual And Not lComparativo Then
      For k = 1 To 12
         Mes = C_INI_MES + k - 1
         
         Grid.ColWidth(Mes) = G_VALWIDTH + 100
         GridTot.ColWidth(Mes) = G_VALWIDTH + 100
         Grid.TextMatrix(0, Mes) = gNomMes(k)
         
      Next k
         
   End If

   
   Grid.Rows = Grid.FixedRows
   i = Grid.Rows - 1
   
   CurNiv = 0
   CurCta = ""
   CodPadre = ""
   NomPadre = ""
   
   Do While Rs.EOF = False
   
      'Obtengo el Padre de la cuenta y pongo el total del padre al cambiar a otro
      
      If vFld(Rs("Nivel")) = 1 Then
      
         If CodPadre <> "" Then 'había uno antes
            
            'agregamos la cuenta de resultado del ejercicio
            If lBalClasif And ClasifPadre = CLASCTA_PASIVO Then
               Call AddResEjercicio(i, LinPatrimonio, LinResEjercicio)
            End If
            
            i = i + 1
            Grid.Rows = i + 1
            Call FGrSetRowStyle(Grid, i, "B")
            
            Grid.TextMatrix(i, C_CUENTA) = "TOTAL " & UCase(NomPadre)
            Grid.TextMatrix(i, C_IDCUENTA) = TOT_CUENTA
            Grid.TextMatrix(i, C_CLASCTA) = ClasifPadre
            Grid.TextMatrix(i, C_FMT) = "B"
            
            Grid.TextMatrix(i, C_DEBITOS) = Format(Total(1).Debe, NEGNUMFMT)
            Grid.TextMatrix(i, C_CREDITOS) = Format(Total(1).Haber, NEGNUMFMT)
            
            If ClasifPadre = CLASCTA_RESULTADO Then        'si es Resultado, pueden haber 2 cuentas de primer nivel de este tipo, por lo que el total no se pone a este nivel sino abajo en la grilla de totales
               If Ch_VerSubTot = 0 Then
                  Grid.RowHeight(i) = 0   'ocultamos el total
               End If
'            Else
'               If TotClasif(ClasifPadre) <> 0 Then
'                  MsgBox1 "ATENCIÓN:" & vbCrLf & vbCrLf & "Revise su plan de cuentas. Hay dos o más cuentas de primer nivel clasificadas como " & UCase(gClasCta(ClasifPadre)) & "." & vbCrLf & "Esto puede generar errores en el Balance.", vbExclamation
'               End If
            End If
            
            If ClasifPadre = CLASCTA_ACTIVO Then
               TotClasif(ClasifPadre) = Total(1).Debe - Total(1).Haber
            ElseIf ClasifPadre = CLASCTA_PASIVO Or ClasifPadre = CLASCTA_RESULTADO Then
               TotClasif(ClasifPadre) = Total(1).Haber - Total(1).Debe
            End If
            
            LinTotClasif(ClasifPadre) = i
            
            If lMensual Then
               For k = 1 To 12
                  Mes = C_INI_MES + k - 1
                  If ClasifPadre = CLASCTA_ACTIVO Then
                     Grid.TextMatrix(i, Mes) = Format(TotalMes(1, k).Debe - TotalMes(1, k).Haber, NEGNUMFMT)
                  ElseIf ClasifPadre = CLASCTA_PASIVO Or ClasifPadre = CLASCTA_RESULTADO Then
                     Grid.TextMatrix(i, Mes) = Format(TotalMes(1, k).Haber - TotalMes(1, k).Debe, NEGNUMFMT)
                  End If
               Next k
            End If
            
            'Salto una línea
            i = i + 1
            Grid.Rows = i + 1
            Grid.TextMatrix(i, C_IDCUENTA) = "*"
            
         End If
         
         CodPadre = vFld(Rs("Codigo"))
         NomPadre = FCase(vFld(Rs("Descripcion"), True))
         ClasifPadre = vFld(Rs("Clasificacion"))

      End If
            
      If vFld(Rs("Nivel")) < CurNiv Then    'disminuye el nivel
         'asignamos los totales hacia arriba
         For j = CurNiv - 1 To vFld(Rs("Nivel")) Step -1
            If j > 1 Then
               Grid.TextMatrix(Total(j).Linea, C_DEBITOS) = Format(Total(j).Debe, BL_NUMFMT)
               Grid.TextMatrix(Total(j).Linea, C_CREDITOS) = Format(Total(j).Haber, BL_NUMFMT)
               
               If lMensual Then
                  For k = 1 To 12
                     Mes = C_INI_MES + k - 1
                     If ClasifPadre = CLASCTA_ACTIVO Then
                        Grid.TextMatrix(Total(j).Linea, Mes) = Format(TotalMes(j, k).Debe - TotalMes(j, k).Haber, NEGNUMFMT)
                     ElseIf ClasifPadre = CLASCTA_PASIVO Or ClasifPadre = CLASCTA_RESULTADO Then
                       Grid.TextMatrix(Total(j).Linea, Mes) = Format(TotalMes(j, k).Haber - TotalMes(j, k).Debe, NEGNUMFMT)
                     End If
                  Next k
               End If

            End If
            Total(j).Debe = 0
            Total(j).Haber = 0
            Total(j).Linea = 0
            
            For k = 0 To 12
               TotalMes(j, k).Debe = 0
               TotalMes(j, k).Haber = 0
               TotalMes(j, k).Linea = 0
            Next k
            
         Next j
      End If
         

      If CurCta <> vFld(Rs("Codigo")) Then
      
         If CurCta <> "" And CurNiv > 1 Then
            'ponemos totales de cuenta actual
            Grid.TextMatrix(Total(CurNiv).Linea, C_DEBITOS) = Format(Total(CurNiv).Debe, BL_NUMFMT)
            Grid.TextMatrix(Total(CurNiv).Linea, C_CREDITOS) = Format(Total(CurNiv).Haber, BL_NUMFMT)
            
            If lMensual Then
               For k = 1 To 12
                  Mes = C_INI_MES + k - 1
                  If ClasifPadre = CLASCTA_ACTIVO Then
                     Grid.TextMatrix(Total(CurNiv).Linea, Mes) = Format(TotalMes(CurNiv, k).Debe - TotalMes(CurNiv, k).Haber, NEGNUMFMT)
                  ElseIf ClasifPadre = CLASCTA_PASIVO Or ClasifPadre = CLASCTA_RESULTADO Then
                     Grid.TextMatrix(Total(CurNiv).Linea, Mes) = Format(TotalMes(CurNiv, k).Haber - TotalMes(CurNiv, k).Debe, NEGNUMFMT)
                  End If
               Next k
            End If
         End If
      
         If lBalClasif Then
            
            If LinPatrimonio = 0 Then
               If vFld(Rs("IdCuenta")) = gCtasBas.IdCtaPatrimonio Then
                  LinPatrimonio = i + 1 'se agrega después
               End If
            ElseIf LinResEjercicio(0) = 0 Then
               If vFld(Rs("Nivel")) < CurNiv Then    'disminuye el nivel
                  Call AddResEjercicio(i, LinPatrimonio, LinResEjercicio)
               End If
            End If
            
         End If
                    
         'actualizamos el nivel
         CurNiv = vFld(Rs("Nivel"))
         
         'agregamos la nueva cuenta
         i = i + 1
         Grid.Rows = i + 1
         CurCta = vFld(Rs("Codigo"))
  
         Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("idCuenta"))
         Grid.TextMatrix(i, C_CLASCTA) = vFld(Rs("Clasificacion"))
         Grid.TextMatrix(i, C_NIVEL) = CurNiv
         
         
         If CurNiv <> 1 Then 'Es un hijo
            Call FGrSetRowStyle(Grid, i, "FC", gColores(CurNiv))
            Grid.TextMatrix(i, C_FMT) = "C" & gColores(CurNiv)
            Grid.TextMatrix(i, C_CODIGO) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
            Grid.TextMatrix(i, C_CUENTA) = String(REP_INDENT * (CurNiv - 2), " ") & FCase(vFld(Rs("Descripcion")))
                        
         Else 'Es el Padre  Nivel 1
            Call FGrSetRowStyle(Grid, i, "B")
            Grid.TextMatrix(i, C_CUENTA) = UCase(vFld(Rs("Descripcion")))
            Grid.TextMatrix(i, C_FMT) = "B"
            
         End If
                  
         Total(CurNiv).Debe = 0
         Total(CurNiv).Haber = 0
         Total(CurNiv).Linea = i
         
         For k = 0 To 12
            TotalMes(CurNiv, k).Debe = 0
            TotalMes(CurNiv, k).Haber = 0
            TotalMes(CurNiv, k).Linea = i
         Next k
         
      End If
   
      'sumamos los totales al nivel actual y a los niveles anteriores
      For j = CurNiv To 1 Step -1
         Total(j).Debe = Total(j).Debe + vFld(Rs("Debe"))
         Total(j).Haber = Total(j).Haber + vFld(Rs("Haber"))
         
         If lMensual Then
            If vFld(Rs("Mes")) > 0 Then
               Mes = vFld(Rs("Mes"))
               TotalMes(j, Mes).Debe = TotalMes(j, Mes).Debe + vFld(Rs("Debe"))
               TotalMes(j, Mes).Haber = TotalMes(j, Mes).Haber + vFld(Rs("Haber"))
            End If
         End If
      Next j
            
      Rs.MoveNext
      
   Loop
      
   'ponemos el total de la última línea
   If CurCta <> "" Then
      'ponemos totales de cuenta actual
      Grid.TextMatrix(Total(CurNiv).Linea, C_DEBITOS) = Format(Total(CurNiv).Debe, BL_NUMFMT)
      Grid.TextMatrix(Total(CurNiv).Linea, C_CREDITOS) = Format(Total(CurNiv).Haber, BL_NUMFMT)
      
      If lMensual Then
         For k = 1 To 12
            Mes = C_INI_MES + k - 1
            Grid.TextMatrix(Total(CurNiv).Linea, Mes) = Format(TotalMes(CurNiv, k).Debe - TotalMes(CurNiv, k).Haber, NEGNUMFMT)
         Next k
      End If
      
      'asignamos los totales hacia arriba
      For j = CurNiv - 1 To 2 Step -1
         If j > 1 Then
            Grid.TextMatrix(Total(j).Linea, C_DEBITOS) = Format(Total(j).Debe, BL_NUMFMT)
            Grid.TextMatrix(Total(j).Linea, C_CREDITOS) = Format(Total(j).Haber, BL_NUMFMT)
            
            If lMensual Then
               For k = 1 To 12
                  Mes = C_INI_MES + k - 1
                  If ClasifPadre = CLASCTA_ACTIVO Then
                     Grid.TextMatrix(Total(j).Linea, Mes) = Format(TotalMes(j, k).Debe - TotalMes(j, k).Haber, NEGNUMFMT)
                  ElseIf ClasifPadre = CLASCTA_PASIVO Or ClasifPadre = CLASCTA_RESULTADO Then
                    Grid.TextMatrix(Total(j).Linea, Mes) = Format(TotalMes(j, k).Haber - TotalMes(j, k).Debe, NEGNUMFMT)
                  End If
               Next k
            End If

         End If
      Next j
      
      If CodPadre <> "" Then 'había uno antes
         
         'agregamos la cuenta de resultado del ejercicio
         If lBalClasif And ClasifPadre = CLASCTA_PASIVO Then
            Call AddResEjercicio(i, LinPatrimonio, LinResEjercicio)
         End If
         
         i = i + 1
         Grid.Rows = i + 1
         Call FGrSetRowStyle(Grid, i, "B")
         
         Grid.TextMatrix(i, C_CUENTA) = "TOTAL " & UCase(NomPadre)
         Grid.TextMatrix(i, C_IDCUENTA) = TOT_CUENTA
         Grid.TextMatrix(i, C_CLASCTA) = ClasifPadre
         Grid.TextMatrix(i, C_FMT) = "B"
         
         Grid.TextMatrix(i, C_DEBITOS) = Format(Total(1).Debe, BL_NUMFMT)
         Grid.TextMatrix(i, C_CREDITOS) = Format(Total(1).Haber, BL_NUMFMT)
         
         If ClasifPadre = CLASCTA_RESULTADO Then           'si es Resultado, pueden haber 2 cuentas de primer nivel de este tipo, por lo que el total no se pone a este nivel sino abajo en la grilla de totales
            If Ch_VerSubTot = 0 Then
               Grid.RowHeight(i) = 0   'ocultamos el total
            End If
'         Else
'            If TotClasif(ClasifPadre) <> 0 Then
'               MsgBox1 "ATENCIÓN:" & vbCrLf & vbCrLf & "Revise su plan de cuentas. Hay dos o más cuentas de primer nivel clasificadas como " & UCase(gClasCta(ClasifPadre)) & "." & vbCrLf & "Esto puede generar errores en el Balance.", vbExclamation
'            End If
         End If
         
         If ClasifPadre = CLASCTA_ACTIVO Then
            TotClasif(ClasifPadre) = Total(1).Debe - Total(1).Haber
         ElseIf ClasifPadre = CLASCTA_PASIVO Or ClasifPadre = CLASCTA_RESULTADO Then
            TotClasif(ClasifPadre) = Total(1).Haber - Total(1).Debe
         End If
         
         LinTotClasif(ClasifPadre) = i
         
         If lMensual Then
            For k = 1 To 12
               Mes = C_INI_MES + k - 1
               If ClasifPadre = CLASCTA_ACTIVO Then
                  Grid.TextMatrix(i, Mes) = Format(TotalMes(1, k).Debe - TotalMes(1, k).Haber, NEGNUMFMT)
               ElseIf ClasifPadre = CLASCTA_PASIVO Or ClasifPadre = CLASCTA_RESULTADO Then
                  Grid.TextMatrix(i, Mes) = Format(TotalMes(1, k).Haber - TotalMes(1, k).Debe, NEGNUMFMT)
               End If
            Next k
         End If
         
      End If
      
   End If
   
   Call CloseRs(Rs)
      
   TotalFinal = 0
   
   'calculamos la columna Valor como la diferencia de Créditos y Débitos y ocultamos filas con valor 0
   For Row = Grid.FixedRows To Grid.Rows - 1
   
      If Trim(Grid.TextMatrix(Row, C_CODIGO)) <> "" And vFmt(Grid.TextMatrix(Row, C_IDCUENTA)) <> TOT_CUENTA Then
         
         If Val(Grid.TextMatrix(Row, C_CLASCTA)) = CLASCTA_ACTIVO Then
            Diff = vFmt(Grid.TextMatrix(Row, C_DEBITOS)) - vFmt(Grid.TextMatrix(Row, C_CREDITOS))
         Else
            Diff = vFmt(Grid.TextMatrix(Row, C_CREDITOS)) - vFmt(Grid.TextMatrix(Row, C_DEBITOS))
         End If
         
         Grid.TextMatrix(Row, C_VALOR) = Format(Diff, NEGNUMFMT)
                  
         If lMensual Then
            Grid.TextMatrix(Row, C_SALDOFIN) = Format(Diff, NEGNUMFMT)
            
            RowVisible = False
            For Col = C_INI_MES To C_SALDOFIN
               If vFmt(Grid.TextMatrix(Row, Col)) <> 0 Then
                  RowVisible = True
                  Exit For
               End If
            Next Col
            
            If Not RowVisible Then
               Grid.RowHeight(Row) = 0
            End If
            
         Else
            If Diff = 0 Then
               Grid.RowHeight(Row) = 0
            End If
         
         End If
         
      ElseIf vFmt(Grid.TextMatrix(Row, C_IDCUENTA)) = TOT_CUENTA Then
         
         Call FGrSetRowStyle(Grid, Row, "B")
         
         If Val(Grid.TextMatrix(Row, C_CLASCTA)) = CLASCTA_ACTIVO Then
            Diff = vFmt(Grid.TextMatrix(Row, C_DEBITOS)) - vFmt(Grid.TextMatrix(Row, C_CREDITOS))
         Else
            Diff = vFmt(Grid.TextMatrix(Row, C_CREDITOS)) - vFmt(Grid.TextMatrix(Row, C_DEBITOS))
         End If
         
         Grid.TextMatrix(Row, C_VALOR) = Format(Diff, NEGNUMFMT)
         
         If Abs(TotalFinal) < Abs(Diff) Or Abs(TotalFinal) = 0 Then
            NombreCtaTotal = Grid.TextMatrix(Row, C_CUENTA)
         End If

         TotalFinal = TotalFinal + Diff
         
         If lMensual Then
            For k = 1 To 12
               Mes = C_INI_MES + k - 1
               TotalFinalMes(k) = TotalFinalMes(k) + vFmt(Grid.TextMatrix(Row, Mes))
               
               If k < Month(FDesde) Or k > Month(FHasta) Then
                  Grid.ColWidth(Mes) = 0
                  GridTot.ColWidth(Mes) = 0
               End If
            
            Next k
         
            Grid.TextMatrix(Row, C_SALDOFIN) = Format(Diff, NEGNUMFMT)
            
         End If
      End If
      
   Next Row
   
   'ponemos el resultado del ejercicio si correponde
   If lBalClasif Then
   
      ResEjercicio = TotClasif(CLASCTA_ACTIVO) - TotClasif(CLASCTA_PASIVO)
      
      If ResEjercicio <> 0 Then
         For k = 0 To UBound(LinResEjercicio)
            If LinResEjercicio(k) = 0 Then
               Exit For
            End If
            Grid.TextMatrix(LinResEjercicio(k), C_VALOR) = Format(ResEjercicio, NEGNUMFMT)
            If vFmt(Grid.TextMatrix(LinResEjercicio(k), C_VALOR)) <> 0 And Nivel >= Val(Grid.TextMatrix(LinResEjercicio(k), C_NIVEL)) Then
               Grid.RowHeight(LinResEjercicio(k)) = Grid.RowHeight(k)
            End If
         Next k
         If LinPatrimonio > 0 Then
            Grid.TextMatrix(LinPatrimonio, C_VALOR) = Format(vFmt(Grid.TextMatrix(LinPatrimonio, C_VALOR)) + ResEjercicio, NEGNUMFMT)
            If vFmt(Grid.TextMatrix(LinPatrimonio, C_VALOR)) <> 0 And Nivel >= Val(Grid.TextMatrix(LinPatrimonio, C_NIVEL)) Then
               Grid.RowHeight(LinPatrimonio) = Grid.RowHeight(0)
            End If
         End If
         
         If LinTotClasif(CLASCTA_PASIVO) > 0 Then
            Grid.TextMatrix(LinTotClasif(CLASCTA_PASIVO), C_VALOR) = Format(vFmt(Grid.TextMatrix(LinTotClasif(CLASCTA_PASIVO), C_VALOR)) + ResEjercicio, NEGNUMFMT)
         End If
      
         TotalFinal = TotalFinal + ResEjercicio
      Else
         For k = 0 To UBound(LinResEjercicio)
            If LinResEjercicio(k) = 0 Then
               Exit For
            End If
            Grid.RowHeight(LinResEjercicio(k)) = 0
         Next k
      End If
   End If

   If ClasifPadre <> CLASCTA_RESULTADO Then
      GridTot.TextMatrix(0, C_CUENTA) = "TOTAL"
   Else
      GridTot.TextMatrix(0, C_CUENTA) = NombreCtaTotal
   End If
   GridTot.TextMatrix(0, C_VALOR) = Format(TotalFinal, NEGNUMFMT)

   If lMensual Then
      
      If lComparativo Then
      
         For Row = Grid.FixedRows To Grid.Rows - 1
         
            If Grid.TextMatrix(Row, C_IDCUENTA) = "" Then
               Exit For
            End If
         
            TotMesHasta = 0
            TotMesMenos1 = 0
            For k = Month(FDesde) To Month(FHasta)
               TotMesHasta = TotMesHasta + vFmt(Grid.TextMatrix(Row, C_INI_MES + k - 1))
               
               If k < Month(FHasta) Then
                  TotMesMenos1 = TotMesMenos1 + vFmt(Grid.TextMatrix(Row, C_INI_MES + k - 1))
               End If
            Next k
            
            If Trim(Grid.TextMatrix(Row, C_CODIGO)) <> "" And vFmt(Grid.TextMatrix(Row, C_IDCUENTA)) <> TOT_CUENTA Then
               Grid.TextMatrix(Row, C_INI_COMP) = Format(TotMesMenos1, NEGNUMFMT)
               Grid.TextMatrix(Row, C_INI_COMP + 1) = Format(TotMesHasta - TotMesMenos1, NEGNUMFMT)
               Grid.TextMatrix(Row, C_INI_COMP + 2) = Format(TotMesHasta, NEGNUMFMT)
            
            ElseIf vFmt(Grid.TextMatrix(Row, C_IDCUENTA)) = TOT_CUENTA Then
               Grid.TextMatrix(Row, C_INI_COMP) = Format(TotMesMenos1, NEGNUMFMT)
               Grid.TextMatrix(Row, C_INI_COMP + 1) = Format(TotMesHasta - TotMesMenos1, NEGNUMFMT)
               Grid.TextMatrix(Row, C_INI_COMP + 2) = Format(TotMesHasta, NEGNUMFMT)
            
               TotFinMenos1 = TotFinMenos1 + TotMesMenos1
               TotFinHasta = TotFinHasta + TotMesHasta
            End If
                  
         Next Row
      
         Call FirstLastMonthDay(DateAdd("m", -1, FHasta), FirstDiaMes, LastDiaMes)
      
         Grid.TextMatrix(1, C_INI_COMP) = "Al " & Format(LastDiaMes, SDATEFMT)
         Grid.TextMatrix(1, C_INI_COMP + 2) = "Al " & Format(FHasta, SDATEFMT)
         
         GridTot.TextMatrix(0, C_INI_COMP) = Format(TotFinMenos1, NEGNUMFMT)
         GridTot.TextMatrix(0, C_INI_COMP + 1) = Format(TotFinHasta - TotFinMenos1, NEGNUMFMT)
         GridTot.TextMatrix(0, C_INI_COMP + 2) = Format(TotFinHasta, NEGNUMFMT)
      
      Else
         For k = 1 To 12
            Mes = C_INI_MES + k - 1
            GridTot.TextMatrix(0, Mes) = Format(TotalFinalMes(k), NEGNUMFMT)
         Next k
         
         GridTot.TextMatrix(0, C_SALDOFIN) = Format(TotalFinal, NEGNUMFMT)
      End If
   End If
   
   
   Call FGrVRows(Grid, 1)
   
   'borramos los títulos de las columnas con ancho 0
   For i = 0 To Grid.Cols - 1
      If Grid.ColWidth(i) = 0 Then
         Grid.TextMatrix(0, i) = ""
      End If
   Next i
   
   'Grid.Rows = Grid.Rows + 25
      
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.Col = C_CODIGO
   Grid.RowSel = Grid.Row
   Grid.ColSel = Grid.Col
   
   Grid.Redraw = True
   Call EnableFrm(False)
  
End Sub
Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - GridTot.Height - 500
   GridTot.Top = Grid.Top + Grid.Height + 30
   Grid.Width = Me.Width - 330
   GridTot.Width = Grid.Width - 230
   
   Call FGrVRows(Grid, 1)

End Sub
Private Sub EnableFrm(bool As Boolean)
   Bt_Buscar.Enabled = bool
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

Private Sub Grid_Click()
   Fr_MultiSelect.Visible = False

End Sub

Private Sub Grid_DblClick()
   Call Bt_VerLibMayor_Click
End Sub


Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol

End Sub

Private Sub Ls_AreaNeg_Click()

   Call EnableFrm(True)
   
   If Ls_AreaNeg.SelCount > 0 Then
      Cb_AreaNeg.ListIndex = 0
      Cb_AreaNeg.Enabled = False
   Else
      Cb_AreaNeg.Enabled = True
   End If

End Sub

Private Sub Ls_CCosto_Click()
   Call EnableFrm(True)
   
   If Ls_CCosto.SelCount > 0 Then
      Cb_CCosto.ListIndex = 0
      Cb_CCosto.Enabled = False
   Else
      Cb_CCosto.Enabled = True
   End If
   
End Sub


Private Sub tx_Desde_Change()
   Call EnableFrm(True)
End Sub

Private Sub Tx_Desde_GotFocus()
   Call DtGotFocus(Tx_Desde)
End Sub

Private Sub Tx_Desde_LostFocus()
   
   If Trim$(Tx_Desde) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Desde)
   
End Sub

Private Sub Tx_Desde_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub

Private Sub tx_Hasta_Change()
   Call EnableFrm(True)
End Sub

Private Sub Tx_Hasta_GotFocus()
   Call DtGotFocus(Tx_Hasta)
   
End Sub

Private Sub Tx_Hasta_LostFocus()
   
   If Trim$(Tx_Hasta) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Hasta)
      
End Sub

Private Sub Tx_Hasta_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Bt_Fecha_Click(Index As Integer)
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   If Index = 0 Then
      Call Frm.TxSelDate(Tx_Desde)
   Else
      Call Frm.TxSelDate(Tx_Hasta)
   End If
   
   Set Frm = Nothing
   
End Sub
Public Function FViewBalClasif(Optional ByVal Mes As Integer = 0)
   lMes = Mes

   lClasCta = CLASCTA_ACTIVO & "," & CLASCTA_PASIVO
   lCaption = "Balance Clasificado"
   lLibOf = LIBOF_CLASIFICADO
   
   lBalClasif = True
   
   Me.Show vbModeless
End Function
Public Function FViewEstResultClasif(Optional ByVal Mes As Integer = 0)
   lMes = Mes

   lClasCta = CLASCTA_RESULTADO
   lCaption = "Estado de Resultado Clasificado"
   lLibOf = LIBOF_ESTRESCLASIF
   
   Me.Show vbModeless
End Function
Public Function FViewEstResultMensual(Optional ByVal Mes As Integer = 0)
   lMes = Mes

   lMensual = True
   lClasCta = CLASCTA_RESULTADO
   lCaption = "Estado de Resultado Mensual"
   lLibOf = LIBOF_ESTRESMENSUAL

   Me.Show vbModeless
End Function

Public Function FViewEstResultComparativo(Optional ByVal Mes As Integer = 0)
   lMes = Mes

   lMensual = True
   lComparativo = True
   lClasCta = CLASCTA_RESULTADO
   lCaption = "Estado de Resultado Comparativo Mes Anterior"
   lLibOf = LIBOF_ESTRESCOMP
   
   Me.Show vbModeless
End Function

Private Sub Cb_AreaNeg_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_CCosto_Click()
   Call EnableFrm(True)

End Sub

Private Sub AddResEjercicio(i As Integer, LinPatrimonio As Integer, LinResEjercicio() As Integer)
   Dim Nivel As Integer
   Dim NivRes As Integer, NivAux As Integer
   Dim NivPat As Integer
   Dim SpIndent As String
   Dim j As Integer, k As Integer
   Dim IdPadre As Long, IdPadreAux As Long
   Dim DescRes As String
   Dim CodRes As String
      
   If LinPatrimonio = 0 Then
      MsgBox1 "No se encontró la cuenta de Patrimonio, de acuerdo a la ""Definición de Cuentas Básicas"" de la ""Configuración Inicial"" de la empresa.", vbExclamation
      Exit Sub
   End If
   
   If LinResEjercicio(0) > 0 Then   'ya se agregó el resultado
      Exit Sub
   End If
   
   Nivel = Val(Cb_Nivel)
   NivPat = Val(Grid.TextMatrix(LinPatrimonio, C_NIVEL))
   
   If lResEje(gLastNivel).Codigo = "" Then
      MsgBox1 "No se encontró la cuenta de Resultado Ejercicio, de acuerdo a la ""Definición de Cuentas Básicas"" de la ""Configuración Inicial"" de la empresa.", vbExclamation
      Exit Sub
   End If
     
   k = 0
   For j = NivPat + 1 To Nivel
      
      i = i + 1
      Grid.Rows = i + 1
      
      SpIndent = String(REP_INDENT * (j - 2), " ")
      Call FGrSetRowStyle(Grid, i, "FC", gColores(j))
      Grid.TextMatrix(i, C_FMT) = "C" & gColores(j)
      Grid.TextMatrix(i, C_CODIGO) = FmtCodCuenta(lResEje(j).Codigo)
      Grid.TextMatrix(i, C_CUENTA) = SpIndent & lResEje(j).Descripcion
      Grid.TextMatrix(i, C_NIVEL) = lResEje(j).Nivel
      Grid.TextMatrix(i, C_IDCUENTA) = "*"
      
      LinResEjercicio(k) = i
      k = k + 1
   Next j
   
End Sub
Private Sub ReadResEje()
   Dim i As Integer
   
   i = gLastNivel
   lResEje(gLastNivel).id = gCtasBas.IdCtaResEje
   
   Do While lResEje(i).id > 0
      lResEje(i).Codigo = GetDatosCuenta(lResEje(i).id, lResEje(i).Descripcion, lResEje(i).IdPadre, lResEje(i).Nivel)
      i = i - 1
      lResEje(i).id = lResEje(i + 1).IdPadre
   Loop

End Sub
Private Sub SetupPriv()
   If Not ChkPriv(PRV_IMP_LIBOF) Then
      Ch_LibOficial = 0
      Ch_LibOficial.Enabled = False
   End If
End Sub

Private Sub Ch_VerSubTot_Click()
   Call EnableFrm(True)

End Sub

