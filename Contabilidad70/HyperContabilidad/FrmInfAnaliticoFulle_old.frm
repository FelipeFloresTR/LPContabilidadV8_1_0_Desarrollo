VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmInfAnaliticoFulle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Analítico de Cuentas"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   Icon            =   "FrmInfAnaliticoFulle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   13860
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   30
      Top             =   0
      Width           =   13635
      Begin VB.CommandButton Bt_VerComp 
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
         Picture         =   "FrmInfAnaliticoFulle.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Detalle comprobante seleccionado"
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
         Left            =   540
         Picture         =   "FrmInfAnaliticoFulle.frx":0485
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Detalle documento seleccionado"
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
         Left            =   1500
         Picture         =   "FrmInfAnaliticoFulle.frx":08F9
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   11520
         TabIndex        =   28
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
         Left            =   1080
         Picture         =   "FrmInfAnaliticoFulle.frx":0DB3
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   1920
         Picture         =   "FrmInfAnaliticoFulle.frx":125A
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   3840
         Picture         =   "FrmInfAnaliticoFulle.frx":169F
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   3000
         Picture         =   "FrmInfAnaliticoFulle.frx":1AC8
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   3420
         Picture         =   "FrmInfAnaliticoFulle.frx":1E66
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   2460
         Picture         =   "FrmInfAnaliticoFulle.frx":21C7
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   120
      TabIndex        =   27
      Top             =   600
      Width           =   13635
      Begin VB.ComboBox Cb_Glosa 
         Height          =   315
         Left            =   10080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1380
         Width           =   2655
      End
      Begin VB.CheckBox Ch_SaldoApertura 
         Caption         =   "Saldo de Apertura"
         Height          =   195
         Left            =   1860
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Bt_SelCuenta 
         Height          =   315
         Left            =   9300
         Picture         =   "FrmInfAnaliticoFulle.frx":226B
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar cuenta"
         Top             =   240
         Width           =   315
      End
      Begin VB.ComboBox Cb_TipoInforme 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1380
         Width           =   6015
      End
      Begin VB.TextBox Tx_HastaComp 
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   600
         Width           =   1275
      End
      Begin VB.CommandButton Bt_FechaComp 
         Height          =   315
         Left            =   2760
         Picture         =   "FrmInfAnaliticoFulle.frx":2660
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   225
      End
      Begin VB.CheckBox Ch_CuentasRUT 
         Caption         =   "Sólo cuentas con RUT asociado"
         Height          =   195
         Left            =   7020
         TabIndex        =   12
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CheckBox Ch_Rut 
         Caption         =   "RUT:"
         Height          =   195
         Left            =   7740
         TabIndex        =   7
         Top             =   600
         Width           =   225
      End
      Begin VB.CheckBox Ch_DocsComp 
         Caption         =   "Sólo documentos contabilizados en el Libro Mayor a la fecha informe"
         Height          =   195
         Left            =   1920
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.CheckBox Ch_SaldosVig 
         Caption         =   "Saldos Vigentes"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Ch_InfRes 
         Caption         =   "Informe Resumido"
         Height          =   195
         Left            =   3660
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Ch_DetComp 
         Caption         =   "Detalle por Comprobante"
         Height          =   195
         Left            =   7320
         TabIndex        =   13
         Top             =   1440
         Width           =   2055
      End
      Begin VB.ComboBox Cb_Nombre 
         Height          =   315
         Left            =   5040
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   2655
      End
      Begin VB.TextBox Tx_Rut 
         Height          =   315
         Left            =   8460
         MaxLength       =   12
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   540
         Width           =   1155
      End
      Begin VB.ComboBox Cb_Cuenta 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   4275
      End
      Begin VB.CommandButton Bt_Search 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   11520
         Picture         =   "FrmInfAnaliticoFulle.frx":296A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Left            =   2760
         Picture         =   "FrmInfAnaliticoFulle.frx":2DA8
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   225
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Lbl_Glosa 
         Caption         =   "Glosa:"
         Height          =   255
         Left            =   9600
         TabIndex        =   39
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo informe:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Comprobantes al:"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   660
         Width           =   1230
      End
      Begin VB.Label Lb_FechaEmi 
         AutoSize        =   -1  'True
         Caption         =   "(fecha emisión)"
         Height          =   195
         Left            =   3060
         TabIndex        =   35
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Left            =   7980
         TabIndex        =   34
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label3 
         Caption         =   "Entidad:"
         Height          =   255
         Left            =   4380
         TabIndex        =   32
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta:"
         Height          =   255
         Left            =   4380
         TabIndex        =   31
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Lb_DocAl 
         AutoSize        =   -1  'True
         Caption         =   "Documentos al:"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   240
         Width           =   1110
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5775
      Left            =   60
      TabIndex        =   17
      Top             =   2520
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   10186
      _Version        =   393216
      Rows            =   4
      Cols            =   14
      FixedCols       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   60
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   8280
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   8
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmInfAnaliticoFulle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDENTIDAD = 0
Const C_ENTIDAD = 1
Const C_IDCOMP = 2
Const C_IDDOC = 3
Const C_DOC = 4         'TipoDoc-NumDoc
Const C_FECHADOC = 5
Const C_FECHAVENC = 6
Const C_GLOSA = 7
Const C_GLOSACOMP = 8
Const C_DEBE = 9
Const C_HABER = 10
Const C_SALDO = 11
Const C_SALDOCALC = 12
Const C_FMT = 13
Const C_SUBTITULO = 14
Const C_OBLIGATORIA = 14

Const NCOLS = C_OBLIGATORIA

Const M_IDENTIDAD = 1
Const M_RUT = 2
Const M_NOTVALIDRUT = 3

'Tipo de informe
Const TI_TODOS = 0
Const TI_SOLODOCS_CONCOMP = 1
Const TI_SOLODOCS_SINCOMP = 2


'Glosa Comprobante
Const G_AMBOS = 1
Const G_DOCUMENTO = 2
Const G_COMPROBANTE = 3


Dim lcbNombre As ClsCombo

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
'2758025
Dim lInfoPreliminar As Boolean

Dim lMes As Integer

Dim lPorEntidad As Boolean

Dim lViewOtrosDocs As Boolean


Public Function FViewPorEntidad(ByVal Mes As Integer)
   Dim MesActual As Integer

   lPorEntidad = True
   
   lMes = Mes
   
   MesActual = GetMesActual()
   
   If lMes = 0 Then
      If MesActual > 0 Then
         lMes = MesActual
      Else
         lMes = GetUltimoMesConMovs()
      End If
   End If
         
   Me.Show vbModeless
   
End Function

Public Function FViewPorCuenta(ByVal Mes As Integer)
   Dim MesActual As Integer

   lPorEntidad = False
   
   lMes = Mes
   
   MesActual = GetMesActual()
   
   If lMes = 0 Then
      If MesActual > 0 Then
         lMes = MesActual
      Else
         lMes = GetUltimoMesConMovs()
      End If
   End If
         
   Me.Show vbModeless
   
End Function
Public Function FViewOtrosDocs(ByVal Mes As Integer)
   Dim MesActual As Integer

   lPorEntidad = True
   lViewOtrosDocs = True
   
   lMes = Mes
   
   MesActual = GetMesActual()
   
   If lMes = 0 Then
      If MesActual > 0 Then
         lMes = MesActual
      Else
         lMes = GetUltimoMesConMovs()
      End If
   End If
         
   Me.Show vbModeless
   
End Function


Private Sub SetUpGrid()
   Dim Col As Integer
   
   Grid.Cols = NCOLS + 1
   
   Grid.ColWidth(C_IDENTIDAD) = 0
   Grid.ColWidth(C_ENTIDAD) = 2300
   Grid.ColWidth(C_IDCOMP) = 0
   Grid.ColWidth(C_IDDOC) = 0
   Grid.ColWidth(C_DOC) = 1300
   Grid.ColWidth(C_FECHADOC) = 800
   Grid.ColWidth(C_FECHAVENC) = 800
   Grid.ColWidth(C_GLOSA) = 2300
   Grid.ColWidth(C_GLOSACOMP) = 2300
   Grid.ColWidth(C_DEBE) = 1300
   Grid.ColWidth(C_HABER) = 1300
   Grid.ColWidth(C_SALDO) = 1350
   Grid.ColWidth(C_SALDOCALC) = 0 '1450
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_SUBTITULO) = 0
   Grid.ColWidth(C_OBLIGATORIA) = 0
   
   Grid.ColAlignment(C_ENTIDAD) = flexAlignLeftCenter
   Grid.ColAlignment(C_DOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_FECHADOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_FECHAVENC) = flexAlignLeftCenter
   Grid.ColAlignment(C_GLOSA) = flexAlignLeftCenter
   Grid.ColAlignment(C_GLOSACOMP) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDO) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDOCALC) = flexAlignCenterCenter ' flexAlignRightCenter
   
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot)
   
   If lPorEntidad Then
      Grid.TextMatrix(0, C_ENTIDAD) = "Entidad"
   Else
      Grid.TextMatrix(0, C_ENTIDAD) = "Cuenta"
   End If
   Grid.TextMatrix(0, C_DOC) = "Documento"
   Grid.TextMatrix(0, C_FECHADOC) = "Emisión"
   Grid.TextMatrix(0, C_FECHAVENC) = "Vencim."
   Grid.TextMatrix(0, C_GLOSA) = "Glosa"
   Grid.TextMatrix(0, C_GLOSACOMP) = "Glosa Comprobante"
   Grid.TextMatrix(0, C_DEBE) = "Debe"
   Grid.TextMatrix(0, C_HABER) = "Haber"
   Grid.TextMatrix(0, C_SALDO) = "Saldo"
   Grid.TextMatrix(0, C_SALDOCALC) = "" '"Saldo Otr Doc Año"
   Grid.TextMatrix(0, C_FMT) = "          .FMT"
   
   Call FGrVRows(Grid)
      
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer, j As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Titulos(2) As String
   Dim Encabezados(0) As String
   Dim FontTit(0) As FontDef_t
   Dim FontEnc(0) As FontDef_t
   Dim Nombres(5) As String
   Dim OldOrient As Integer
   Dim Total(NCOLS) As String
   
   Set gPrtReportes.Grid = Grid
   
    Call FolioEncabEmpresa(Not lPapelFoliado, lOrientacion)
   
   Printer.Orientation = lOrientacion
   
   Titulos(0) = Caption
   If ItemData(Cb_Cuenta) > 0 Then
      Titulos(1) = Trim(Mid(Cb_Cuenta, InStr(Cb_Cuenta, " ")))
   End If
   
   '2758025
   If lInfoPreliminar Then
      Titulos(2) = INFO_PRELIMINAR
      FontTit(2).FontBold = True
   End If
   
   Encabezados(0) = "Documentos pendientes al " & Format(GetTxDate(Tx_Hasta), DATEFMT)
   If Tx_Rut <> "" Then
      Encabezados(0) = Encabezados(0) & " de " & Cb_Nombre & " RUT: " & Tx_Rut
   End If
   
   gPrtReportes.Titulos = Titulos
   
   FontTit(0).FontBold = True
   Call gPrtReportes.FntTitulos(FontTit())
   
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   gPrtReportes.Encabezados = Encabezados
   FontEnc(0).FontBold = True
   FontEnc(0).FontName = "Arial"
   FontEnc(0).FontSize = 10
   Call gPrtReportes.FntEncabezados(FontEnc())
    
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
      
   If Grid.ColWidth(C_GLOSA) <> 0 Then
        ColWi(C_GLOSA) = ColWi(C_GLOSA) - 200
   End If
   
   If Grid.ColWidth(C_GLOSACOMP) <> 0 Then
        ColWi(C_GLOSACOMP) = ColWi(C_GLOSACOMP) - 200
   End If
   
   If lOrientacion = ORIENT_HOR Then
    If ItemData(Cb_Glosa) = G_AMBOS Then
         ColWi(C_GLOSA) = ColWi(C_GLOSA) - 400
         ColWi(C_GLOSACOMP) = ColWi(C_GLOSACOMP) - 400
    End If
   End If
   
    If lOrientacion = ORIENT_VER Then
      gPrtReportes.GrFontSize = 8
   Else
      gPrtReportes.GrFontSize = 9
      For i = 0 To Grid.Cols - 1
         ColWi(i) = ColWi(i) * 1.2
      Next i
      
   End If
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_OBLIGATORIA
   gPrtReportes.Total = Total
   gPrtReportes.NTotLines = 1
   
End Sub
Private Sub Bt_Search_Click()
      
   MousePointer = vbHourglass
   
   If ExitDemo() Then
      Unload Me
   End If

   Call LoadAll
   
   MousePointer = vbDefault
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub
Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   
   '2758025
   Dim FrmPrt As FrmPrtSetup
   Dim OldOrientacion As Integer
   
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de seleccionar la vista previa.", vbExclamation
      Exit Sub
   End If
   
   lPapelFoliado = False
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   '2758025
   Set FrmPrt = New FrmPrtSetup
   If FrmPrt.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar, False) = vbOK Then
   
   OldOrientacion = Printer.Orientation
   '2758025
    Printer.Orientation = lOrientacion
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault

   Call Frm.FView(Caption)
   Set Frm = Nothing
   
    Printer.Orientation = OldOrientacion
      lInfoPreliminar = False
   
   End If
   
End Sub
Private Sub Bt_Print_Click()
   Dim OldOrientation As Integer
   '2758025
     Dim Frm As FrmPrtSetup
     
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   '2758025
    Set Frm = New FrmPrtSetup
   If Frm.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar) = vbOK Then
   
   OldOrientation = Printer.Orientation
    
    '2758025
    Printer.Orientation = lOrientacion
   Call SetUpPrtGrid
 
   Me.MousePointer = vbHourglass
  
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
   lInfoPreliminar = False
   End If
End Sub
Private Sub Bt_CopyExcel_Click()
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de copiar.", vbExclamation
      Exit Sub
   End If
   
   Call FGr2Clip(Grid, "Cuenta: " & Cb_Cuenta & "   Documentos pendientes al: " & Tx_Hasta)
End Sub


Private Sub Bt_SelCuenta_Click()
   Dim FrmPlan As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Nombre As String
   Dim Descrip As String
   
   Set FrmPlan = New FrmPlanCuentas
   If FrmPlan.FSelect(IdCuenta, Codigo, Nombre, Descrip, False) = vbOK Then
      Call CbSelItem(Cb_Cuenta, IdCuenta)
   End If
   
   Set FrmPlan = Nothing
   
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

Private Sub Bt_VerComp_Click()
   Dim Row As Integer
   Dim Col As Integer
   Dim Frm As FrmComprobante
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDCOMP)) > 0 Then
      Set Frm = New FrmComprobante
      Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDCOMP)), False)
      Set Frm = Nothing
   End If

End Sub

Private Sub Bt_VerDoc_Click()
   Dim Row As Integer
   Dim Col As Integer
   Dim Frm As FrmDoc
   Dim TipDoc As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 Then
      TipDoc = 0
      If InStr(Grid.TextMatrix(Row, C_DOC), "ODF") > 0 Then
        TipDoc = 8
      End If
      Set Frm = New FrmDoc
      Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDDOC)), TipDoc)
      Set Frm = Nothing
   End If

End Sub



Private Sub Cb_Glosa_Click()
Call EnableFrm(True)
End Sub

Private Sub Cb_TipoInforme_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_CuentasRUT_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_DetComp_Click()

   Call EnableFrm(True)
   
   If Ch_DetComp <> 0 Then
'      Ch_SaldosVig = 0                'tmptbl
'      Ch_SaldosVig.Enabled = False    'tmptbl
      Bt_VerDoc.Enabled = True
      Bt_VerComp.Enabled = True
      Cb_Glosa.Enabled = True
   Else
'      Ch_SaldosVig.Enabled = True     'tmptbl
      'Bt_VerDoc.Enabled = False
      Bt_VerComp.Enabled = False
      Cb_Glosa.Enabled = False
      If lPorEntidad Then
        Call CbSelItem(Cb_Glosa, G_DOCUMENTO)
      End If
   End If
   
End Sub

Private Sub Ch_DocsComp_Click()
   Call EnableFrm(True)

   If Ch_DocsComp = 1 Then
      Lb_FechaEmi.visible = False
'      Lb_DocAl.Caption = "Informe al:"
   Else
      Lb_FechaEmi.visible = True
      Lb_DocAl.Caption = "Documentos al:"
   End If
      
End Sub

Private Sub Ch_InfRes_Click()
   
   Call EnableFrm(True)
   
   If Ch_InfRes <> 0 Then
      Ch_DetComp = 0
      Ch_DetComp.Enabled = False
   Else
      Ch_DetComp.Enabled = True
   End If
   
End Sub

Private Sub Ch_Rut_Click()
   Call EnableFrm(True)
End Sub

Private Sub Ch_SaldoApertura_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_SaldosVig_Click()
   Call EnableFrm(True)

End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim D1 As Long, D2 As Long
   Dim ActDate As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim WhereRut As String
      
   Call recalcularSaldos
   ActDate = DateSerial(gEmpresa.Ano, lMes, 1)
   
   Call FirstLastMonthDay(ActDate, D1, D2)
   Call SetTxDate(Tx_Hasta, D2)
   Call SetTxDate(Tx_HastaComp, D2)
      
   If lPorEntidad Then
      WhereRut = " AND Atrib" & ATRIB_RUT & "<> 0"
      Cb_Cuenta.ToolTipText = "Listado de cuentas con atributo de RUT asociado"
   Else
      Cb_Cuenta.ToolTipText = "Listado de todas las cuentas"
   End If
   
   Ch_CuentasRUT.visible = False
   If lPorEntidad Then
      If lViewOtrosDocs Then
         Me.Caption = "Informe Otros Documentos Asociados a Comprobantes"
      Else
         Me.Caption = "Informe Analítico ODF por Entidad"
      End If
      Cb_Glosa.Enabled = IIf(Ch_DetComp <> 0, True, False)
   Else
      Me.Caption = "Informe Analítico ODF por Cuentas"
      Ch_CuentasRUT.visible = True
      Ch_CuentasRUT = 1
      Lbl_Glosa.Left = 7320
      Cb_Glosa.Left = 7800
   End If
   
   Q1 = "SELECT IdCuenta, Codigo, Descripcion FROM Cuentas "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & WhereRut
   Q1 = Q1 & " ORDER BY Codigo "
   Set Rs = OpenRs(DbMain, Q1)
   
   Cb_Cuenta.AddItem " "
   Cb_Cuenta.ItemData(Cb_Cuenta.NewIndex) = 0
   
   Do While Rs.EOF = False
      
      Cb_Cuenta.AddItem FmtCodCuenta(vFld(Rs("Codigo"))) & "  " & FCase(vFld(Rs("Descripcion"), True))
      Cb_Cuenta.ItemData(Cb_Cuenta.NewIndex) = vFld(Rs("IdCuenta"))
   
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   If Not lPorEntidad Then
      Ch_DetComp.visible = False
   End If
   
   Set lcbNombre = New ClsCombo
   Call lcbNombre.SetControl(Cb_Nombre)
   
   '2758025
   'se comenta orientacion vertical
   lOrientacion = ORIENT_VER
   'lOrientacion = ORIENT_HOR
     
   Call SelCbEntidad(-1)
   Ch_SaldosVig = 1
   Ch_Rut = 1
   Ch_DocsComp = 1
   
   Bt_VerComp.Enabled = IIf(Ch_DetComp <> 0, True, False)
   
   Call CbAddItem(Cb_TipoInforme, "Todos los documentos", TI_TODOS)
   Call CbAddItem(Cb_TipoInforme, "Sólo documentos contabilizados en el Libro Mayor", TI_SOLODOCS_CONCOMP)
   Call CbAddItem(Cb_TipoInforme, "Sólo documentos NO contabilizados en el Libro Mayor", TI_SOLODOCS_SINCOMP)
   Call CbSelItem(Cb_TipoInforme, TI_SOLODOCS_CONCOMP)
   Call CargaCbGlosa
   Call CbSelItem(Cb_Glosa, G_DOCUMENTO)
   Call SetUpGrid
   
   
   Ch_SaldoApertura = 1
   
   Call LoadAll

End Sub

Private Sub CargaCbGlosa()

   Cb_Glosa.AddItem "Ambos"
   Cb_Glosa.ItemData(Cb_Glosa.NewIndex) = G_AMBOS
   Cb_Glosa.AddItem "Comprobante"
   Cb_Glosa.ItemData(Cb_Glosa.NewIndex) = G_COMPROBANTE
   Cb_Glosa.AddItem "Documento"
   Cb_Glosa.ItemData(Cb_Glosa.NewIndex) = G_DOCUMENTO


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
   Dim Row As Integer
   Dim Col As Integer
   Dim Frm As Form
   
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Col = C_ENTIDAD Then
      Call PostClick(Bt_VerComp)
   ElseIf Col = C_DOC Then
      Call PostClick(Bt_VerDoc)
   End If
      
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
Private Sub Bt_Fecha_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_Hasta)
   Set Frm = Nothing
   
End Sub
Private Sub tx_HastaComp_Change()
   Call EnableFrm(True)
End Sub

Private Sub Tx_HastaComp_GotFocus()
   Call DtGotFocus(Tx_HastaComp)
   
End Sub

Private Sub Tx_HastaComp_LostFocus()
   
   If Trim$(Tx_HastaComp) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_HastaComp)
      
End Sub

Private Sub Tx_HastaComp_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Bt_FechaComp_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_HastaComp)
   Set Frm = Nothing
   
End Sub

Private Function LoadAll()
   Dim TipoValLib As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim IdEntidad As Long
   Dim IdEnt As Long
   Dim Debe As Double
   Dim Haber As Double
   Dim Saldo As Double
   Dim TotalDoc As Double
   Dim SubTotDebe As Double
   Dim SubTotHaber As Double
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim TotDebeParc As Double
   Dim TotHaberParc As Double
   Dim TotSaldo As Double
   Dim Wh As String
   Dim WhDoc As String
   Dim WhComp As String
   Dim NombEnt As String
   Dim DetComp1 As String
   Dim DetComp2 As String
   Dim IdCuenta As Long
   Dim CurIdDoc As Long
   Dim AnteIdDoc As Long
   Dim TipoComp As String
   Dim DocDebe As Double
   Dim DocHaber As Double
   Dim j As Integer
   Dim NotValidRut As Boolean
   Dim CondOtrosDocs As String
   Dim RsSaldos As Recordset
   Dim TotPagadoAnoAnt As Double
   Dim InsPagoAnoAnt As Boolean, MsgInsPagoAnoAnt As Boolean
   Dim DocPagoAnt As Long
   Dim CondDocOtrosEnAnalitico As String
   Dim TmpTbl As String, QName As String, TmpTbl2 As String
   Dim Rc As Long
   Dim WhSaldoAp As String
   Dim WhCtaRut As String
   Dim WhCtaRut1 As String
   Dim TipoInforme As Long
   Dim RowTit As Integer
   Dim sDelFrom As String, sDelWhere As String
   Dim Glosa As Integer
   
   'ffv 2855046
   Dim SubTotal As Double
   'ffv 2855046
   TmpTbl = DbGenTmpName2(SQL_ACCESS, "tanalit_")       'le ponemos siempre Access para que no le ponga el # al nombre
   TmpTbl2 = DbGenTmpName2(SQL_ACCESS, "tanalit2_")
   QName = DbGenTmpName2(SQL_ACCESS, "qanalit_")

   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl2)
   Call ExecSQL(DbMain, "DROP TABLE " & QName)
   
   Grid.Redraw = False
   
   'ajustamos títulos
   If Ch_InfRes = 0 Then
      Grid.TextMatrix(0, C_DOC) = "Documento"
      Grid.TextMatrix(0, C_FECHADOC) = "Emisión"
      Grid.TextMatrix(0, C_FECHAVENC) = "Vencim."
      Grid.TextMatrix(0, C_GLOSA) = "Glosa"
   Else
      Grid.TextMatrix(0, C_DOC) = "RUT"
      Grid.TextMatrix(0, C_FECHADOC) = ""
      Grid.TextMatrix(0, C_FECHAVENC) = ""
      Grid.TextMatrix(0, C_GLOSA) = ""
   End If
   
   Grid.ColWidth(C_GLOSA) = 2300
   Grid.ColWidth(C_GLOSACOMP) = 2300
   Grid.TextMatrix(0, C_GLOSA) = "Glosa"
   Grid.TextMatrix(0, C_GLOSACOMP) = "Glosa Comprobante"
   Glosa = ItemData(Cb_Glosa)
   If Glosa = G_DOCUMENTO Then
        Grid.ColWidth(C_GLOSA) = 2300
        Grid.ColWidth(C_GLOSACOMP) = 0
        Grid.TextMatrix(0, C_GLOSA) = "Glosa"
        Grid.TextMatrix(0, C_GLOSACOMP) = ""
   ElseIf Glosa = G_COMPROBANTE Then
        Grid.ColWidth(C_GLOSA) = 0
        Grid.ColWidth(C_GLOSACOMP) = 2300
        Grid.TextMatrix(0, C_GLOSA) = ""
        Grid.TextMatrix(0, C_GLOSACOMP) = "Glosa Comprobante"
   End If
   
   If lPorEntidad Then
      If Ch_DetComp.Value <> 0 Then
         Grid.TextMatrix(0, C_ENTIDAD) = "Entidad / Comprobante"
      Else
         Grid.TextMatrix(0, C_ENTIDAD) = "Entidad"
      End If
   End If
   
   'definimos los filtros
   If ItemData(Cb_Cuenta) > 0 Then
      WhDoc = WhDoc & " AND MovDocumento.IdCuenta = " & ItemData(Cb_Cuenta)
      WhComp = WhComp & " AND Documento.IdCtaBanco = " & ItemData(Cb_Cuenta)
      If lPorEntidad Then
         Call AppendWhere(" DetSaldosAp.IdCuenta = " & ItemData(Cb_Cuenta), WhSaldoAp)
      Else
         Call AppendWhere(" Cuentas.IdCuenta = " & ItemData(Cb_Cuenta), WhSaldoAp)
      End If
   End If

   IdEnt = 0
   If Trim(Tx_Rut) <> "" Then
      IdEnt = GetIdEntidad(Trim(Tx_Rut), NombEnt, NotValidRut)
      If IdEnt > 0 Then
         Wh = Wh & " AND Documento.IdEntidad = " & IdEnt
         Call AppendWhere(" DetSaldosAp.IdEntidad =" & IdEnt, WhSaldoAp)
      ElseIf Cb_Nombre.ListCount > 0 Then
         Tx_Rut = ""
         Cb_Nombre.ListIndex = 0
      End If
   
   End If
   
   If Ch_DetComp <> 0 Then
      DetComp1 = ", 0 AS IdComp, 0 AS Tipo, 0 AS Correlativo, 0 AS Fecha, 0 as TipoAjuste "
      DetComp2 = ", MovComprobante.IdComp, Comprobante.Tipo, Comprobante.Correlativo, Comprobante.Fecha, Comprobante.TipoAjuste "
   End If
   
   
   If Ch_CuentasRUT <> 0 Then
      WhCtaRut = " AND Cuentas.Atrib" & ATRIB_RUT & " <> 0 "
      WhCtaRut1 = " AND Documento.IdCtaBanco <> 0 "
   End If
   
   If lViewOtrosDocs Then
      CondOtrosDocs = " NOT "
   Else
      CondDocOtrosEnAnalitico = " OR (Documento.TipoLib IN (" & LIB_OTROFULL & ") AND Documento.DocOtrosEnAnalitico <> 0)"
   End If
   
   'vemos si hay año anterior
   
   If Ch_SaldoApertura And Not ExistFile(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb") Then
   
      'no hay año anteriro, se usa la información ingresada en el detalle de saldos de apertura

      If gFunciones.DetSaldoApertura And Not lViewOtrosDocs Then
         'primero obtenemos los saldos de Apertura
         If lPorEntidad Then
            Q1 = "SELECT DetSaldosAp.IdEntidad, Nombre, Rut, NotValidRut, Sum(Debe) As SumDebe, Sum(Haber) As SumHaber "
            Q1 = Q1 & " FROM DetSaldosAp INNER JOIN Entidades ON DetSaldosAp.IdEntidad = Entidades.IdEntidad "
            Q1 = Q1 & " AND DetSaldosAp.IdEmpresa = Entidades.IdEmpresa "
            Q1 = Q1 & WhSaldoAp
            If WhSaldoAp <> "" Then
               Q1 = Q1 & " AND "
            Else
               Q1 = Q1 & " WHERE "
            End If
            Q1 = Q1 & " DetSaldosAp.IdEmpresa = " & gEmpresa.id & " AND DetSaldosAp.Ano = " & gEmpresa.Ano
            Q1 = Q1 & " GROUP BY DetSaldosAp.IdEntidad, Nombre, Rut, NotValidRut "
            Q1 = Q1 & " ORDER BY Nombre, DetSaldosAp.IdEntidad "
            
         ElseIf IdEnt > 0 Then
            Q1 = "SELECT Cuentas.IdCuenta, Codigo, Descripcion, 0, Sum(DetSaldosAp.Debe) As SumDebe, Sum(DetSaldosAp.Haber) As SumHaber "
            Q1 = Q1 & " FROM Cuentas INNER JOIN DetSaldosAp ON Cuentas.IdCuenta = DetSaldosAp.IdCuenta "
            Q1 = Q1 & JoinEmpAno(gDbType, "DetSaldosAp", "Cuentas")
            Q1 = Q1 & WhSaldoAp
            If WhSaldoAp <> "" Then
               Q1 = Q1 & " AND "
            Else
               Q1 = Q1 & " WHERE "
            End If
            Q1 = Q1 & " Nivel = " & gLastNivel & " AND (DetSaldosAp.Debe <> 0 OR DetSaldosAp.Haber <> 0) " & WhCtaRut
            Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
            Q1 = Q1 & " GROUP BY Cuentas.IdCuenta, Codigo, Descripcion"
            Q1 = Q1 & " ORDER BY Codigo  "
         
         Else
            Q1 = "SELECT IdCuenta, Codigo, Descripcion, 0, Debe As SumDebe, Haber As SumHaber"
            Q1 = Q1 & " FROM Cuentas "
            Q1 = Q1 & WhSaldoAp
            If WhSaldoAp <> "" Then
               Q1 = Q1 & " AND "
            Else
               Q1 = Q1 & " WHERE "
            End If
            Q1 = Q1 & " Nivel = " & gLastNivel & " AND (Debe <> 0 OR Haber <> 0) " & WhCtaRut
            Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
            Q1 = Q1 & " ORDER BY Codigo  "
         
         End If
      
'         Set RsSaldos = OpenRs(DbMain, Q1, , dbOpenSnapshot)
         Set RsSaldos = OpenRs(DbMain, Q1)
         
         Q1 = ""
         
      End If
   
   Else   'hay año anterior, no se insertan saldos (consulta nula)
   
'      Set RsSaldos = OpenRs(DbMain, "SELECT * FROM Cuentas WHERE 1=0", , dbOpenSnapshot)
      Set RsSaldos = OpenRs(DbMain, "SELECT * FROM Cuentas WHERE 1=0")
   
   End If
   
   
   'If Ch_DocsComp = 0 Then
   TipoInforme = CbItemData(Cb_TipoInforme)
   
   If TipoInforme = TI_TODOS Or TipoInforme = TI_SOLODOCS_SINCOMP Then
   
      'consulta de docs que no están en enlazados a ningún comprobante
      Q1 = Q1 & " SELECT Documento.IdDoc, Documento.NumDoc, Documento.TipoLib, Documento.TotPagadoAnoAnt, "
      Q1 = Q1 & "  Documento.TipoDoc, Documento.IdEntidad, Entidades.Nombre, Entidades.RUT," & SqlInt("Entidades.NotValidRut") & " As NotValidRut, "
      Q1 = Q1 & "  Documento.FEmisionOri, Documento.FVenc, Documento.Total, Documento.Descrip, "
      Q1 = Q1 & "  MovDocumento.IdCuenta, Cuentas.Codigo, Cuentas.Descripcion " & DetComp1 & ","
      Q1 = Q1 & "  Sum(MovDocumento.Debe) As Debe, Sum(MovDocumento.Haber) As Haber, Documento.SaldoDoc "
         
      Q1 = Q1 & " FROM (((Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True) & " )"
      Q1 = Q1 & "  LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
      Q1 = Q1 & "  LEFT JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento") & " )"
      'Q1 = Q1 & "  LEFT JOIN MovComprobante ON Documento.IdDoc = MovComprobante.IdDoc"
      Q1 = Q1 & "  LEFT JOIN vMovCompIdDoc ON Documento.IdDoc = vMovCompIdDoc.IdDoc "
      Q1 = Q1 & JoinEmpAno(gDbType, "vMovCompIdDoc", "Documento")
   
      'OJO: si es ViewOtrosDocs entonces EsTotalDoc siempre es cero, por lo que esta parte del Union del query no entrega nada
      '7 abr 2008: ahra mostramos Otros Documentos aunque no estén asociados a un comprobante (?¿¡! usuarios)
      Q1 = Q1 & " WHERE "
      
      If Not lViewOtrosDocs Then
         Q1 = Q1 & " (EsTotalDoc <> 0 OR EsTotalDoc IS NULL) AND Documento.IdEntidad > 0 AND "   'EsTotalDoc IS NULL para los Otros Documentos, que no tienen movimientos
      End If
      
      'si NO ViewOtrosDocs, incluimos los OtrosDocs que el usuario marca para que se incluyan
      Q1 = Q1 & "  (Documento.TipoLib " & CondOtrosDocs & " IN(" & LIB_OTROFULL & ") "
      Q1 = Q1 & CondDocOtrosEnAnalitico & ")"
      
      Q1 = Q1 & "  AND Documento.Estado <> " & ED_ANULADO
            'tomamos los que no están enlazados a un comprobante y los que están marcados como centralizados pero no tienen comprobante asociado (docs pendientes del año anterior)
      'Q1 = Q1 & "  AND (MovComprobante.IdComp IS NULL "
      Q1 = Q1 & "  AND (vMovCompIdDoc.IdDoc IS NULL "
      Q1 = Q1 & "  OR (Documento.Estado IN (" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND Documento.IdCompCent = 0 ))"
      Q1 = Q1 & Wh & WhDoc
      Q1 = Q1 & " AND Documento.FEmisionOri <= " & GetTxDate(Tx_Hasta)
      Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
     'Q1 = Q1 & " AND MovDocumento.idcuenta >0 " '2748525 - 2751484 OTROS DOCU
   
      Q1 = Q1 & " GROUP BY Documento.IdDoc, Documento.NumDoc, Documento.TipoLib, Documento.TotPagadoAnoAnt, "
      Q1 = Q1 & "  Documento.TipoDoc, Documento.IdEntidad, Entidades.Nombre, Entidades.RUT, " & SqlInt("Entidades.NotValidRut") & ", "
      Q1 = Q1 & "  Documento.FEmisionOri, Documento.FVenc, Documento.Total, Documento.Descrip, "
      Q1 = Q1 & "  MovDocumento.IdCuenta, Cuentas.Codigo, Cuentas.Descripcion, Documento.SaldoDoc "
   
         
   End If
   
   
   Q1 = ""
   If TipoInforme = TI_TODOS Or TipoInforme = TI_SOLODOCS_SINCOMP Then
      
   
      'consulta Docs sin Comprobantes
      Q1 = Q1 & " SELECT iif(MovComprobante.IDDOC IS NULL, Documento.IdDoc, MovComprobante.IDDOC) as IdDoc, Documento.NumDoc, Documento.TipoLib, Documento.TotPagadoAnoAnt, "
      Q1 = Q1 & "  Documento.TipoDoc, "
      
      If Not lViewOtrosDocs Then
         Q1 = Q1 & "  Documento.IdEntidad, "
      Else
         Q1 = Q1 & "  iif( Documento.IdEntidad = 0, -1, Documento.IdEntidad) As IdEntidad, "
      End If
      
      Q1 = Q1 & "  Entidades.Nombre, Entidades.RUT, " & SqlInt("Entidades.NotValidRut") & " As NotValidRut, "
      Q1 = Q1 & "  Documento.FEmisionOri, Documento.FVenc, Documento.Total, Documento.Descrip, "
      Q1 = Q1 & "  Documento.IdCtaBanco AS IdCuenta, Cuentas.Codigo, Cuentas.Descripcion " & DetComp1 & ","
      'Q1 = Q1 & "  Sum(MovComprobante.Debe) As Debe, Sum(MovComprobante.Haber) As Haber "
      
      'Q1 = Q1 & "  iif(Documento.SaldoDoc <> 0 AND  Documento.SaldoDoc IS NOT NULL AND Documento.Tratamiento = 1 , Documento.SaldoDoc, iif(Documento.Tratamiento = 1 AND (Documento.SaldoDoc IS NULL OR Documento.SaldoDoc = 0) ,Sum(MovComprobante.Haber),iif(Documento.Tratamiento = 2 AND (Documento.SaldoDoc IS NULL OR Documento.SaldoDoc = 0) ,Sum(MovComprobante.Debe),0))) As Debe, "
      'Q1 = Q1 & "  iif(Documento.SaldoDoc <> 0 AND  Documento.SaldoDoc IS NOT NULL AND Documento.Tratamiento = 2 , Documento.SaldoDoc, iif(Documento.Tratamiento = 2 AND (Documento.SaldoDoc IS NULL OR Documento.SaldoDoc = 0) ,Sum(MovComprobante.Debe),iif(Documento.Tratamiento = 1 AND (Documento.SaldoDoc IS NULL OR Documento.SaldoDoc = 0) ,Sum(MovComprobante.Haber),0))) As Haber "
      'ffv 2855046
      Q1 = Q1 & "  iif(MovComprobante.Debe = 0 or MovComprobante.Debe is null AND Documento.Tratamiento = 1  ,Total, iif(Documento.Tratamiento = 1 AND MovComprobante.Debe <> 0  ,Sum(MovComprobante.Debe),iif(Documento.Tratamiento = 2 AND MovComprobante.Debe <> 0 ,Sum(MovComprobante.Debe),0))) As Debe, "
      Q1 = Q1 & "  iif(MovComprobante.HABER = 0 or MovComprobante.Haber is null AND Documento.Tratamiento = 2 , Total, iif(Documento.Tratamiento = 2 AND MovComprobante.HABER  <> 0,Sum(MovComprobante.Haber),iif(Documento.Tratamiento = 1 AND MovComprobante.Haber <> 0 ,Sum(MovComprobante.Haber),0))) As Haber "
      'ffv 2855046
      Q1 = Q1 & "  , Documento.SaldoDoc,Documento.Tratamiento, MovComprobante.Glosa "
      
      Q1 = Q1 & " FROM (((Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True) & " )"
      Q1 = Q1 & "  LEFT JOIN MovComprobante ON Documento.IdDoc = MovComprobante.IdDoc "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
      Q1 = Q1 & "  LEFT JOIN Cuentas ON Documento.IdCtaBanco = Cuentas.IdCuenta) "
   
      'Q1 = Q1 & " WHERE Comprobante.Estado <> " & EC_ANULADO ' & " AND Documento.NumDoc=56 "
      Q1 = Q1 & " WHERE Documento.Estado <> " & ED_ANULADO   'para que no muestre docs anulados en el Listado de Otros Documentos
      Q1 = Q1 & "  AND (MovComprobante.IdDoc IS NULL) "
      If Not lViewOtrosDocs Then
         Q1 = Q1 & " AND Documento.IdEntidad > 0  "   'para que se puedan ver los documentos que no tienen entidad asociada en el informe de Otros Docs
      End If
      
      'si NO ViewOtrosDocs, incluimos los OtrosDocs que el usuario marca para que se incluyan
      Q1 = Q1 & " AND (Documento.TipoLib " & CondOtrosDocs & " IN(" & LIB_OTROFULL & ") "
      Q1 = Q1 & CondDocOtrosEnAnalitico & ")"
   
      Q1 = Q1 & Wh & WhComp & WhCtaRut1
      
      Q1 = Q1 & " AND ( Documento.FEmisionOri <= " & GetTxDate(Tx_Hasta)
      
      If lViewOtrosDocs Then    'para tomar los Otros docs que tiene FEmisionOri en NULL (esto ya no debiera ocurrir pero para los docs antiguos)
         Q1 = Q1 & " OR Documento.FEmision <= " & GetTxDate(Tx_Hasta)
      End If
      Q1 = Q1 & ")"
      
      'Q1 = Q1 & " AND Comprobante.Fecha <= " & GetTxDate(Tx_HastaComp)
      Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
     ' Q1 = Q1 & " AND MovComprobante.idcuenta >0 " '2748525 - 2751484 OTROS DOCU
      
      Q1 = Q1 & " GROUP BY MovComprobante.IdDoc, Documento.NumDoc, Documento.TipoLib, Documento.TotPagadoAnoAnt, "
      Q1 = Q1 & "  Documento.TipoDoc, Documento.IdEntidad, Entidades.Nombre, Entidades.RUT, " & SqlInt("Entidades.NotValidRut") & ", "
      Q1 = Q1 & "  Documento.FEmisionOri, Documento.FVenc, Documento.Total, Documento.Descrip, "
      Q1 = Q1 & "  Documento.IdCtaBanco, Cuentas.Codigo, Cuentas.Descripcion, Documento.SaldoDoc, Documento.Tratamiento, Documento.IdDoc, MovComprobante.Glosa " '& DetComp1
      
       'ffv 2855046
      Q1 = Q1 & "  ,MovComprobante.debe,MovComprobante.HABER"
      'ffv 2855046
      
   End If
   
   
   
   If TipoInforme = TI_TODOS Or TipoInforme = TI_SOLODOCS_CONCOMP Then
   
      If Q1 <> "" Then
         Q1 = Q1 & " UNION "
      End If
      
   
      'consulta de movs. comprobantes que tienen docs enlazados
      Q1 = Q1 & " SELECT iif(MovComprobante.IDDOC IS NULL, Documento.IdDoc, MovComprobante.IDDOC) as IdDoc, Documento.NumDoc, Documento.TipoLib, Documento.TotPagadoAnoAnt, "
      Q1 = Q1 & "  Documento.TipoDoc, "
      
      If Not lViewOtrosDocs Then
         Q1 = Q1 & "  Documento.IdEntidad, "
      Else
         Q1 = Q1 & "  iif( Documento.IdEntidad = 0, -1, Documento.IdEntidad) As IdEntidad, "
      End If
      
      Q1 = Q1 & "  Entidades.Nombre, Entidades.RUT, " & SqlInt("Entidades.NotValidRut") & " As NotValidRut, "
      Q1 = Q1 & "  Documento.FEmisionOri, Documento.FVenc, Documento.Total, Documento.Descrip, "
      Q1 = Q1 & "  Documento.IdCtaBanco AS IdCuenta, Cuentas.Codigo, Cuentas.Descripcion " & DetComp2 & ","
      
'       Q1 = Q1 & "  iif((MovComprobante.Debe = 0 or MovComprobante.Debe is null) AND Documento.Tratamiento = 1  ,Total, iif(Documento.Tratamiento = 1 AND MovComprobante.Debe <> 0  ,Sum(MovComprobante.Debe),iif(Documento.Tratamiento = 2 AND MovComprobante.Debe <> 0 ,Sum(MovComprobante.Debe),0))) As Debe, "
'      Q1 = Q1 & "  iif((MovComprobante.HABER = 0 or MovComprobante.Haber is null)  AND Documento.Tratamiento = 2 , Total, iif(Documento.Tratamiento = 2 AND MovComprobante.HABER  <> 0,Sum(MovComprobante.Haber),iif(Documento.Tratamiento = 1 AND MovComprobante.Haber <> 0 ,Sum(MovComprobante.Haber),0))) As Haber, "
      
      Q1 = Q1 & "  Sum(MovComprobante.Debe) As Debe, "
      Q1 = Q1 & "  Sum(MovComprobante.Haber) As Haber, "
      
      'ffv 2855046
      Q1 = Q1 & "  Documento.SaldoDoc,Documento.Tratamiento, MovComprobante.Glosa "
            
      Q1 = Q1 & " FROM (((MovComprobante INNER JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
      Q1 = Q1 & "  INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
      Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
      Q1 = Q1 & "  LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True) & " )"
      Q1 = Q1 & "  LEFT JOIN Cuentas ON Documento.IdCtaBanco = Cuentas.IdCuenta"
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante")
   
      Q1 = Q1 & " WHERE Comprobante.Estado <> " & EC_ANULADO ' & " AND Documento.NumDoc=56 "
      Q1 = Q1 & " AND Documento.Estado <> " & ED_ANULADO   'para que no muestre docs anulados en el Listado de Otros Documentos
      
      If Not lViewOtrosDocs Then
         Q1 = Q1 & " AND Documento.IdEntidad > 0  "   'para que se puedan ver los documentos que no tienen entidad asociada en el informe de Otros Docs
      End If
      
      'si NO ViewOtrosDocs, incluimos los OtrosDocs que el usuario marca para que se incluyan
      Q1 = Q1 & " AND (Documento.TipoLib " & CondOtrosDocs & " IN(" & LIB_OTROFULL & ") "
      Q1 = Q1 & CondDocOtrosEnAnalitico & ")"
   
      Q1 = Q1 & Wh & WhComp & WhCtaRut1
      
      Q1 = Q1 & " AND ( Documento.FEmisionOri <= " & GetTxDate(Tx_Hasta)
      
      If lViewOtrosDocs Then    'para tomar los Otros docs que tiene FEmisionOri en NULL (esto ya no debiera ocurrir pero para los docs antiguos)
         Q1 = Q1 & " OR Documento.FEmision <= " & GetTxDate(Tx_Hasta)
      End If
      Q1 = Q1 & ")"
      
      Q1 = Q1 & " AND Comprobante.Fecha <= " & GetTxDate(Tx_HastaComp)
      Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
     ' Q1 = Q1 & " AND MovComprobante.idcuenta >0 " '2748525 - 2751484 OTROS DOCU
      
      Q1 = Q1 & " GROUP BY MovComprobante.IdDoc, Documento.NumDoc, Documento.TipoLib, Documento.TotPagadoAnoAnt, "
      Q1 = Q1 & "  Documento.TipoDoc, Documento.IdEntidad, Entidades.Nombre, Entidades.RUT, " & SqlInt("Entidades.NotValidRut") & ", "
      Q1 = Q1 & "  Documento.FEmisionOri, Documento.FVenc, Documento.Total, Documento.Descrip, "
      Q1 = Q1 & "  Documento.IdCtaBanco, Cuentas.Codigo, Cuentas.Descripcion, Documento.SaldoDoc, Documento.IdDoc, Documento.Tratamiento " & DetComp2
      'ffv 2855046
      Q1 = Q1 & "  ,MovComprobante.debe,MovComprobante.HABER, MovComprobante.Glosa"
      'ffv 2855046
     
   End If

   'creamos una vista con este query
   Rc = CreateQry(DbMain, QName, Q1)
   'tiramos el resultado del query a una tabla temporal
   Q1 = "SELECT * INTO " & TmpTbl & " FROM " & QName
   Call ExecSQL(DbMain, Q1)
   
   
   If Ch_SaldosVig <> 0 Then   'tmptbl
   
      'seleccionamos en una tabla temporal 2 los que tienen Debe = Haber (Saldo = 0)
      Q1 = " SELECT IdDoc, NumDoc, TipoLib, TotPagadoAnoAnt, TipoDoc, IdEntidad, Nombre, RUT, NotValidRut, FEmisionOri,"
      Q1 = Q1 & " FVenc, Total, Descrip, IdCuenta, Codigo, Descripcion, Sum(Debe) as SumDebe, Sum(Haber) as SumHaber, "
      Q1 = Q1 & " SaldoDoc, Glosa INTO " & TmpTbl2 & " FROM " & TmpTbl
      Q1 = Q1 & " Where SaldoDoc = 0 "
      Q1 = Q1 & " GROUP BY IdDoc, NumDoc, TipoLib, TotPagadoAnoAnt, tipoDoc, IdEntidad, Nombre, RUT, NotValidRut, "
      Q1 = Q1 & " FEmisionOri, FVenc, Total, Descrip, IdCuenta, Codigo, Descripcion, SaldoDoc, Glosa "
      'Q1 = Q1 & " HAVING Sum(Debe) = Sum(Haber)"
      Call ExecSQL(DbMain, Q1)
      
      
      sDelFrom = TmpTbl & " LEFT JOIN " & TmpTbl2 & " ON "
      sDelFrom = sDelFrom & " (" & TmpTbl & ".IdDoc = " & TmpTbl2 & ".IdDoc) AND"
      sDelFrom = sDelFrom & " (" & TmpTbl & ".NumDoc = " & TmpTbl2 & ".NumDoc) AND "
      sDelFrom = sDelFrom & " (" & TmpTbl & ".TipoLib = " & TmpTbl2 & ".TipoLib) AND "
      sDelFrom = sDelFrom & " (" & TmpTbl & ".TipoDoc = " & TmpTbl2 & ".TipoDoc) AND "
      sDelFrom = sDelFrom & " (" & TmpTbl & ".IdEntidad = " & TmpTbl2 & ".IdEntidad) "
      sDelWhere = " WHERE NOT (" & TmpTbl2 & ".IdDoc IS NULL) And NOT (" & TmpTbl2 & ".NumDoc IS NULL)"
      
      Call DeleteJSQL(DbMain, TmpTbl, sDelFrom, sDelWhere)
      
      Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl2)
     
      'seleccionamos en una tabla temporal 2 los que tienen TotPagadoAnoAnt = Debe or TotPagadoAnoAnt = Haber (Saldo = 0)
      Q1 = " SELECT IdDoc, NumDoc, TipoLib, TotPagadoAnoAnt, TipoDoc, IdEntidad, Nombre, RUT, NotValidRut, FEmisionOri,"
      Q1 = Q1 & " FVenc, Total, Descrip, IdCuenta, Codigo, Descripcion, Debe as SumDebe, Haber as SumHaber, "
      Q1 = Q1 & " SaldoDoc, Glosa  INTO " & TmpTbl2 & " FROM " & TmpTbl
      Q1 = Q1 & " WHERE (Haber <> 0 AND Haber = -1 * TotPagadoAnoAnt) OR (Debe <> 0 AND Debe = TotPagadoAnoAnt)"
      Call ExecSQL(DbMain, Q1)
      

      sDelFrom = TmpTbl & " LEFT JOIN " & TmpTbl2 & " ON "
      sDelFrom = sDelFrom & " (" & TmpTbl & ".IdDoc = " & TmpTbl2 & ".IdDoc) AND"
      sDelFrom = sDelFrom & " (" & TmpTbl & ".NumDoc = " & TmpTbl2 & ".NumDoc) AND "
      sDelFrom = sDelFrom & " (" & TmpTbl & ".TipoLib = " & TmpTbl2 & ".TipoLib) AND "
      sDelFrom = sDelFrom & " (" & TmpTbl & ".TipoDoc = " & TmpTbl2 & ".TipoDoc) AND "
      sDelFrom = sDelFrom & " (" & TmpTbl & ".IdEntidad = " & TmpTbl2 & ".IdEntidad) "
      sDelWhere = " WHERE NOT (" & TmpTbl2 & ".IdDoc IS NULL) And NOT (" & TmpTbl2 & ".NumDoc IS NULL)"
      
      Call DeleteJSQL(DbMain, TmpTbl, sDelFrom, sDelWhere)
      
   End If
     
   Q1 = "SELECT * FROM " & TmpTbl
   
   If lPorEntidad Then
      Q1 = Q1 & " ORDER BY Nombre, Rut, FEmisionOri, TipoLib, TipoDoc, NumDoc ,debe,haber desc "
   Else  'por cuenta
      Q1 = Q1 & " ORDER BY Codigo, FEmisionOri, Rut, TipoLib, TipoDoc, NumDoc ,debe,haber desc "
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
  
   i = Grid.FixedRows
   Grid.rows = Grid.FixedRows
   IdEntidad = 0
   IdCuenta = 0
   
   TotDebe = 0
   TotHaber = 0
   TotSaldo = 0
   SubTotDebe = 0
   SubTotHaber = 0
   DocDebe = 0
   DocHaber = 0
   TotDebeParc = 0
   TotHaberParc = 0
      
   InsPagoAnoAnt = False
   MsgInsPagoAnoAnt = False
   DocPagoAnt = 0  'se usa para el caso de informe resumen
   
   If Rs.EOF Then    'no hay docs seleccionados, mostramos todos los saldos de apertura solamente
      
      Grid.rows = Grid.rows + 1
      
      If lPorEntidad Then
      
         If gFunciones.DetSaldoApertura And Not lViewOtrosDocs Then
            Call InsertLstSaldoAp(IdEntidad, 0, i, RsSaldos, "", TotDebe, TotHaber)
         End If

      Else  'Por Cuenta
         If gFunciones.DetSaldoApertura And Not lViewOtrosDocs Then
            Call InsertLstSaldoAp(0, IdCuenta, i, RsSaldos, "", TotDebe, TotHaber)
         End If
      End If
   End If
   
   Do While Rs.EOF = False
   
      If FGrChkMaxSize(Grid) = True Then
         Exit Do
      End If
     
      Grid.rows = Grid.rows + 1
      
'      If Ch_InfRes = 0 Then
         If CurIdDoc <> vFld(Rs("IdDoc")) And InsPagoAnoAnt = True Then
            InsPagoAnoAnt = False
         End If
'      Else    'info resumido
'         If DocPagoAnt <> vFld(Rs("IdDoc")) And InsPagoAnoAnt = True Then
'            InsPagoAnoAnt = False
'         End If
'      End If
      
            'cambio de entidad o cuenta
      If (lPorEntidad And IdEntidad <> vFld(Rs("IdEntidad"))) Or (Not lPorEntidad And IdCuenta <> vFld(Rs("IdCuenta"))) Then
         
         If IdEntidad <> 0 Or IdCuenta <> 0 Then    'la primera vez es cero
                        
            Saldo = SubTotDebe - SubTotHaber
            
            If (Ch_SaldosVig <> 0 And Saldo <> 0) Or (Ch_SaldosVig = 0) Then
         
               If CurIdDoc > 0 And (lPorEntidad And Ch_DetComp <> 0) Then  'ponemos el total doc
               
                  'ponemos el saldo doc
                  Grid.TextMatrix(i, C_GLOSA) = "Saldo Documento"
                  Grid.TextMatrix(i, C_SALDO) = Format(DocDebe - DocHaber, NEGNUMFMT)
                  Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
                  Call FGrSetRowStyle(Grid, i, "FC", vbBlue)
                  Grid.TextMatrix(i, C_FMT) = "B"
                  Grid.rows = Grid.rows + 1
                  i = i + 1
                  DocDebe = 0
                  DocHaber = 0
               End If
            
               If Ch_InfRes = 0 Then
                  Grid.TextMatrix(i, C_DOC) = "TOTAL"
                  Grid.TextMatrix(i, C_IDDOC) = ""
               Else
                  i = i - 1
               End If
               
               
               Grid.TextMatrix(i, C_DEBE) = Format(SubTotDebe, NUMFMT)
               Grid.TextMatrix(i, C_HABER) = Format(SubTotHaber, NUMFMT)
               Grid.TextMatrix(i, C_SALDO) = Format(SubTotDebe - SubTotHaber, NEGNUMFMT)
               
               If Ch_InfRes = 0 Then
                  Grid.TextMatrix(i, C_FMT) = "B"
                  Call FGrSetRowStyle(Grid, i, "B")
               End If
               
               Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
               
               TotDebe = TotDebe + SubTotDebe
               TotHaber = TotHaber + SubTotHaber
               
               SubTotDebe = 0
               SubTotHaber = 0
               
               If Ch_InfRes = 0 Then
                  Grid.rows = Grid.rows + 2
                  i = i + 2
               Else
                  i = i + 1
               End If
               Grid.TextMatrix(i - 2, C_OBLIGATORIA) = "O"
               Grid.TextMatrix(i - 1, C_OBLIGATORIA) = "O"
               If vFmt(Grid.TextMatrix(i - 1, C_SALDO)) = 0 Then
                  Grid.TextMatrix(i - 1, C_SALDO) = ""
               End If
               
            Else     'Ch_SaldosVig <> 0 And Saldo = 0
            
               If Ch_InfRes = 0 Then
                  'retrocedemos hasta encontrar la línea de la entidad anterior (esto en el caso de haber uno o más docuemtnos con saldo 0 y que el total de la entidad da cero también)
                  'OJO que puede ocurrir que una cuenta o una entidad tenga Saldo 0 pero los documentos por si solos no den saldo 0. Igual se oculta la cuenta o entidad completa.
                  For j = i - 1 To Grid.FixedRows Step -1
                     If Grid.TextMatrix(j, C_IDENTIDAD) <> "" And Grid.TextMatrix(j, C_ENTIDAD) <> "" And Grid.TextMatrix(j, C_DOC) <> "" And Grid.TextMatrix(j, C_SUBTITULO) <> "" Then   'se elimina la columna DOC porque en algunos casos produce problemas
                        i = j
                        SubTotDebe = 0
                        SubTotHaber = 0
                        DocDebe = 0
                        DocHaber = 0
                        Exit For
                     Else
                        Grid.TextMatrix(j, C_DOC) = ""
                        Grid.TextMatrix(j, C_FECHADOC) = ""
                        Grid.TextMatrix(j, C_FECHAVENC) = ""
                        Grid.TextMatrix(j, C_GLOSA) = ""
                        Grid.TextMatrix(j, C_DEBE) = ""
                        Grid.TextMatrix(j, C_HABER) = ""
                        Grid.TextMatrix(j, C_SALDO) = ""
                        Grid.TextMatrix(j, C_IDENTIDAD) = ""
                        Grid.TextMatrix(j, C_ENTIDAD) = ""
                        Grid.TextMatrix(j, C_DOC) = ""
                        Call FGrSetRowStyle(Grid, j, "FC", vbBlack)
                        Grid.TextMatrix(j, C_FMT) = ""
                     End If
                  Next j
               Else
                  i = i - 1
                  SubTotDebe = 0
                  SubTotHaber = 0
               End If
               
               Grid.TextMatrix(i, C_SUBTITULO) = ""
               Grid.TextMatrix(i, C_IDENTIDAD) = ""
               Grid.TextMatrix(i, C_ENTIDAD) = ""
               Grid.TextMatrix(i, C_DOC) = ""
               
            End If
         
         End If
         
         If lPorEntidad Then
         
            Grid.TextMatrix(i, C_SALDO) = ""
            
            If gFunciones.DetSaldoApertura And Not lViewOtrosDocs Then
               Call InsertLstSaldoAp(IdEntidad, 0, i, RsSaldos, vFld(Rs("Nombre")), TotDebe, TotHaber)
            End If
            
            If vFld(Rs("IdEntidad")) > 0 Then
               Grid.TextMatrix(i, C_IDENTIDAD) = vFld(Rs("IdEntidad"))
               Grid.TextMatrix(i, C_ENTIDAD) = vFld(Rs("Nombre"), True)
               Grid.TextMatrix(i, C_DOC) = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = 0)
               
            ElseIf vFld(Rs("IdEntidad")) < 0 Then
               Grid.TextMatrix(i, C_IDENTIDAD) = vFld(Rs("IdEntidad"))
               Grid.TextMatrix(i, C_ENTIDAD) = "(Sin Entidad Asociada)"
               Grid.TextMatrix(i, C_DOC) = "(Sin RUT)"
            End If
            
            Grid.TextMatrix(i, C_IDDOC) = ""
            Grid.TextMatrix(i, C_SUBTITULO) = "1"
            
            If Ch_InfRes = 0 Then
               If IdEntidad <> 0 Then
                  Grid.TextMatrix(i, C_FMT) = "LB"
               Else
                  Grid.TextMatrix(i, C_FMT) = "B"
               End If
               Call FGrSetRowStyle(Grid, i, "B")
            End If
            
            Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
            
            IdEntidad = vFld(Rs("IdEntidad"))
            CurIdDoc = 0
            
            If IdEntidad > 0 And gFunciones.DetSaldoApertura And Not lViewOtrosDocs Then
               If InsertSaldoApertura(0, IdEntidad, vFld(Rs("Nombre"), True), i, RsSaldos) = True Then
                  SubTotDebe = vFld(RsSaldos("SumDebe"))
                  SubTotHaber = vFld(RsSaldos("SumHaber"))
                  RsSaldos.MoveNext
               End If
            End If
            
         Else  'Por Cuenta
            Grid.TextMatrix(i, C_SALDO) = ""
            
            If gFunciones.DetSaldoApertura And Not lViewOtrosDocs Then
               Call InsertLstSaldoAp(0, IdCuenta, i, RsSaldos, vFld(Rs("Codigo")), TotDebe, TotHaber)
            End If
            
            Grid.TextMatrix(i, C_IDENTIDAD) = vFld(Rs("IdCuenta"))
            Grid.TextMatrix(i, C_ENTIDAD) = FCase(vFld(Rs("Descripcion"), True))
            Grid.TextMatrix(i, C_DOC) = FmtCodCuenta(vFld(Rs("Codigo")))
            Grid.TextMatrix(i, C_IDDOC) = ""
            Grid.TextMatrix(i, C_SUBTITULO) = "1"
            
            If Ch_InfRes = 0 Then
               If IdCuenta > 0 Then
                  Grid.TextMatrix(i, C_FMT) = "LB"
               Else
                  Grid.TextMatrix(i, C_FMT) = "B"
               End If
               Call FGrSetRowStyle(Grid, i, "B")
            End If
            
            Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
            
            IdCuenta = vFld(Rs("IdCuenta"))
            
            If IdCuenta > 0 And gFunciones.DetSaldoApertura And Not lViewOtrosDocs Then
               If InsertSaldoApertura(IdCuenta, 0, vFld(Rs("Codigo")), i, RsSaldos) = True Then
                  SubTotDebe = vFld(RsSaldos("SumDebe"))
                  SubTotHaber = vFld(RsSaldos("SumHaber"))
                  RsSaldos.MoveNext
               End If
            End If
            
         End If
         
         If FGrChkMaxSize(Grid) = True Then
            Exit Do
         End If
            
         Grid.rows = Grid.rows + 1
         i = i + 1
      End If
                   
      Debe = vFld(Rs("Debe"), True)
      Haber = vFld(Rs("Haber"), True)
      Saldo = Debe - Haber
      
      'detalle doc
      If Ch_InfRes.Value = 0 And ((Ch_SaldosVig <> 0 And Saldo <> 0) Or (Ch_SaldosVig = 0)) Then
         
'                  'Solo para Debugging
'                  If vFld(Rs("Correlativo")) = 69 Or vFld(Rs("Correlativo")) = 72 Then
'                     MsgBeep vbExclamation
'                  End If

         If CurIdDoc <> vFld(Rs("IdDoc")) Or (lPorEntidad And Ch_DetComp <> 0) Then  'si estamos presentando detalle por comprobante, separamos los movimientos de un mismo doc, para ver ambos comprobantes
            
            If CurIdDoc <> vFld(Rs("IdDoc")) And CurIdDoc > 0 And (lPorEntidad And Ch_DetComp <> 0) Then  'ponemos el total doc
               'ponemos el saldo doc
               Grid.TextMatrix(i, C_GLOSA) = "Saldo Documento"
               
               
               
               Grid.TextMatrix(i, C_SALDO) = Format(DocDebe - DocHaber, NEGNUMFMT)
               Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
               Call FGrSetRowStyle(Grid, i, "FC", vbBlue)
               Grid.TextMatrix(i, C_FMT) = "B"
               Grid.rows = Grid.rows + 1
               i = i + 1
               DocDebe = 0
               DocHaber = 0
            End If
            
            'detalle comprobante
            If lPorEntidad And Ch_DetComp <> 0 Then
               If vFld(Rs("IdComp")) > 0 Then
                  Grid.TextMatrix(i, C_IDCOMP) = vFld(Rs("IdComp"))
                  TipoComp = UCase(Left(gTipoComp(vFld(Rs("Tipo"))), 1))
                  If TipoComp = "I" Then
                     TipoComp = " " & TipoComp
                  End If
                  Grid.TextMatrix(i, C_ENTIDAD) = String(4, " ") & TipoComp & " " & vFld(Rs("Correlativo")) & IIf(vFld(Rs("TipoAjuste")) = TAJUSTE_TRIBUTARIO, "-T", "") & String((9 - Len(vFld(Rs("Correlativo")))) * 2, " ") & Format(vFld(Rs("Fecha")), SDATEFMT)
               End If
            End If
            
            Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
            Grid.TextMatrix(i, C_DOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc"))) & " " & vFld(Rs("NumDoc"), True)
            Grid.TextMatrix(i, C_FECHADOC) = Format(vFld(Rs("FEmisionOri")), SDATEFMT)
            If vFld(Rs("FVenc")) > 0 Then
               Grid.TextMatrix(i, C_FECHAVENC) = Format(vFld(Rs("FVenc")), SDATEFMT)
            End If
            Grid.TextMatrix(i, C_GLOSA) = vFld(Rs("Descrip"), True)
            Grid.TextMatrix(i, C_GLOSACOMP) = vFld(Rs("Glosa"), True)
            Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
            
            If Not lPorEntidad Then   'por cuenta
               'Grid.TextMatrix(i, C_ENTIDAD) = Right(String(14, " ") & FmtCID(vFld(Rs("Rut"))), 14) & " " & vFld(Rs("Nombre"), True)
               Grid.TextMatrix(i, C_ENTIDAD) = String(4, " ") & FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = 0) & " " & vFld(Rs("Nombre"), True)
            End If
            
            If Not InsPagoAnoAnt And vFld(Rs("TotPagadoAnoAnt")) <> 0 Then
               If vFld(Rs("TotPagadoAnoAnt")) > 0 Then
                  Haber = Haber + vFld(Rs("TotPagadoAnoAnt"))
               Else
                  Debe = Debe + Abs(vFld(Rs("TotPagadoAnoAnt")))
               End If
               InsPagoAnoAnt = True
               
               If Not MsgInsPagoAnoAnt Then
                  MsgBox1 "Atención: los documentos inluyen el monto pagado en el año anterior.", vbInformation
                  MsgInsPagoAnoAnt = True
               End If

            End If
            
            '            If i = 95 Then
            '               MsgBeep vbExclamation
            '            End If
                        'ERRORRRRRRR saldo erroneo se inserta acá con i=95
            If Ch_DetComp = 0 Then
                 If vFld(Rs("Tratamiento")) = 1 And Debe = 0 Then
                    Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("total")), BL_NUMFMT)
                 Else
                    Grid.TextMatrix(i, C_DEBE) = Format(Debe, BL_NUMFMT)
                 End If
                
                 If vFld(Rs("Tratamiento")) = 2 And Haber = 0 Then
                     Grid.TextMatrix(i, C_HABER) = Format(vFld(Rs("total")), BL_NUMFMT)
                 Else
                     Grid.TextMatrix(i, C_HABER) = Format(Haber, BL_NUMFMT)
                 End If
            
            'Grid.TextMatrix(i, C_HABER) = Format(Haber, BL_NUMFMT)
                 Grid.TextMatrix(i, C_SALDO) = Format(vFmt(Grid.TextMatrix(i, C_DEBE)) - vFmt(Grid.TextMatrix(i, C_HABER)), NEGNUMFMT)
            Else

                Grid.TextMatrix(i, C_DEBE) = Format(Debe, BL_NUMFMT)
                Grid.TextMatrix(i, C_HABER) = Format(Haber, BL_NUMFMT)
                Grid.TextMatrix(i, C_SALDO) = Format(Debe - Haber, NEGNUMFMT)
            End If
            
'            Grid.TextMatrix(i, C_DEBE) = Format(Debe, BL_NUMFMT)
'            Grid.TextMatrix(i, C_HABER) = Format(Haber, BL_NUMFMT)
'            Grid.TextMatrix(i, C_SALDO) = Format(Debe - Haber, NEGNUMFMT)
                       
             'pipe
            'If vFld(Rs("TipoLib")) = 5 And Year(Format(vFld(Rs("FEmisionOri")), SDATEFMT)) = gEmpresa.Ano Then
            If vFld(Rs("TipoLib")) = 5 Or vFld(Rs("TipoLib")) = 4 Then
            
                If Haber = 0 And Debe = 0 Then
                                       
                 Grid.ColWidth(C_SALDOCALC) = 1450
                 Grid.TextMatrix(i, C_SALDOCALC) = Format(vFld(Rs("Total")), NEGNUMFMT)
                End If
             
            End If
      
         Else    'sin detalle comprobante
                     
            i = i - 1
            
            If Not InsPagoAnoAnt And vFld(Rs("TotPagadoAnoAnt")) <> 0 Then
               If vFld(Rs("TotPagadoAnoAnt")) > 0 Then
                  Haber = Haber + vFld(Rs("TotPagadoAnoAnt"))
               Else
                  Debe = Debe + Abs(vFld(Rs("TotPagadoAnoAnt")))
               End If
               InsPagoAnoAnt = True
               
               If Not MsgInsPagoAnoAnt Then
                  MsgBox1 "Atención: los documentos inluyen el monto pagado en el año anterior.", vbInformation
                  MsgInsPagoAnoAnt = True
               End If
   
            End If
            
            If vFld(Rs("Tratamiento")) = 1 And Debe <> 0 And Debe = vFld(Rs("total")) And Haber = 0 Then
            Grid.TextMatrix(i, C_DEBE) = Format(Debe, BL_NUMFMT)
            Grid.TextMatrix(i, C_HABER) = Format(Haber + vFmt(Grid.TextMatrix(i, C_HABER)), BL_NUMFMT)
            ElseIf vFld(Rs("Tratamiento")) = 1 And Debe <> 0 And Haber <> 0 And CurIdDoc <> vFld(Rs("IdDoc")) Then
            Grid.TextMatrix(i, C_DEBE) = Format(Debe + vFmt(Grid.TextMatrix(i, C_DEBE)), BL_NUMFMT)
            Grid.TextMatrix(i, C_HABER) = Format(Haber + vFmt(Grid.TextMatrix(i, C_HABER)), BL_NUMFMT)
             ElseIf vFld(Rs("Tratamiento")) = 1 And Debe = 0 And Haber <> 0 And CurIdDoc = vFld(Rs("IdDoc")) Then
            Grid.TextMatrix(i, C_DEBE) = Format(Debe + vFmt(Grid.TextMatrix(i, C_DEBE)), BL_NUMFMT)
            Grid.TextMatrix(i, C_HABER) = Format(Haber + vFmt(Grid.TextMatrix(i, C_HABER)), BL_NUMFMT)
            End If
            
            If vFld(Rs("Tratamiento")) = 2 And Haber <> 0 And Haber = vFld(Rs("total")) And Debe = 0 Then
             Grid.TextMatrix(i, C_HABER) = Format(Haber, BL_NUMFMT)
             Grid.TextMatrix(i, C_DEBE) = Format(Debe + vFmt(Grid.TextMatrix(i, C_DEBE)), BL_NUMFMT)
            ElseIf vFld(Rs("Tratamiento")) = 2 And Haber <> 0 And Debe <> 0 And CurIdDoc <> vFld(Rs("IdDoc")) Then
            Grid.TextMatrix(i, C_HABER) = Format(Haber + vFmt(Grid.TextMatrix(i, C_HABER)), BL_NUMFMT)
            Grid.TextMatrix(i, C_DEBE) = Format(Debe + vFmt(Grid.TextMatrix(i, C_DEBE)), BL_NUMFMT)
            ElseIf vFld(Rs("Tratamiento")) = 2 And Haber = 0 And Debe <> 0 And CurIdDoc = vFld(Rs("IdDoc")) Then
            Grid.TextMatrix(i, C_HABER) = Format(Haber + vFmt(Grid.TextMatrix(i, C_HABER)), BL_NUMFMT)
            Grid.TextMatrix(i, C_DEBE) = Format(Debe + vFmt(Grid.TextMatrix(i, C_DEBE)), BL_NUMFMT)
            
            End If
            
            'Grid.TextMatrix(i, C_DEBE) = Format(Debe + vFmt(Grid.TextMatrix(i, C_DEBE)), BL_NUMFMT)
            'Grid.TextMatrix(i, C_HABER) = Format(Haber + vFmt(Grid.TextMatrix(i, C_HABER)), BL_NUMFMT)
            Grid.TextMatrix(i, C_SALDO) = Format(vFmt(Grid.TextMatrix(i, C_DEBE)) - vFmt(Grid.TextMatrix(i, C_HABER)), NEGNUMFMT)
            
            Grid.rows = Grid.rows - 1
            
            'ocultamos línea con saldo 0, si corresponde
            '9 Jun 06 NO SE HACE porque genera errores en los totales y oculta más líneas de la cuenta en algunos casos
'            If Ch_SaldosVig <> 0 Then
'               If vFmt(Grid.TextMatrix(i, C_IDDOC)) <> 0 And Grid.TextMatrix(i, C_IDENTIDAD) = "" And vFmt(Grid.TextMatrix(i, C_SALDO)) = 0 Then
'                  Debe = Debe - Grid.TextMatrix(i, C_DEBE)
'                  Haber = Haber - Grid.TextMatrix(i, C_HABER)
'                  Grid.RowHeight(i) = 0
'               End If
'
'            End If
            'CurIdDoc = vFld(Rs("IdDoc"))
         End If
         
         'If Ch_DetComp = 1 Then
           If vFld(Rs("Tratamiento")) = 1 And Debe = 0 Then
              Debe = vFld(Rs("total"))
              If CurIdDoc <> vFld(Rs("IdDoc")) Then
                DocDebe = DocDebe + Debe
'              Else
'                DocDebe = Debe
              End If

               DocHaber = DocHaber + Haber

               'Debe = 0
            ElseIf vFld(Rs("Tratamiento")) = 1 And Debe <> 0 And Debe = vFld(Rs("total")) And Haber = 0 Then
                DocDebe = Debe
               DocHaber = DocHaber + Haber
            Else

'               DocDebe = DocDebe + Debe
'               DocHaber = DocHaber + Haber
            End If

            If vFld(Rs("Tratamiento")) = 2 And Haber = 0 And CurIdDoc <> vFld(Rs("IdDoc")) Then
              Haber = vFld(Rs("total"))
              DocDebe = DocDebe + Debe
              If CurIdDoc <> vFld(Rs("IdDoc")) Then
                DocHaber = DocHaber + Haber
'              Else
'                DocHaber = Haber
              End If
            ElseIf vFld(Rs("Tratamiento")) = 2 And Haber <> 0 Then
             DocDebe = DocDebe + Debe
             DocHaber = DocHaber + Haber
            ElseIf vFld(Rs("Tratamiento")) = 2 And Haber = 0 And CurIdDoc = vFld(Rs("IdDoc")) And Debe <> 0 Then
             DocDebe = DocDebe + Debe
             DocHaber = DocHaber + Haber
             
            End If
         ''Else
         'DocDebe = DocDebe + Debe
         'DocHaber = DocHaber + Haber
         '
         'End If
                              
         AnteIdDoc = CurIdDoc
         CurIdDoc = vFld(Rs("IdDoc"))

         i = i + 1
      Else
         Grid.rows = Grid.rows - 1
         
         If Ch_InfRes = 1 Then
         AnteIdDoc = CurIdDoc
         End If
        CurIdDoc = vFld(Rs("IdDoc"))
         
         If Not InsPagoAnoAnt And vFld(Rs("TotPagadoAnoAnt")) <> 0 Then
            If vFld(Rs("TotPagadoAnoAnt")) > 0 Then
               Haber = Haber + vFld(Rs("TotPagadoAnoAnt"))
            Else
               Debe = Debe + Abs(vFld(Rs("TotPagadoAnoAnt")))
            End If
            InsPagoAnoAnt = True
            
            If Not MsgInsPagoAnoAnt Then
               MsgBox1 "Atención: los documentos inluyen el monto pagado en el año anterior.", vbInformation
               MsgInsPagoAnoAnt = True
            End If
            
            DocPagoAnt = vFld(Rs("IdDoc"))
            
         End If

      End If
         
      If (Ch_SaldosVig <> 0 And Saldo <> 0) Or (Ch_SaldosVig = 0) Then
        If Ch_InfRes = 1 Then
            If vFld(Rs("Tratamiento")) = 1 And Debe = 0 Then
              Debe = vFld(Rs("total"))
            End If

            If vFld(Rs("Tratamiento")) = 2 And Haber = 0 And Debe = 0 Then 'And Debe = 0 Then
              Haber = vFld(Rs("total"))
            ElseIf vFld(Rs("Tratamiento")) = 2 And Haber = 0 And Debe <> 0 Then 'And Debe = 0 Then
              Haber = vFld(Rs("total"))
            End If
        
         End If
         
'          If vFld(Rs("Tratamiento")) = 1 And AnteIdDoc = vFld(Rs("IdDoc")) And Debe <> 0 And Haber <> 0 Then
'            SubTotDebe = Debe
'            SubTotHaber = SubTotHaber + Haber
'          ElseIf vFld(Rs("Tratamiento")) = 1 And AnteIdDoc = vFld(Rs("IdDoc")) And Haber <> 0 Then
'            SubTotDebe = SubTotDebe + Debe
'            SubTotHaber = SubTotHaber + Haber
'          ElseIf vFld(Rs("Tratamiento")) = 1 And AnteIdDoc <> vFld(Rs("IdDoc")) And Haber <> 0 Then
'            SubTotDebe = SubTotDebe + Debe
'            SubTotHaber = SubTotHaber + Haber
'          ElseIf vFld(Rs("Tratamiento")) = 1 And AnteIdDoc <> vFld(Rs("IdDoc")) And Haber = 0 And Debe <> 0 Then
'            SubTotDebe = SubTotDebe + Debe
'            SubTotHaber = SubTotHaber + Haber
'          ElseIf vFld(Rs("Tratamiento")) = 2 And AnteIdDoc = vFld(Rs("IdDoc")) And Haber <> 0 And Debe <> 0 Then
'            SubTotDebe = SubTotDebe + Debe
'            SubTotHaber = SubTotHaber + Haber
'          ElseIf vFld(Rs("Tratamiento")) = 2 And AnteIdDoc = vFld(Rs("IdDoc")) And Debe <> 0 Then
'          SubTotDebe = SubTotDebe + Debe
'          SubTotHaber = SubTotHaber + Haber
'          ElseIf vFld(Rs("Tratamiento")) = 2 And AnteIdDoc <> vFld(Rs("IdDoc")) And Debe <> 0 Then
'          SubTotDebe = SubTotDebe + Debe
'          SubTotHaber = SubTotHaber + Haber
'         If vFld(Rs("Tratamiento")) = 2 And AnteIdDoc <> vFld(Rs("IdDoc")) And Debe = 0 And Haber <> 0 Then
'            SubTotDebe = SubTotDebe + Debe
'            SubTotHaber = SubTotHaber + Haber + vFld(Rs("IdDoc"))
'          ElseIf vFld(Rs("Tratamiento")) = 1 And AnteIdDoc <> vFld(Rs("IdDoc")) And Debe <> 0 And Haber = 0 Then
'            SubTotDebe = SubTotDebe + Debe + vFld(Rs("IdDoc"))
'            SubTotHaber = SubTotHaber + Haber
'          Else
            SubTotDebe = SubTotDebe + Debe
            SubTotHaber = SubTotHaber + Haber
'          End If
         
         
         'End If
      
      
        
      End If
       
      Rs.MoveNext
      
   Loop
   
   'ponemos el último total
   If IdEntidad <> 0 Or IdCuenta <> 0 Then
      
      If CurIdDoc > 0 And lPorEntidad And Ch_DetComp <> 0 Then  'ponemos el total doc
      
         'ponemos el saldo doc
         Grid.rows = Grid.rows + 1
         Grid.TextMatrix(i, C_GLOSA) = "Saldo Documento"
         Grid.TextMatrix(i, C_SALDO) = Format(DocDebe - DocHaber, NEGNUMFMT)
         Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
         Call FGrSetRowStyle(Grid, i, "FC", vbBlue)
         Grid.TextMatrix(i, C_FMT) = "B"
         Grid.rows = Grid.rows + 1
         i = i + 1
         DocDebe = 0
         DocHaber = 0

      End If
            
      If Ch_InfRes = 0 Then
         Grid.rows = Grid.rows + 1
         Grid.TextMatrix(i, C_DOC) = "TOTAL"
      Else
         i = i - 1
      End If
      
'      If Ch_InfRes = 1 Then
      Grid.TextMatrix(i, C_DEBE) = Format(SubTotDebe, NUMFMT)
      Grid.TextMatrix(i, C_HABER) = Format(SubTotHaber, NUMFMT)
      Grid.TextMatrix(i, C_SALDO) = Format(SubTotDebe - SubTotHaber, NEGNUMFMT)
'      Else
'
'      Grid.TextMatrix(i, C_DEBE) = Format(DocDebe, NUMFMT)
'      Grid.TextMatrix(i, C_HABER) = Format(DocHaber, NUMFMT)
'      Grid.TextMatrix(i, C_SALDO) = Format(DocDebe - DocHaber, NEGNUMFMT)
'      End If
      
      If Ch_InfRes = 0 Then
         Grid.TextMatrix(i, C_FMT) = "B"
         Call FGrSetRowStyle(Grid, i, "B")
      End If
      
      Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
      
'      TotDebe = TotDebe + SubTotDebe
'      TotHaber = TotHaber + SubTotHaber
'      If Ch_InfRes = 0 Then
'        TotDebe = TotDebe + DocDebe
'        TotHaber = TotHaber + DocHaber
'      Else
        TotDebe = TotDebe + SubTotDebe
        TotHaber = TotHaber + SubTotHaber
'      End If
      
      If gFunciones.DetSaldoApertura Then   'ponemos los últimos saldos de apertura, si quedan
         
         Grid.rows = Grid.rows + 1
         Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
         Grid.rows = Grid.rows + 1
         Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
         
         If Not lViewOtrosDocs Then
            If lPorEntidad Then
               Call InsertLstSaldoAp(IdEntidad, 0, i + 2, RsSaldos, "", TotDebe, TotHaber)
   
            Else  'Por Cuenta
               Call InsertLstSaldoAp(0, IdCuenta, i + 2, RsSaldos, "", TotDebe, TotHaber)
            End If
         End If
      
      End If
      
   End If
   
   Call CloseRs(Rs)
   
   If gFunciones.DetSaldoApertura Then
      Call CloseRs(RsSaldos)
   End If
   
   'vemos si quedó algún saldo en 0 y está seleccionado Ch_SaldoVig, para ocultarlo
   RowTit = 0
   If Ch_SaldosVig <> 0 Then
      For i = Grid.FixedRows To Grid.rows - 1
         '8 jun 2006 ahora se hace arriba, uno por uno, para calcular bien los totales
'         If vFmt(Grid.TextMatrix(i, C_IDDOC)) <> 0 And Grid.TextMatrix(i, C_IDENTIDAD) = "" And vFmt(Grid.TextMatrix(i, C_SALDO)) = 0 Then
'            Grid.RowHeight(i) = 0
'         End If

         'si una cuenta tiene muchos detalles, esto no funciona
'         If Grid.TextMatrix(i, C_DOC) = "TOTAL" And vFmt(Grid.TextMatrix(i, C_SALDO)) = 0 Then
'            Grid.RowHeight(i - 1) = 0
'            Grid.RowHeight(i) = 0
'            If Grid.Rows > i + 1 Then
'               Grid.RowHeight(i + 1) = 0
'            End If
'         End If

         If Grid.RowHeight(i) > 0 Then
            If Grid.TextMatrix(i, C_SUBTITULO) <> "" Then
               RowTit = i
            End If

            If Grid.TextMatrix(i, C_DOC) = "TOTAL" And vFmt(Grid.TextMatrix(i, C_SALDO)) = 0 And RowTit > 0 Then    'FCA 25/7/2017
               For j = RowTit To i
                  Grid.RowHeight(j) = 0
               Next j
               If Grid.rows > i + 1 Then
                  Grid.RowHeight(i + 1) = 0
               End If
               RowTit = 0
            End If
         End If
      Next i
   End If
   
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl2)
   Call ExecSQL(DbMain, "DROP TABLE " & QName)


 'SF 14465862
    Call SetUpGridTotal
 'SF 14465862

   'ponemos los totales generales
   GridTot.TextMatrix(0, C_DOC) = "TOTAL"
   GridTot.TextMatrix(0, C_DEBE) = Format(TotDebe, NUMFMT)
   GridTot.TextMatrix(0, C_HABER) = Format(TotHaber, NUMFMT)
   GridTot.TextMatrix(0, C_SALDO) = Format(TotDebe - TotHaber, NEGNUMFMT)
      
   Call FGrVRows(Grid)
   Grid.rows = Grid.rows + 1
   Grid.TopRow = Grid.FixedRows
   
   Grid.Row = Grid.FixedRows
   Grid.RowSel = Grid.Row
   Grid.Col = C_ENTIDAD
   Grid.ColSel = Grid.Col
   
   Grid.Redraw = True
   
   Call EnableFrm(False)
   
   
End Function
Private Sub EnableFrm(bool As Boolean)
   Bt_Search.Enabled = bool
'   bt_Print.Enabled = Not bool
'   Bt_Preview.Enabled = Not bool
'   Bt_CopyExcel.Enabled = Not bool
   
End Sub
Private Sub Tx_Rut_Change()
   Call EnableFrm(True)
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

Private Sub Tx_Rut_LostFocus()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdEnt As Long
   Dim i As Integer
   Dim AuxRut As String

   If Tx_Rut = "" Then
      Cb_Nombre.Clear
      Exit Sub
   End If
   
'   If Not MsgValidCID(Tx_Rut) Then
'      Tx_Rut.SetFocus
'      Exit Sub
'
'   End If
      
   Q1 = "SELECT IdEntidad, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5 FROM Entidades WHERE Rut = '" & vFmtCID(Tx_Rut, Ch_Rut <> 0) & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   IdEnt = 0
   
   If Rs.EOF = False Then   'existe
      IdEnt = vFld(Rs("IdEntidad"))
      
      'seleccionamos el tipo de entidad y esto llena la lista de nombres de entidades
      If vFld(Rs("Clasif" & ENT_CLIENTE)) <> 0 Then
         Call SelCbEntidad(ENT_CLIENTE)
      ElseIf vFld(Rs("Clasif" & ENT_PROVEEDOR)) <> 0 Then
         Call SelCbEntidad(ENT_PROVEEDOR)
      Else
         Call SelCbEntidad(-1)
      End If
      
      'seleccionamos la entidad
      For i = 0 To Cb_Nombre.ListCount - 1
         If lcbNombre.Matrix(M_IDENTIDAD, i) = IdEnt Then
            lcbNombre.ListIndex = i
            Exit For
         End If
      Next i
      
      Call EnableFrm(True)

   Else
      MsgBox1 "Este RUT no ha sido ingresado al sistema.", vbExclamation + vbOKOnly
      Cb_Nombre.Clear
      
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
   
   Tx_Rut = ""
   
   If lcbNombre.ListIndex > 0 Then
      Tx_Rut = FmtCID(lcbNombre.Matrix(M_RUT), Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0)
      Ch_Rut = IIf(Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0, 1, 0)
   End If
   
   Call EnableFrm(True)

End Sub
Private Sub Cb_Cuenta_Click()
      
   Cb_Nombre.Clear
   
   If ItemData(Cb_Cuenta) > 0 Then
      
      If InStr(LCase(Cb_Cuenta), "cliente") > 0 Then
         Call SelCbEntidad(ENT_CLIENTE)
      ElseIf InStr(LCase(Cb_Cuenta), "proveedor") > 0 Then
         Call SelCbEntidad(ENT_PROVEEDOR)
      Else
         Call SelCbEntidad(-1)
      End If
      
   Else
      Tx_Rut = ""
   End If
   
   Call EnableFrm(True)

End Sub
Private Sub SelCbEntidad(Clasif As Integer)
   Dim Q1 As String
   
   lcbNombre.Clear
   
   Q1 = "SELECT Nombre, idEntidad, Rut, abs(NotValidRut) FROM Entidades"
   
   If Clasif >= 0 Then
      Q1 = Q1 & " WHERE Clasif" & Clasif & "=" & CON_CLASIF
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Else
      Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
   
   End If
      
   Q1 = Q1 & " ORDER BY Nombre "
   
   Call lcbNombre.AddItem(" ", 0)
   Call lcbNombre.FillCombo(DbMain, Q1, -1)

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
Private Function InsertSaldoApertura(ByVal IdCuenta As Long, ByVal IdEntidad As Long, NombCod As String, Row As Integer, RsSaldos As Recordset) As Boolean
      
   InsertSaldoApertura = False
   
   'RsSaldos.MoveFirst
   
      
   Do While RsSaldos.EOF = False
   
      If IdCuenta > 0 Then
      
         'If IdCuenta = vFld(RsSaldos("IdCuenta")) Then
         If NombCod = vFld(RsSaldos("Codigo")) Then
            
            Call FillSaldoApertura(Row, vFld(RsSaldos("SumDebe")), vFld(RsSaldos("SumHaber")))
            
            InsertSaldoApertura = True
                        
            Exit Function
         
         'ElseIf IdCuenta < vFld(RsSaldos("IdCuenta")) Then   'no está la cuenta
         ElseIf NombCod < vFld(RsSaldos("Codigo")) Then   'no está la cuenta
            
            Call FillSaldoApertura(Row, 0, 0)
                     
            Exit Function
         
         Else
            RsSaldos.MoveNext
            
         End If
         
      ElseIf IdEntidad > 0 Then
      
         'If IdEntidad = vFld(RsSaldos("IdEntidad")) Then
         If LCase(NombCod) = LCase(vFld(RsSaldos("Nombre"))) Then
            
            Call FillSaldoApertura(Row, vFld(RsSaldos("SumDebe")), vFld(RsSaldos("SumHaber")))
            
            InsertSaldoApertura = True
            
            Exit Function
         
         'ElseIf IdEntidad < vFld(RsSaldos("IdEntidad")) Then   'no está la entidad
         ElseIf LCase(NombCod) < LCase(vFld(RsSaldos("Nombre"))) Then   'no está la entidad
            
            Call FillSaldoApertura(Row, 0, 0)
                     
            Exit Function
         
         Else
            RsSaldos.MoveNext
            
         End If
      
      End If
      
   Loop
    
End Function

Private Sub FillSaldoApertura(Row As Integer, ByVal Debe As Double, Haber As Double)

   If Debe = 0 And Haber = 0 Then
      Exit Sub
   End If

   Row = Row + 1
   Grid.rows = Row + 1
   
   Grid.TextMatrix(Row, C_GLOSA) = "Saldo Apertura"
   Grid.TextMatrix(Row, C_DEBE) = Format(Debe, NEGNUMFMT)
   Grid.TextMatrix(Row, C_HABER) = Format(Haber, NEGNUMFMT)
   Grid.TextMatrix(Row, C_SALDO) = Format(Debe - Haber, NEGNUMFMT)
   'Grid.TextMatrix(Row, C_FMT) = Grid.TextMatrix(Row, C_FMT) & "B"
   Grid.TextMatrix(Row, C_OBLIGATORIA) = "    O"
   'Call FGrSetRowStyle(Grid, Row, "B")
   
End Sub
'inserta todos los saldos apertura de las cuentas o entidades intermedias que no tienen docs, hasta una cuenta o entidad especificada
Private Sub InsertLstSaldoAp(IdEntidad As Long, IdCuenta As Long, Row As Integer, RsSaldos As Recordset, ByVal NombCodHasta As String, TotDebe As Double, TotHaber As Double)
   Dim SetLine As Boolean
   
   If IdEntidad <> 0 Or IdCuenta <> 0 Then
      SetLine = True
   End If

   Do While RsSaldos.EOF = False
   
      If vFld(RsSaldos("SumDebe")) - vFld(RsSaldos("SumHaber")) <> 0 Then
         If lPorEntidad Then
            If NombCodHasta <> "" And LCase(vFld(RsSaldos("Nombre"))) >= LCase(NombCodHasta) Then
               Exit Sub
            End If
            
            Grid.TextMatrix(Row, C_IDENTIDAD) = vFld(RsSaldos("IdEntidad"))
            Grid.TextMatrix(Row, C_ENTIDAD) = vFld(RsSaldos("Nombre"), True)
            Grid.TextMatrix(Row, C_DOC) = FmtCID(vFld(RsSaldos("Rut")), vFld(RsSaldos("NotValidRut")) = 0)
            Grid.TextMatrix(Row, C_SUBTITULO) = 1
         
         Else
            If NombCodHasta <> "" And vFld(RsSaldos("Codigo")) >= NombCodHasta Then
               Exit Sub
            End If
            
            Grid.TextMatrix(Row, C_IDENTIDAD) = vFld(RsSaldos("IdCuenta"))
            Grid.TextMatrix(Row, C_ENTIDAD) = FCase(vFld(RsSaldos("Descripcion"), True))
            Grid.TextMatrix(Row, C_DOC) = FmtCodCuenta(vFld(RsSaldos("Codigo")))
            Grid.TextMatrix(Row, C_SUBTITULO) = 1
         
         End If
         
         If Ch_InfRes = 0 Then
            If SetLine <> 0 Then
               Grid.TextMatrix(Row, C_FMT) = "LB"
            Else
               Grid.TextMatrix(Row, C_FMT) = "B"
            End If
            SetLine = True
            If lPorEntidad Then
               If IdEntidad = 0 Then
                  IdEntidad = -2
               End If
            ElseIf IdCuenta = 0 Then
               IdCuenta = -2
            End If
            
            Call FGrSetRowStyle(Grid, Row, "B")
         End If
               
         Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
   
         Call FillSaldoApertura(Row, vFld(RsSaldos("SumDebe")), vFld(RsSaldos("SumHaber")))
         
         If Ch_InfRes = 0 Then
            Row = Row + 1
            Grid.rows = Row + 1
            Grid.TextMatrix(Row, C_DOC) = "TOTAL"
            Grid.TextMatrix(Row, C_DEBE) = Format(vFld(RsSaldos("SumDebe")), NUMFMT)
            Grid.TextMatrix(Row, C_HABER) = Format(vFld(RsSaldos("SumHaber")), NUMFMT)
            Grid.TextMatrix(Row, C_SALDO) = Format(vFld(RsSaldos("SumDebe")) - vFld(RsSaldos("SumHaber")), NEGNUMFMT)
            Grid.TextMatrix(Row, C_FMT) = "B"
            Call FGrSetRowStyle(Grid, Row, "B")
            Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
         End If
               
         TotDebe = TotDebe + vFld(RsSaldos("SumDebe"))
         TotHaber = TotHaber + vFld(RsSaldos("SumHaber"))
         
         Row = Row + 1
         Grid.rows = Row + 1
         Grid.TextMatrix(Row, C_OBLIGATORIA) = "O"
         Row = Row + 1
         Grid.rows = Row + 1
         
      End If
      
      RsSaldos.MoveNext
      
   Loop

End Sub

Private Sub RecalcSubTot(ByVal RowInit As Integer, ByVal RowEnd As Integer, SubTotDebe As Double, SubTotHaber As Double)
   Dim i As Integer
   Dim TotDebe As Double
   Dim TotHaber As Double
   
   For i = RowInit To RowEnd
   
      If Grid.RowHeight(i) > 0 Then
         TotDebe = TotDebe + vFmt(Grid.TextMatrix(i, C_DEBE))
         TotHaber = TotHaber + vFmt(Grid.TextMatrix(i, C_HABER))
      End If
      
   Next i
   
   SubTotDebe = TotDebe
   SubTotHaber = TotHaber
   
End Sub


Private Sub recalcularSaldos()
Dim WhLib As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim n As Long
   
   
   WhLib = " Documento.TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & "," & LIB_REMU & "," & LIB_OTROS & ") "

   Set Rs = OpenRs(DbMain, "SELECT Count(*) FROM Documento WHERE " & WhLib & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   If Not Rs.EOF Then
      n = vFld(Rs(0))
   End If
   Call CloseRs(Rs)
   
'   If MsgBox1("Esta operación recalcula los saldos de TODOS los documentos (" & n & " documentos)." & vbNewLine & vbNewLine & "Puede tomar un poco de tiempo ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'      Exit Sub
'   End If
   
   Me.MousePointer = vbHourglass
   
   'asignamos SaldoDoc = NULL para los docs de compras, ventas y retenciones para que los recalcule TOTOS
   Call ExecSQL(DbMain, "UPDATE Documento SET SaldoDoc = NULL WHERE " & WhLib & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   
   'asignamos SaldoDoc = Total para Otros Documentos, por el error que se asignaron todos en NULL
'   WhLib = " Documento.TipoLib =" & LIB_OTROS
'   Call ExecSQL(DbMain, "UPDATE Documento SET SaldoDoc = Total WHERE " & WhLib & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
   
   Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano)
   Call RecalcSaldosFulle(gEmpresa.id, gEmpresa.Ano)
    
   Me.MousePointer = vbDefault
   
   'MsgBox1 "Cálculo de saldos vigentes finalizado.", vbInformation
   

End Sub

Private Sub SetUpGridTotal()
   Dim Col As Integer
   
   GridTot.Cols = NCOLS + 1
   
   GridTot.ColWidth(C_IDENTIDAD) = 0
   GridTot.ColWidth(C_ENTIDAD) = 0
   GridTot.ColWidth(C_IDCOMP) = 0
   GridTot.ColWidth(C_IDDOC) = 0
   GridTot.ColWidth(C_DOC) = 1300
   GridTot.ColWidth(C_FECHADOC) = 800
   GridTot.ColWidth(C_FECHAVENC) = 800
   GridTot.ColWidth(C_GLOSA) = 2300
   GridTot.ColWidth(C_GLOSACOMP) = 2300
   GridTot.ColWidth(C_DEBE) = 1300
   GridTot.ColWidth(C_HABER) = 1300
   GridTot.ColWidth(C_SALDO) = 1350
   GridTot.ColWidth(C_SALDOCALC) = 0 '1450
   GridTot.ColWidth(C_FMT) = 0
   GridTot.ColWidth(C_SUBTITULO) = 0
   GridTot.ColWidth(C_OBLIGATORIA) = 0
   
   GridTot.ColAlignment(C_ENTIDAD) = flexAlignLeftCenter
   GridTot.ColAlignment(C_DOC) = flexAlignLeftCenter
   GridTot.ColAlignment(C_FECHADOC) = flexAlignLeftCenter
   GridTot.ColAlignment(C_FECHAVENC) = flexAlignLeftCenter
   GridTot.ColAlignment(C_GLOSA) = flexAlignLeftCenter
   GridTot.ColAlignment(C_GLOSACOMP) = flexAlignLeftCenter
   GridTot.ColAlignment(C_DEBE) = flexAlignRightCenter
   GridTot.ColAlignment(C_HABER) = flexAlignRightCenter
   GridTot.ColAlignment(C_SALDO) = flexAlignRightCenter
   GridTot.ColAlignment(C_SALDOCALC) = flexAlignRightCenter
      
End Sub

