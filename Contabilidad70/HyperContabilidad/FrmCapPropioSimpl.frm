VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmCapPropioSimpl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capital Propio Tributario Simplificado - Artículo 14 D Ley de Renta"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   6315
      Left            =   480
      TabIndex        =   0
      Top             =   900
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   11139
      Cols            =   2
      Rows            =   2
      FixedCols       =   0
      FixedRows       =   0
      ScrollBars      =   2
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.TextBox Tx_CapPropio 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Capital Propio Tributario Simplificado"
      Top             =   7380
      Width           =   6075
   End
   Begin VB.TextBox Tx_TotCapPropio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   7380
      Width           =   1575
   End
   Begin VB.Frame Fr_Notas 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   420
      TabIndex        =   12
      Top             =   7920
      Width           =   7755
      Begin VB.Label Label1 
         Caption         =   $"FrmCapPropioSimpl.frx":0000
         ForeColor       =   &H00FF0000&
         Height          =   435
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   7635
      End
      Begin VB.Label Label1 
         Caption         =   "Recuerde presionar el botón Aceptar en cada ventana de detalle, para que el total sea traspasado a este reporte."
         ForeColor       =   &H00FF0000&
         Height          =   435
         Index           =   1
         Left            =   60
         TabIndex        =   13
         Top             =   600
         Width           =   7575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   60
      Width           =   9015
      Begin VB.CommandButton Bt_Manual 
         Caption         =   "Manual CPT Simplificado"
         Height          =   315
         Left            =   4980
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   2055
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
         Left            =   1500
         Picture         =   "FrmCapPropioSimpl.frx":0097
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   2460
         Picture         =   "FrmCapPropioSimpl.frx":013B
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   2040
         Picture         =   "FrmCapPropioSimpl.frx":049C
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   2880
         Picture         =   "FrmCapPropioSimpl.frx":083A
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   960
         Picture         =   "FrmCapPropioSimpl.frx":0C63
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "FrmCapPropioSimpl.frx":10A8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   7500
         TabIndex        =   9
         Top             =   180
         Width           =   1035
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
         Picture         =   "FrmCapPropioSimpl.frx":154F
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6330
      Left            =   480
      TabIndex        =   10
      Top             =   900
      Visible         =   0   'False
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   11165
      _Version        =   393216
      Rows            =   10
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483626
      GridColor       =   -2147483626
      GridLinesFixed  =   1
      ScrollBars      =   2
   End
End
Attribute VB_Name = "FrmCapPropioSimpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_TITULO = 0
Const C_MONTO = 1
Const C_SIGNO = 2
Const C_FMT = 3
Const C_OBLIGATORIA = 4

Const NCOLS = C_OBLIGATORIA

Dim lTipoInforme As Integer

Dim lRowCapAportado As Integer
Dim lRowBaseImp As Integer
Dim lRowParticipaciones As Integer
Dim lRowDisminucionesCapital As Integer
Dim lRowGastosRechazados As Integer
Dim lRowRetDiv As Integer
Dim lRowCapPropioSimplificado As Integer
Dim lRowAumentosCapital As Integer
Dim lRowINRPropios As Integer
Dim lRowINRPropiosPerdidas As Integer
Dim lRowOtrosAjustesAumentos As Integer
Dim lRowOtrosAjustesDisminuciones As Integer
Dim lRowCapPropioTribAnoAnt As Integer
Dim lRowRepPerdidaArrastre As Integer
Dim lRowUtilidadesPerdida As Integer
Dim lRowIngresoDiferido As Integer
Dim lRowCTDImputableIPE As Integer
Dim lRowIncentivoAhorro As Integer
Dim lRowIDPCVoluntario As Integer
Dim lRowCredActFijos As Integer
Dim lRowCredParticipaciones As Integer

Dim lTotalCapPropioSimplificado As Double

Public Sub FView(ByVal TipoInforme As Integer)

   lTipoInforme = TipoInforme
   
   Me.Show vbModeless

End Sub

Private Sub bt_Cerrar_Click()
   Dim Q1 As String
   Dim sSet As String, sFrom As String, sWhere As String

   sFrom = " EmpresasAno "
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      sSet = "  CPS_CapPropioSimplificado = " & lTotalCapPropioSimplificado
   Else
      sSet = "  CPS_CapPropioSimplVarAnual = " & lTotalCapPropioSimplificado
   End If
   If gEmpresa.ProPymeGeneral Then
      sSet = sSet & ", CPS_BaseImpPrimCat_14DN3 = " & vFmt(Grid.TextMatrix(lRowBaseImp, C_MONTO))
      sSet = sSet & ", CPS_BaseImpPrimCat_14DN8 = 0"
   Else
      sSet = sSet & ", CPS_BaseImpPrimCat_14DN3 = 0"
      sSet = sSet & ", CPS_BaseImpPrimCat_14DN8 = " & vFmt(Grid.TextMatrix(lRowBaseImp, C_MONTO))
   End If
   If lTipoInforme = CPS_TIPOINFO_VARANUAL Then
      sSet = sSet & ", CPS_CapPropioTribAnoAnt = " & vFmt(Grid.TextMatrix(lRowCapPropioTribAnoAnt, C_MONTO))
      sSet = sSet & ", CPS_RepPerdidaArrastre = " & vFmt(Grid.TextMatrix(lRowRepPerdidaArrastre, C_MONTO))
   End If
   sWhere = " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

   Call UpdateSQL(DbMain, "EmpresasAno", sSet, sFrom, sWhere)

   Unload Me
End Sub

Private Sub Bt_Manual_Click()
   Dim Rc As Long
   Dim Buf As String
   
   MousePointer = vbHourglass
   DoEvents
   
   Buf = gAppPath & "\Manual_CPT_Simplificado.pdf"
   Rc = ExistFile(Buf)
      
   If Rc = 0 Then
      MsgBox1 "No se encontró el archivo que contiene el Manual de Capital Propio Tributario Simplificado, por favor contáctese con su proveedor para obtenerlo.", vbExclamation
   Else

      Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
      If Rc < 32 Then
         MsgBox1 "Error " & Rc & " al abrir el archivo '" & Buf & "' que contiene el Manual de Capital Propio Tributario Simplificado." & vbLf & "Trate de abrir este archivo con otro programa.", vbExclamation
      End If
   End If

   MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Dim Msg As String

   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Me.Caption = "Capital Propio Tributario Simplificado General - Artículo 14 D Ley de Renta"
   Else
      Me.Caption = "Capital Propio Tributario Simplificado por Variación Anual - Artículo 14 D Ley de Renta"
   End If

   Call SetUpGrid
   
   Call LoadAll
   
   MsgBox1 "ATENCIÓN:" & vbNewLine & vbNewLine & "Este informe se genera seleccionando solamente los comprobantes en estado APROBADO.", vbInformation
   
   If gEmpresa.ProPymeTransp Then
      
      Msg = Val(GetIniString(gIniFile, "Msg", "14DN8Ingresos", "0"))

      If Msg = 0 Then
         MsgBox1 "Si la empresa se encuentra acogida al Art. 14 D N° 8 y tiene ingresos que no superan 50.000 UF, no esta obligada a efectuar el CPT Simplificado." & vbCrLf & vbCrLf & "Queda a su criterio utilizar este informe.", vbInformation
         Call SetIniString(gIniFile, "Msg", "14DN8Ingresos", "1")
      End If
   End If
   
End Sub

Private Function SetUpGrid()

   Grid.Cols = NCOLS + 1
   Grid.rows = 40
   
   Call FGrSetup(Grid)
   
   Grid.ColAlignment(C_MONTO) = flexAlignRightCenter
   Grid.ColAlignment(C_SIGNO) = flexAlignCenterCenter
   
   Grid.ColWidth(C_TITULO) = 5600
   Grid.ColWidth(C_MONTO) = 1500
   Grid.ColWidth(C_SIGNO) = 400
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_OBLIGATORIA) = 0
   
   Grid.FlxGrid.BackColorFixed = &H8000000F
   Grid.FlxGrid.BackColor = &H80000016
   Grid.FlxGrid.GridColor = &H80000016
   
End Function

Private Function LoadAll()
   Dim Row As Integer, i As Integer
   Dim Q1 As String
   Dim Rs As Recordset, Rs2 As Recordset
   
   
   lRowCapAportado = 0
   lRowBaseImp = 0
   lRowBaseImp = 0
   lRowParticipaciones = 0
   lRowDisminucionesCapital = 0
   lRowGastosRechazados = 0
   lRowRetDiv = 0
   lRowCapPropioSimplificado = 0
   lRowAumentosCapital = 0
   lRowINRPropios = 0
   lRowINRPropiosPerdidas = 0
   lRowOtrosAjustesAumentos = 0
   lRowOtrosAjustesDisminuciones = 0
   lRowCapPropioTribAnoAnt = 0
   lRowRepPerdidaArrastre = 0
   lRowIngresoDiferido = 0
   lRowCTDImputableIPE = 0
   lRowIncentivoAhorro = 0
   lRowIDPCVoluntario = 0
   lRowCredActFijos = 0
   lRowCredParticipaciones = 0
   
   Q1 = "SELECT CPS_CapPropioTribAnoAnt, CPS_CapitalAportado, CPS_BaseImpPrimCat_14DN3, CPS_BaseImpPrimCat_14DN8, CPS_Participaciones, CPS_Disminuciones, CPS_GastosRechazados,"
   Q1 = Q1 & " CPS_RetirosDividendos, CPS_RepPerdidaArrastre, CPS_CapPropioSimplificado, CPS_CapPropioSimplVarAnual, "
   Q1 = Q1 & " CPS_AumentosCapital, CPS_GastosRechazadosNoPagan40, CPS_INRPropios, CPS_INRPropiosPerdidas, CPS_OtrosAjustesAumentos, "
   Q1 = Q1 & " CPS_OtrosAjustesDisminuciones, CPS_UtilidadesPerdida, CPS_IngresoDiferido, CPS_CTDImputableIPE, "
   Q1 = Q1 & " CPS_IncentivoAhorro, CPS_IDPCVoluntario, CPS_CredActFijos, CPS_CredParticipaciones "
   Q1 = Q1 & " FROM EmpresasAno WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)

   Grid.Redraw = False
   
   Row = 0
   Grid.TextMatrix(0, C_FMT) = "              .FMT"
   Grid.RowHeight(0) = 0  'Row con el formateo
  
   'Capital Propio Tributario Año Anterior (sólo para TipoInforme Variación Anual)
   If lTipoInforme = CPS_TIPOINFO_VARANUAL Then
      Row = Row + 1
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_TITULO) = "CPT al inicio del ejercicio"
      
      Q1 = "SELECT CPS_CapPropioTrib, CPS_CapPropioSimplificado "
      Q1 = Q1 & " FROM EmpresasAno WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1
      Set Rs2 = OpenRs(DbMain, Q1)
      If Rs2.EOF Then  'No existe año anterior en el sistema, usamos CPS_CapPropioTribAnoAnt de este año
         Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_CapPropioTribAnoAnt")), NUMFMT), 0)
      Else    'si existe año anterior
         If gEmpresa.Ano >= 2021 Then
            Grid.TextMatrix(Row, C_MONTO) = Format(vFld(Rs2("CPS_CapPropioSimplificado")), NUMFMT)
         Else
            Grid.TextMatrix(Row, C_MONTO) = Format(vFld(Rs2("CPS_CapPropioTrib")), NUMFMT)
         End If
      End If
      Call CloseRs(Rs2)
      
      Grid.TextMatrix(Row, C_SIGNO) = "+/-"
      Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
      lRowCapPropioTribAnoAnt = Row
   End If
 
   'Capital Aportado
   Row = Row + 1
   If lTipoInforme = CPS_TIPOINFO_VARANUAL Then
      Row = Row + 1
   End If
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_TITULO) = "Capital Aportado"
   Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_CapitalAportado")), NUMFMT), 0)
   Grid.TextMatrix(Row, C_SIGNO) = "+"
   Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
   lRowCapAportado = Row
      
   'Aumentos Posteriores de Capital
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_TITULO) = "Aumentos posteriores de capital"
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_AumentosCapital")), NUMFMT), 0)
   Else
      Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_AUMENTOSCAP), NUMFMT)
   End If
   Grid.TextMatrix(Row, C_SIGNO) = "+"
   Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
   lRowAumentosCapital = Row
   
   'Reposición Pérdida de Arrastre
   If lTipoInforme = CPS_TIPOINFO_VARANUAL Then
      Row = Row + 2
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_TITULO) = "Reposición pérdida de arrastre"
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_RepPerdidaArrastre")), NUMFMT), 0)
      Grid.TextMatrix(Row, C_SIGNO) = "+"
      Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
      lRowRepPerdidaArrastre = Row
   End If
      
   'Base Imponible del Ejercicio                                        'Base de Impuesto Primera Categoría
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_TITULO) = "Base imponible del ejercicio"     '"Base de Impuesto Primera Categoría"
   If gEmpresa.ProPymeGeneral Then
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_BaseImpPrimCat_14DN3")), NUMFMT), 0)
   Else
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_BaseImpPrimCat_14DN8")), NUMFMT), 0)
   End If
   Grid.TextMatrix(Row, C_SIGNO) = "+/-"
   Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
   lRowBaseImp = Row
  
   'Rentas exentas e ingresos no renta propios - INR Propios
   If gEmpresa.ProPymeGeneral Then
      Row = Row + 2
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_TITULO) = "Rentas exentas e ingresos no renta propios"
      If lTipoInforme = CPS_TIPOINFO_GENERAL Then
         Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_INRPropios")), NUMFMT), 0)
      Else
         Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_INRPROPIOS), NUMFMT)
      End If
      
      Grid.TextMatrix(Row, C_SIGNO) = "+"
      Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
      lRowINRPropios = Row
      
   ElseIf gEmpresa.ProPymeTransp And vFmt(Grid.TextMatrix(Row, C_MONTO)) <> 0 Then
      Grid.TextMatrix(Row, C_MONTO) = 0
      
      'eliminamos todos los registros de detalle de ingreso automático o manual (información de detalle del año en la contabilidad)  y los volvemos a agregar
      Call DeleteSQL(DbMain, "DetCapPropioSimpl", " WHERE TipoDetCPS = " & CPS_INRPROPIOS & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      
      'y ahora el acumunlado anual
      Call DeleteSQL(DbMain, "CapPropioSimplAnual", " WHERE TipoDetCPS = " & CPS_INRPROPIOS & " AND IdEmpresa = " & gEmpresa.id)
            
   End If
  
   'Pérdida por rentas exentas e ingresos no renta del ejercicio    - Pérdidas INR Propios
   If gEmpresa.ProPymeGeneral Then
      Row = Row + 2
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_TITULO) = "Pérdida por rentas exentas e ingresos no renta del ejercicio"
      If lTipoInforme = CPS_TIPOINFO_GENERAL Then
         Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_INRPropiosPerdidas")), NUMFMT), 0)
      Else
         Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_INRPROPIOSPERDIDAS), NUMFMT)
      End If
      
      Grid.TextMatrix(Row, C_SIGNO) = "-"
      Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
      lRowINRPropiosPerdidas = Row
   
   ElseIf gEmpresa.ProPymeTransp And Grid.TextMatrix(Row, C_MONTO) <> 0 Then
         Grid.TextMatrix(Row, C_MONTO) = 0
         
         'eliminamos todos los registros de detalle de ingreso automático o manual (información de detalle del año en la contabilidad)  y los volvemos a agregar
         Call DeleteSQL(DbMain, "DetCapPropioSimpl", " WHERE TipoDetCPS = " & CPS_INRPROPIOSPERDIDAS & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
         
         'y ahora el acumunlado anual
         Call DeleteSQL(DbMain, "CapPropioSimplAnual", " WHERE TipoDetCPS = " & CPS_INRPROPIOSPERDIDAS & " AND IdEmpresa = " & gEmpresa.id)
      
   End If
       
   'Rentas percibidas por paricipaciones
   If gEmpresa.ProPymeGeneral Then
      Row = Row + 2
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_TITULO) = "Rentas percibidas por participaciones"
      
      If lTipoInforme = CPS_TIPOINFO_GENERAL Then
         Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_Participaciones")), NUMFMT), 0)
      Else
         Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_PARTICIPACIONES), NUMFMT)
      End If
      
      Grid.TextMatrix(Row, C_SIGNO) = "+"
      Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
      lRowParticipaciones = Row
      
   ElseIf gEmpresa.ProPymeTransp And vFld(Rs("CPS_Participaciones")) <> 0 Then
      Grid.TextMatrix(Row, C_MONTO) = 0
      
      'eliminamos todos los registros de detalle de ingreso automático o manual (información de detalle del año en la contabilidad)  y los volvemos a agregar
      Call DeleteSQL(DbMain, "DetCapPropioSimpl", " WHERE TipoDetCPS = " & CPS_PARTICIPACIONES & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      
      'y ahora el acumunlado anual
      Call DeleteSQL(DbMain, "CapPropioSimplAnual", " WHERE TipoDetCPS = " & CPS_PARTICIPACIONES & " AND IdEmpresa = " & gEmpresa.id)

   End If
   
   'Utilidades percibidas imputadas a la pérdida del ejercicio
   If gEmpresa.ProPymeGeneral Then
      Row = Row + 2
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_TITULO) = "Utilidades percibidas imputadas a la pérdida del ejercicio"
      
      If lTipoInforme = CPS_TIPOINFO_GENERAL Then
         Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_UtilidadesPerdida")), NUMFMT), 0)
      Else
         Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_UTILIDADESPERDIDA), NUMFMT)
      End If
         
      Grid.TextMatrix(Row, C_SIGNO) = "-"
      Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
      lRowUtilidadesPerdida = Row
   
   ElseIf gEmpresa.ProPymeTransp And vFld(Rs("CPS_UtilidadesPerdida")) <> 0 Then
      Grid.TextMatrix(Row, C_MONTO) = 0
      
      'eliminamos todos los registros de detalle de ingreso automático o manual (información de detalle del año en la contabilidad)  y los volvemos a agregar
      Call DeleteSQL(DbMain, "DetCapPropioSimpl", " WHERE TipoDetCPS = " & CPS_UTILIDADESPERDIDA & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      
      'y ahora el acumunlado anual
      Call DeleteSQL(DbMain, "CapPropioSimplAnual", " WHERE TipoDetCPS = " & CPS_UTILIDADESPERDIDA & " AND IdEmpresa = " & gEmpresa.id)
   
   End If
   
   'Disminuciones formales de capital
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_TITULO) = "Disminuciones formales de capital"
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_Disminuciones")), NUMFMT), 0)
   Else
      Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_DISMINUCIONES), NUMFMT)
   End If
   Grid.TextMatrix(Row, C_SIGNO) = "-"
   Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
   lRowDisminucionesCapital = Row
  
   'Gastos pagados no gravados con el art 21 LIR    - Antes era: Gastos Rechazados  Inc. 2 Art. 21
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_TITULO) = "Gastos pagados no gravados con el art 21 LIR"
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_GastosRechazados")), NUMFMT), 0)
   Else
      Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_GASTOSRECHAZADOS), NUMFMT)
   End If
   Grid.TextMatrix(Row, C_SIGNO) = "-"
   Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
   lRowGastosRechazados = Row
      
   'Retiros o dividendos efectuados a los propietarios
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_TITULO) = "Retiros o dividendos efectuados a los propietarios"
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_RetirosDividendos")), NUMFMT), 0)
   Else
      Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_RETDIV), NUMFMT)
   End If
   Grid.TextMatrix(Row, C_SIGNO) = "-"
   Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
   lRowRetDiv = Row
   

   '-  Ingreso diferido incrementado imputado en el año
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_TITULO) = "Ingreso diferido incrementado imputado en el año"
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_IngresoDiferido")), NUMFMT), 0)
   Else
      Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_INGRESODIFERIDO), NUMFMT)
   End If
   Grid.TextMatrix(Row, C_SIGNO) = "-"
   Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
   lRowIngresoDiferido = Row
      

   '-  CTD imputable contra Impuestos Finales (IPE)
   If gEmpresa.ProPymeTransp And gEmpresa.Ano > 2021 Then
   Else
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_TITULO) = "CTD imputable contra Impuestos Finales (IPE)"
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_CTDImputableIPE")), NUMFMT), 0)
   Else
      Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_CTDIMPUTABLEIPE), NUMFMT)
   End If
   Grid.TextMatrix(Row, C_SIGNO) = "-"
   Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
   lRowCTDImputableIPE = Row
   End If

   '+  Incentivo al ahorro según art. 14 Letra E) de la LIR
   If gEmpresa.ProPymeGeneral Then
      Row = Row + 2
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_TITULO) = "Incentivo al ahorro según art. 14 Letra E) de la LIR"
      If lTipoInforme = CPS_TIPOINFO_GENERAL Then
         Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_IncentivoAhorro")), NUMFMT), 0)
      Else
         Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_INCENTIVOAHORRO), NUMFMT)
      End If
      Grid.TextMatrix(Row, C_SIGNO) = "+"
      Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
      lRowIncentivoAhorro = Row
      
   ElseIf gEmpresa.ProPymeTransp And vFld(Rs("CPS_IncentivoAhorro")) <> 0 Then
      Grid.TextMatrix(Row, C_MONTO) = 0
      
      'eliminamos todos los registros de detalle de ingreso automático o manual (información de detalle del año en la contabilidad)  y los volvemos a agregar
      Call DeleteSQL(DbMain, "DetCapPropioSimpl", " WHERE TipoDetCPS = " & CPS_INCENTIVOAHORRO & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      
      'y ahora el acumunlado anual
      Call DeleteSQL(DbMain, "CapPropioSimplAnual", " WHERE TipoDetCPS = " & CPS_INCENTIVOAHORRO & " AND IdEmpresa = " & gEmpresa.id)
   
   End If

   '+  Base IDPC Voluntario, según art. 14 Letra A n°6 LIR
   If gEmpresa.ProPymeGeneral Then
      Row = Row + 2
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_TITULO) = "Base IDPC Voluntario, según art. 14 Letra A n°6 LIR"
      If lTipoInforme = CPS_TIPOINFO_GENERAL Then
         Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_IDPCVoluntario")), NUMFMT), 0)
      Else
         Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_IDPCVOLUNTARIO), NUMFMT)
      End If
      Grid.TextMatrix(Row, C_SIGNO) = "+"
      Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
      lRowIDPCVoluntario = Row
      
   ElseIf gEmpresa.ProPymeTransp And vFld(Rs("CPS_IDPCVoluntario")) <> 0 Then
      Grid.TextMatrix(Row, C_MONTO) = 0
      
      'eliminamos todos los registros de detalle de ingreso automático o manual (información de detalle del año en la contabilidad)  y los volvemos a agregar
      Call DeleteSQL(DbMain, "DetCapPropioSimpl", " WHERE TipoDetCPS = " & CPS_IDPCVOLUNTARIO & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      
      'y ahora el acumunlado anual
      Call DeleteSQL(DbMain, "CapPropioSimplAnual", " WHERE TipoDetCPS = " & CPS_IDPCVOLUNTARIO & " AND IdEmpresa = " & gEmpresa.id)
   
   End If
 
   '- Crédito por activos fijos adquiridos (art. 33 bis LIR)
   If gEmpresa.ProPymeTransp Then
      Row = Row + 2
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_TITULO) = "Crédito por activos fijos adquiridos (art. 33 bis LIR)"
      If lTipoInforme = CPS_TIPOINFO_GENERAL Then
         Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_CredActFijos")), NUMFMT), 0)
      Else
         Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_CREDACTFIJOS), NUMFMT)
      End If
      Grid.TextMatrix(Row, C_SIGNO) = "-"
      Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
      lRowCredActFijos = Row
      
   ElseIf gEmpresa.ProPymeGeneral And vFld(Rs("CPS_CredActFijos")) <> 0 Then
      Grid.TextMatrix(Row, C_MONTO) = 0
      
      'eliminamos todos los registros de detalle de ingreso automático o manual (información de detalle del año en la contabilidad)  y los volvemos a agregar
      Call DeleteSQL(DbMain, "DetCapPropioSimpl", " WHERE TipoDetCPS = " & CPS_CREDACTFIJOS & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      
      'y ahora el acumunlado anual
      Call DeleteSQL(DbMain, "CapPropioSimplAnual", " WHERE TipoDetCPS = " & CPS_CREDACTFIJOS & " AND IdEmpresa = " & gEmpresa.id)
   
  End If
 
   '-  Crédito por participaciones recibidas
   If gEmpresa.ProPymeTransp Then
      Row = Row + 2
      Call FGrSetRowStyle(Grid, Row, "B")
      Grid.TextMatrix(Row, C_FMT) = "B"
      Grid.TextMatrix(Row, C_TITULO) = "Crédito por participaciones recibidas"
      If lTipoInforme = CPS_TIPOINFO_GENERAL Then
         Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_CredParticipaciones")), NUMFMT), 0)
      Else
         Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_CREDPARTICIPACIONES), NUMFMT)
      End If
      Grid.TextMatrix(Row, C_SIGNO) = "-"
      Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
      lRowCredParticipaciones = Row
      
   ElseIf gEmpresa.ProPymeGeneral And vFld(Rs("CPS_CredParticipaciones")) <> 0 Then
      Grid.TextMatrix(Row, C_MONTO) = 0
      
      'eliminamos todos los registros de detalle de ingreso automático o manual (información de detalle del año en la contabilidad)  y los volvemos a agregar
      Call DeleteSQL(DbMain, "DetCapPropioSimpl", " WHERE TipoDetCPS = " & CPS_CREDPARTICIPACIONES & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
      
      'y ahora el acumunlado anual
      Call DeleteSQL(DbMain, "CapPropioSimplAnual", " WHERE TipoDetCPS = " & CPS_CREDPARTICIPACIONES & " AND IdEmpresa = " & gEmpresa.id)
   
   End If
 
 
   'Otros Ajustes (Aumentos)
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_TITULO) = "Más Otros Ajustes"
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_OtrosAjustesAumentos")), NUMFMT), 0)
   Else
      Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_OTROSAJUSTAUMENTOS), NUMFMT)
   End If
   Grid.TextMatrix(Row, C_SIGNO) = "+"
   Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
   lRowOtrosAjustesAumentos = Row

   'Otros Ajustes (Disminuciones)
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_TITULO) = "Menos Otros Ajustes"
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_OtrosAjustesDisminuciones")), NUMFMT), 0)
   Else
      Grid.TextMatrix(Row, C_MONTO) = Format(GetCPSAnual(CPS_OTROSAJUSTDISMIN), NUMFMT)
   End If
   Grid.TextMatrix(Row, C_SIGNO) = "-"
   Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
   lRowOtrosAjustesDisminuciones = Row
   
   
   '------ TOTAL ------
   
   
   'Monto Capital Propio Tributario Simplificado
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_TITULO) = UCase("Monto Capital Propio Tributario Simplificado")
   Call FGrSetRowStyle(Grid, Row, "FC", vbBlue, C_TITULO, C_TITULO)
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_CapPropioSimplificado")), NUMFMT), 0)
   Else
      Grid.TextMatrix(Row, C_MONTO) = IIf(Not Rs.EOF, Format(vFld(Rs("CPS_CapPropioSimplVarAnual")), NUMFMT), 0)
   End If
   Call FGrSetRowStyle(Grid, Row, "BC", vbInactiveTitleBar, C_MONTO, C_MONTO)
   lRowCapPropioSimplificado = Row
   
   Call CloseRs(Rs)
   
   Call CalcTot     'por si acaso
      
   Grid.rows = Row + 2
   For i = Grid.FixedRows To Row
      Grid.TextMatrix(i, C_OBLIGATORIA) = "."
   Next i
   
   Grid.Redraw = True
End Function

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
   Call ResetPrtBas(gPrtReportes)
   
End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientation As Integer
   
   OldOrientation = Printer.Orientation
   
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
   
   Call ResetPrtBas(gPrtReportes)
      
End Sub

Private Sub Bt_CopyExcel_Click()
   Call LP_FGr2Clip(Grid, Me.Caption)

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

Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(1) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Titulos(0) = "Capital Propio Tributario Simplificado General"
      Titulos(1) = "Artículo 14 D Ley de Renta"
   Else
      Titulos(0) = "Capital Propio Tributario Simplificado por Variación Anual"
      Titulos(1) = "Artículo 14 D Ley de Renta"
   End If
   
   gPrtReportes.Titulos = Titulos
   Encabezados(0) = "Al 31 de Diciembre " & gEmpresa.Ano
   gPrtReportes.Encabezados = Encabezados
         
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
   'Total(C_DESC) = "Capital Pripio Tributario"
   'Total(C_TOTAL) = ""
                  
   gPrtReportes.ColWi = ColWi
   'gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_OBLIGATORIA
   gPrtReportes.NTotLines = 0
   
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   
   Value = Format(vFmt(Value), NUMFMT)
   Grid.TextMatrix(Row, Col) = Value
   Call CalcTot
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
      Exit Sub
   End If
   
   If Col <> C_MONTO Then
      Exit Sub
   End If
   
   If Row = lRowRepPerdidaArrastre Or (Row = lRowCapPropioTribAnoAnt And Not gEmpresa.TieneAnoAnt) Then
      EdType = FEG_Edit
   End If
      
End Sub

Private Sub Grid_DblClick()
   Dim Row As Integer
   Dim Col As Integer
   Dim Rc As Integer
   Dim Valor As Double
   Dim Frm As Form
   
   Row = Grid.Row
   Col = Grid.Col
   
   If Col <> C_MONTO Then
      Exit Sub
   End If
   
   Select Case Row
                     
      Case lRowCapAportado
         
         Set Frm = New FrmCapitalAportado
         Rc = Frm.FEdit(Valor)
         Set Frm = Nothing
         
         If Rc = vbOK Then
            Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
         End If
         
      Case lRowAumentosCapital
         
         Set Frm = New FrmDetCapPropioSimpl
         Rc = Frm.FEdit(CPS_AUMENTOSCAP, 893, lTipoInforme, Valor)
         Set Frm = Nothing
         
         If Rc = vbOK Then
            Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
         End If
         
      Case lRowBaseImp
         
         If gEmpresa.ProPymeGeneral Or gEmpresa.ProPymeTransp Then
        
            Set Frm = New FrmBaseImponible14D
            Rc = Frm.FEdit(lTipoInforme, Valor)
            Set Frm = Nothing
            
            If Rc = vbOK Then
               Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
            End If
         End If
         
      Case lRowParticipaciones
      
         If gEmpresa.ProPymeGeneral Then
   
            Set Frm = New FrmDetCapPropioSimpl
            Rc = Frm.FEdit(CPS_PARTICIPACIONES, 629, lTipoInforme, Valor)
            Set Frm = Nothing
            
            If Rc = vbOK Then
               Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
            End If
            
         End If
            
      Case lRowUtilidadesPerdida
      
         If gEmpresa.ProPymeGeneral Then
   
            Set Frm = New FrmDetCapPropioSimpl
            Rc = Frm.FEdit(CPS_UTILIDADESPERDIDA, 0, lTipoInforme, Valor)
            Set Frm = Nothing
            
            If Rc = vbOK Then
               Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
            End If
            
         End If
            
      Case lRowDisminucionesCapital
            
         Set Frm = New FrmDetCapPropioSimpl
         Rc = Frm.FEdit(CPS_DISMINUCIONES, 894, lTipoInforme, Valor)
         Set Frm = Nothing
         
         If Rc = vbOK Then
            Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
         End If
      
      Case lRowGastosRechazados
      
         Set Frm = New FrmDetCapPropioSimpl
         Rc = Frm.FEdit(CPS_GASTOSRECHAZADOS, 990, lTipoInforme, Valor)
         Set Frm = Nothing
         
         If Rc = vbOK Then
            Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
         End If
     
'      Case lRowGastosRechazadosNoPagan40
'
'         Set Frm = New FrmDetCapPropioSimpl
'         Rc = Frm.FEdit(CPS_GASTOSRECHNOPAGAN40, 1144, lTipoInforme, Valor)
'         Set Frm = Nothing
'
'         If Rc = vbOK Then
'            Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
'         End If
      
     Case lRowRetDiv
     
         Set Frm = New FrmDetCapPropioSimpl
         Rc = Frm.FEdit(CPS_RETDIV, 0, lTipoInforme, Valor)
         Set Frm = Nothing
         
         If Rc = vbOK Then
            Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
         End If
     
      Case lRowINRPropios
      
         If gEmpresa.ProPymeGeneral <> 0 Then
            Set Frm = New FrmDetCapPropioSimpl
            Rc = Frm.FEdit(CPS_INRPROPIOS, 640, lTipoInforme, Valor)
            Set Frm = Nothing
            
            If Rc = vbOK Then
               Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
            End If
            
         End If
         
      Case lRowINRPropiosPerdidas
      
         If gEmpresa.ProPymeGeneral <> 0 Then
            Set Frm = New FrmDetCapPropioSimpl
            Rc = Frm.FEdit(CPS_INRPROPIOSPERDIDAS, 0, lTipoInforme, Valor)
            Set Frm = Nothing
            
            If Rc = vbOK Then
               Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
            End If
            
         End If
         
      Case lRowIngresoDiferido
      
         Set Frm = New FrmDetCapPropioSimpl
         Rc = Frm.FEdit(CPS_INGRESODIFERIDO, 0, lTipoInforme, Valor)
         Set Frm = Nothing
         
         If Rc = vbOK Then
            Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
         End If
            
      Case lRowCTDImputableIPE
      
         Set Frm = New FrmDetCapPropioSimpl
         Rc = Frm.FEdit(CPS_CTDIMPUTABLEIPE, 0, lTipoInforme, Valor)
         Set Frm = Nothing
         
         If Rc = vbOK Then
            Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
         End If
            
         
      Case lRowIncentivoAhorro
      
         If gEmpresa.ProPymeGeneral <> 0 Then
            Set Frm = New FrmDetCapPropioSimpl
            Rc = Frm.FEdit(CPS_INCENTIVOAHORRO, 0, lTipoInforme, Valor)
            Set Frm = Nothing
            
            If Rc = vbOK Then
               Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
            End If
            
         End If
         
      Case lRowIDPCVoluntario
      
         If gEmpresa.ProPymeGeneral <> 0 Then
            Set Frm = New FrmDetCapPropioSimpl
            Rc = Frm.FEdit(CPS_IDPCVOLUNTARIO, 0, lTipoInforme, Valor)
            Set Frm = Nothing
            
            If Rc = vbOK Then
               Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
            End If
            
         End If
         
      Case lRowCredActFijos
      
         If gEmpresa.ProPymeTransp <> 0 Then
            Set Frm = New FrmDetCapPropioSimpl
            Rc = Frm.FEdit(CPS_CREDACTFIJOS, 0, lTipoInforme, Valor)
            Set Frm = Nothing
            
            If Rc = vbOK Then
               Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
            End If
            
         End If
         
      Case lRowCredParticipaciones
      
         If gEmpresa.ProPymeTransp <> 0 Then
            Set Frm = New FrmDetCapPropioSimpl
            Rc = Frm.FEdit(CPS_CREDPARTICIPACIONES, 0, lTipoInforme, Valor)
            Set Frm = Nothing
            
            If Rc = vbOK Then
               Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
            End If
            
         End If
                  
      Case lRowOtrosAjustesAumentos, lRowOtrosAjustesDisminuciones
      
         Set Frm = New FrmDetCapPropioSimplMini
         
         If Row = lRowOtrosAjustesAumentos Then
            Rc = Frm.FEdit(CPS_OTROSAJUSTAUMENTOS, lTipoInforme, Valor)
         Else
            Rc = Frm.FEdit(CPS_OTROSAJUSTDISMIN, lTipoInforme, Valor)
         End If
         
         Set Frm = Nothing
         
         If Rc = vbOK Then
            Grid.TextMatrix(Row, C_MONTO) = Format(Valor, NUMFMT)
         End If
                     
   End Select
   
   Call CalcTot
End Sub
Private Sub CalcTot()
   Dim Total As Double
  
   Total = vFmt(Grid.TextMatrix(lRowCapAportado, C_MONTO))
   Total = Total + vFmt(Grid.TextMatrix(lRowAumentosCapital, C_MONTO))
   Total = Total + vFmt(Grid.TextMatrix(lRowBaseImp, C_MONTO))
   Total = Total + vFmt(Grid.TextMatrix(lRowParticipaciones, C_MONTO))
   Total = Total - vFmt(Grid.TextMatrix(lRowUtilidadesPerdida, C_MONTO))
   Total = Total - vFmt(Grid.TextMatrix(lRowDisminucionesCapital, C_MONTO))
   Total = Total - vFmt(Grid.TextMatrix(lRowGastosRechazados, C_MONTO))
   Total = Total - vFmt(Grid.TextMatrix(lRowRetDiv, C_MONTO))
   Total = Total + vFmt(Grid.TextMatrix(lRowINRPropios, C_MONTO))
   Total = Total - vFmt(Grid.TextMatrix(lRowINRPropiosPerdidas, C_MONTO))
   Total = Total + vFmt(Grid.TextMatrix(lRowOtrosAjustesAumentos, C_MONTO))
   Total = Total - vFmt(Grid.TextMatrix(lRowOtrosAjustesDisminuciones, C_MONTO))
   Total = Total - vFmt(Grid.TextMatrix(lRowIngresoDiferido, C_MONTO))
   Total = Total - vFmt(Grid.TextMatrix(lRowCTDImputableIPE, C_MONTO))
   Total = Total + vFmt(Grid.TextMatrix(lRowIncentivoAhorro, C_MONTO))
   Total = Total + vFmt(Grid.TextMatrix(lRowIDPCVoluntario, C_MONTO))
   Total = Total - vFmt(Grid.TextMatrix(lRowCredActFijos, C_MONTO))
   Total = Total - vFmt(Grid.TextMatrix(lRowCredParticipaciones, C_MONTO))
   
   If lTipoInforme = CPS_TIPOINFO_VARANUAL Then
      Total = Total + vFmt(Grid.TextMatrix(lRowCapPropioTribAnoAnt, C_MONTO))
      Total = Total + vFmt(Grid.TextMatrix(lRowRepPerdidaArrastre, C_MONTO))
   End If
   
   
   If Total < 0 Then
      Grid.TextMatrix(lRowCapPropioSimplificado, C_MONTO) = Format(0, NUMFMT)
   Else
      Grid.TextMatrix(lRowCapPropioSimplificado, C_MONTO) = Format(Total, NUMFMT)
   End If
   
   Tx_TotCapPropio = Grid.TextMatrix(lRowCapPropioSimplificado, C_MONTO)
   
   lTotalCapPropioSimplificado = Total
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)

   If Grid.Row = lRowRepPerdidaArrastre Then
      Call KeyNumPos(KeyAscii)
   ElseIf Grid.Row = lRowCapPropioTribAnoAnt Then
      Call KeyNum(KeyAscii)
   End If
   
End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Dim Row As Integer
   Dim Col As Integer
   
   Row = Grid.Row
   Col = Grid.Col
   
   If Col = C_MONTO And (Row = lRowCapAportado Or Row = lRowBaseImp Or Row = lRowParticipaciones Or Row = lRowDisminucionesCapital Or Row = lRowGastosRechazados Or Row = lRowRetDiv) Then
      Grid.ToolTipText = "Presione doble-click para ingresar y actualizar el detalle"
   Else
      Grid.ToolTipText = ""
   End If
   
   
End Sub

