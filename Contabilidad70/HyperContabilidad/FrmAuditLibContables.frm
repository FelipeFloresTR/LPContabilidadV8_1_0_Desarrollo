VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAuditLibContables 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auditoria de Libros Contables"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Fr_Opciones 
      Height          =   2355
      Left            =   9000
      TabIndex        =   21
      Top             =   1380
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CheckBox Ch_ViewOtrosIngEgr14TER 
         Caption         =   "Ver Otros Ing. Egr.14TER"
         Height          =   195
         Left            =   300
         TabIndex        =   5
         Top             =   360
         Width           =   2220
      End
      Begin VB.CheckBox Ch_ViewCodCuenta 
         Caption         =   "Ver Cód. Cuenta Contable"
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox Ch_ViewAreaNeg 
         Caption         =   "Ver  Áreas de Negocio"
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox Ch_ViewCCosto 
         Caption         =   "Ver Centros de Gestión"
         Height          =   195
         Left            =   300
         TabIndex        =   8
         Top             =   1440
         Width           =   2145
      End
      Begin VB.CheckBox Ch_ViewGlosaComp 
         Caption         =   "Ver Glosa Comprobante"
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   1800
         Width           =   2205
      End
      Begin VB.CommandButton Bt_CerrarOpt 
         Caption         =   "X"
         Height          =   195
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   195
      End
   End
   Begin VB.CommandButton Bt_Opciones 
      Caption         =   "Opciones de Vista"
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      Top             =   1020
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H00FF0000&
      Height          =   795
      Left            =   60
      TabIndex        =   17
      Top             =   780
      Width           =   9675
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   1335
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton Bt_Search 
         Height          =   375
         Left            =   8160
         Picture         =   "FrmAuditLibContables.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Periodo"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   19
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   11775
      Begin VB.CommandButton Bt_DetComp 
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
         Picture         =   "FrmAuditLibContables.frx":0550
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Detalle comprobante seleccionado"
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
         Left            =   1080
         Picture         =   "FrmAuditLibContables.frx":09B5
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10260
         TabIndex        =   15
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
         Left            =   660
         Picture         =   "FrmAuditLibContables.frx":0E6F
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   1500
         Picture         =   "FrmAuditLibContables.frx":1316
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   2040
         Picture         =   "FrmAuditLibContables.frx":175B
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Calendario"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7455
      Left            =   60
      TabIndex        =   3
      Top             =   1680
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13150
      _Version        =   393216
      Rows            =   25
      Cols            =   11
      FixedCols       =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "FrmAuditLibContables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDCOMP = 0
Const C_FECHACOMP = 1
Const C_CORRCOMP = 2
Const C_TIPOCOMP = 3
Const C_OTROSINGEGR = 4
Const C_CODCUENTA = 5
Const C_CUENTA = 6
Const C_GLOSACOMP = 7
Const C_RUTENT = 8
Const C_ENTIDAD = 9
Const C_TIPODOC = 10
Const C_NUMDOC = 11
Const C_FEMISION = 12
Const C_FVENC = 13
Const C_GLOSAMOV = 14
Const C_AREANEG = 15
Const C_CCOSTO = 16
Const C_DEBE = 17
Const C_HABER = 18


Const NCOLS = C_HABER

Dim lOrientacion As Integer


Private Sub Bt_Calendar_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
   
   Call Frm.SelDate(Fecha)
   
   Set Frm = Nothing
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_CerrarOpt_Click()
   Fr_Opciones.visible = False

End Sub

Private Sub Bt_CopyExcel_Click()
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de copiar.", vbExclamation
      Exit Sub
   End If
   
   Call FGr2Clip(Grid, Me.Caption & vbTab & "Periodo: " & vbTab & Cb_Mes & " " & Cb_Ano)
End Sub

Private Sub Bt_DetComp_Click()

   Call ViewDetComp(Grid.Row, Grid.Col)
   
End Sub

Private Sub Bt_Opciones_Click()

   Fr_Opciones.visible = Not Fr_Opciones.visible

End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   
   If Bt_Search.Enabled = True Then
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
   
   Call ResetPrtBas(gPrtReportes)
   
End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientation As Integer
         
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   OldOrientation = Printer.Orientation
      
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
      
   Printer.Orientation = OldOrientation
   
   Call ResetPrtBas(gPrtReportes)
   
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
      
   Encabezados(0) = "Periodo: " & vbTab & Cb_Mes & " " & Cb_Ano
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_IDCOMP
   gPrtReportes.NTotLines = 0
   

End Sub

Private Sub Bt_Search_Click()

   MousePointer = vbHourglass
      
   Call LoadAll
   MousePointer = vbDefault
   
End Sub

Private Sub Cb_Ano_Click()
   Bt_Search.Enabled = True

End Sub

Private Sub Cb_Mes_Click()
   Bt_Search.Enabled = True

End Sub

Private Sub Ch_ViewAreaNeg_Click()

   If Ch_ViewAreaNeg = 0 Then
      Grid.ColWidth(C_AREANEG) = 0
      Grid.TextMatrix(0, C_AREANEG) = ""
      
   Else
      Grid.ColWidth(C_AREANEG) = 2000
      Grid.TextMatrix(0, C_AREANEG) = "Área de Negocio"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerAreaNeg", Abs(Ch_ViewAreaNeg.Value))
   gVarIniFile.VerAreaNeg = Abs(Ch_ViewAreaNeg.Value)


End Sub

Private Sub Ch_ViewCCosto_Click()

   If Ch_ViewCCosto = 0 Then
      Grid.ColWidth(C_CCOSTO) = 0
      Grid.TextMatrix(0, C_CCOSTO) = ""
      
   Else
      Grid.ColWidth(C_CCOSTO) = 2000
      Grid.TextMatrix(0, C_CCOSTO) = "Centro de Gestión"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerCCosto", Abs(Ch_ViewCCosto.Value))
   gVarIniFile.VerCCosto = Abs(Ch_ViewCCosto.Value)


End Sub

Private Sub Ch_ViewCodCuenta_Click()

   If Ch_ViewCodCuenta = 0 Then
      Grid.ColWidth(C_CODCUENTA) = 0
      Grid.TextMatrix(0, C_CODCUENTA) = ""
      
   Else
      Grid.ColWidth(C_CODCUENTA) = FW_CUENTA
      Grid.TextMatrix(0, C_CODCUENTA) = "Cód. Cuenta"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerCodCuenta", Abs(Ch_ViewCodCuenta.Value))
   gVarIniFile.VerCodCuenta = Abs(Ch_ViewCodCuenta.Value)


End Sub

Private Sub Ch_ViewGlosaComp_Click()

   If Ch_ViewGlosaComp = 0 Then
      Grid.ColWidth(C_GLOSACOMP) = 0
      Grid.TextMatrix(0, C_GLOSACOMP) = ""
      
   Else
      Grid.ColWidth(C_GLOSACOMP) = 3000
      Grid.TextMatrix(0, C_GLOSACOMP) = "Glosa Comp."
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerGlosaComp", Abs(Ch_ViewGlosaComp.Value))
   gVarIniFile.VerGlosaComp = Abs(Ch_ViewGlosaComp.Value)



End Sub

Private Sub Ch_ViewOtrosIngEgr14TER_Click()

   If Ch_ViewOtrosIngEgr14TER = 0 Then
      Grid.ColWidth(C_OTROSINGEGR) = 0
      Grid.TextMatrix(0, C_OTROSINGEGR) = ""
      
   Else
      Grid.ColWidth(C_OTROSINGEGR) = 1700
      Grid.TextMatrix(0, C_OTROSINGEGR) = "Otros Ing/Egr. 14TER"
         
   End If
   
   Call SetIniString(gIniFile, "Opciones", "VerOtrosIngEgr14TER", Abs(Ch_ViewOtrosIngEgr14TER.Value))
   gVarIniFile.VerOtrosIngEgr14TER = Abs(Ch_ViewOtrosIngEgr14TER.Value)


End Sub

Private Sub Form_Load()
   Dim MesActual As Integer
   Dim F1 As Long
   Dim F2 As Long
   
   lOrientacion = ORIENT_HOR
         
   MesActual = GetMesActual()
   If MesActual = 0 Then
      MesActual = 1
   End If
   Call CbFillMes(Cb_Mes, MesActual)
      
   Cb_Ano.AddItem gEmpresa.Ano
   Cb_Ano.ListIndex = Cb_Ano.NewIndex
   
   Grid.Cols = NCOLS + 1
   
   Ch_ViewOtrosIngEgr14TER.visible = True
   Ch_ViewOtrosIngEgr14TER = gVarIniFile.VerOtrosIngEgr14TER

   Ch_ViewCodCuenta.visible = True
   Ch_ViewCodCuenta = gVarIniFile.VerCodCuenta

   Ch_ViewAreaNeg.visible = True
   Ch_ViewAreaNeg = gVarIniFile.VerAreaNeg

   Ch_ViewCCosto.visible = True
   Ch_ViewCCosto = gVarIniFile.VerCCosto

   Ch_ViewGlosaComp.visible = True
   Ch_ViewGlosaComp = gVarIniFile.VerGlosaComp

   Call SetUpGrid
   Call LoadAll
   
   DoEvents
   
   MsgBox1 "Este reporte solo considera comprobantes en estado aprobado.", vbInformation + vbOKOnly
 
End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
   Grid.FixedRows = 1
   Grid.FixedCols = C_TIPOCOMP + 1
   
   Call FGrSetup(Grid)
    
   Grid.ColWidth(C_IDCOMP) = 0
   Grid.ColWidth(C_FECHACOMP) = FW_FECHA + 50
   Grid.ColWidth(C_CORRCOMP) = 1000
   Grid.ColWidth(C_TIPOCOMP) = 800
   Grid.ColWidth(C_OTROSINGEGR) = 1700
   Grid.ColWidth(C_CODCUENTA) = FW_CUENTA
   Grid.ColWidth(C_CUENTA) = 2500
   Grid.ColWidth(C_GLOSACOMP) = 3000
   Grid.ColWidth(C_RUTENT) = FW_RUT
   Grid.ColWidth(C_ENTIDAD) = 3000
   Grid.ColWidth(C_TIPODOC) = 500
   Grid.ColWidth(C_NUMDOC) = FW_NUM
   Grid.ColWidth(C_FEMISION) = FW_FECHA + 50
   Grid.ColWidth(C_FVENC) = FW_FECHA + 50
   Grid.ColWidth(C_GLOSAMOV) = 2500
   Grid.ColWidth(C_AREANEG) = 2000
   Grid.ColWidth(C_CCOSTO) = 2000
   Grid.ColWidth(C_DEBE) = 1300
   Grid.ColWidth(C_HABER) = 1300
      
   Grid.ColAlignment(C_IDCOMP) = flexAlignRightCenter
   Grid.ColAlignment(C_FECHACOMP) = flexAlignRightCenter
   Grid.ColAlignment(C_CORRCOMP) = flexAlignRightCenter
   Grid.ColAlignment(C_TIPOCOMP) = flexAlignLeftCenter
   Grid.ColAlignment(C_OTROSINGEGR) = flexAlignCenterCenter
   Grid.ColAlignment(C_CODCUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_GLOSACOMP) = flexAlignLeftCenter
   Grid.ColAlignment(C_RUTENT) = flexAlignRightCenter
   Grid.ColAlignment(C_ENTIDAD) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPODOC) = flexAlignCenterCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_FEMISION) = flexAlignRightCenter
   Grid.ColAlignment(C_FVENC) = flexAlignRightCenter
   Grid.ColAlignment(C_GLOSAMOV) = flexAlignLeftCenter
   Grid.ColAlignment(C_AREANEG) = flexAlignLeftCenter
   Grid.ColAlignment(C_CCOSTO) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_FECHACOMP) = "Fecha Comp."
   Grid.TextMatrix(0, C_CORRCOMP) = "N° Comp."
   Grid.TextMatrix(0, C_TIPOCOMP) = "Tipo"
   Grid.TextMatrix(0, C_OTROSINGEGR) = "Otros Ing/Egr. 14TER"
   Grid.TextMatrix(0, C_CODCUENTA) = "Cód. Cuenta"
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.TextMatrix(0, C_GLOSACOMP) = "Glosa Comp."
   Grid.TextMatrix(0, C_RUTENT) = "RUT Entidad"
   Grid.TextMatrix(0, C_ENTIDAD) = "Razón Social"
   Grid.TextMatrix(0, C_TIPODOC) = "TD"
   Grid.TextMatrix(0, C_NUMDOC) = "N° Doc."
   Grid.TextMatrix(0, C_FEMISION) = "Emisión"
   Grid.TextMatrix(0, C_FVENC) = "Vencimiento"
   Grid.TextMatrix(0, C_GLOSAMOV) = "Glosa Específica"
   Grid.TextMatrix(0, C_AREANEG) = "Área de Negocio"
   Grid.TextMatrix(0, C_CCOSTO) = "Centro de Gestión"
   Grid.TextMatrix(0, C_DEBE) = "Debe"
   Grid.TextMatrix(0, C_HABER) = "Haber"
   
   If Ch_ViewOtrosIngEgr14TER = 0 Then
      Grid.ColWidth(C_OTROSINGEGR) = 0
      Grid.TextMatrix(0, C_OTROSINGEGR) = ""
   End If

   If Ch_ViewCodCuenta = 0 Then
      Grid.ColWidth(C_CODCUENTA) = 0
      Grid.TextMatrix(0, C_CODCUENTA) = ""
   End If

   If Ch_ViewAreaNeg = 0 Then
      Grid.ColWidth(C_AREANEG) = 0
      Grid.TextMatrix(0, C_AREANEG) = ""
   End If

   If Ch_ViewCCosto = 0 Then
      Grid.ColWidth(C_CCOSTO) = 0
      Grid.TextMatrix(0, C_CCOSTO) = ""
   End If

   If Ch_ViewGlosaComp = 0 Then
      Grid.ColWidth(C_GLOSACOMP) = 0
      Grid.TextMatrix(0, C_GLOSACOMP) = ""
   End If

   
   Call FGrVRows(Grid)
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Wh As String
   Dim UsrJoin As String
   Dim IdComp As Long
   
   Q1 = "SELECT Comprobante.IdComp, Comprobante.Fecha, Comprobante.Correlativo, Comprobante.Tipo, Comprobante.OtrosIngEg14TER, Cuentas.Codigo "
   Q1 = Q1 & ", Cuentas.Descripcion as Cuenta, Comprobante.Glosa As GlosaComp, Entidades.Rut, Entidades.Nombre, Entidades.NotValidRut, Documento.TipoLib "
   Q1 = Q1 & ", Documento.TipoDoc, Documento.NumDoc "
   Q1 = Q1 & ", Documento.FEmisionOri, Documento.FVenc, MovComprobante.Glosa As GlosaMov, AreaNegocio.Descripcion As AreaNeg, CentroCosto.Descripcion As CCosto "
   Q1 = Q1 & ", MovComprobante.Debe, MovComprobante.Haber "
   Q1 = Q1 & "  FROM (((((Comprobante INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & "  INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.idCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & "  LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante") & " )"
   Q1 = Q1 & "  LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Entidades", "Documento", True, True) & " )"
   Q1 = Q1 & "  LEFT JOIN CentroCosto ON MovComprobante.idCCosto = CentroCosto.IdCCosto "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovComprobante", "CentroCosto", True, True) & " )"
   Q1 = Q1 & "  LEFT JOIN AreaNegocio ON MovComprobante.idAreaNeg = AreaNegocio.IdAreaNegocio "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovComprobante", "AreaNegocio", True, True)
   Q1 = Q1 & "  WHERE " & SqlMonthLng("Fecha") & " = " & CbItemData(Cb_Mes) & " AND " & SqlYearLng("Fecha") & " = " & Cb_Ano
   Q1 = Q1 & "  AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & "  AND Comprobante.Estado = " & EC_APROBADO
   Q1 = Q1 & "  AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
   Q1 = Q1 & "  ORDER BY Comprobante.Fecha, Comprobante.IdComp "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.Redraw = False
   
   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   Do While Rs.EOF = False
   
      Grid.rows = i + 1
      
      If FGrChkMaxSize(Grid) = True Then
         Exit Do
      End If
      
      
      Grid.TextMatrix(i, C_IDCOMP) = vFld(Rs("IdComp"))
      Grid.TextMatrix(i, C_FECHACOMP) = IIf(vFld(Rs("Fecha")) > 0, Format(vFld(Rs("Fecha")), EDATEFMT), "")
      Grid.TextMatrix(i, C_CORRCOMP) = vFld(Rs("Correlativo"))
      Grid.TextMatrix(i, C_TIPOCOMP) = IIf(IdComp <> vFld(Rs("IdComp")), gTipoComp(vFld(Rs("Tipo"))), "")
      Grid.TextMatrix(i, C_OTROSINGEGR) = IIf(IdComp <> vFld(Rs("IdComp")) And vFld(Rs("OtrosIngEg14TER")) <> 0, "Si", "")
      Grid.TextMatrix(i, C_CODCUENTA) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
      Grid.TextMatrix(i, C_CUENTA) = vFld(Rs("Cuenta"))
      Grid.TextMatrix(i, C_GLOSACOMP) = IIf(IdComp <> vFld(Rs("IdComp")), vFld(Rs("GlosaComp")), "")
      Grid.TextMatrix(i, C_RUTENT) = IIf(vFld(Rs("Rut")) <> "" And vFld(Rs("Rut")) <> "0", FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = False), "")
      Grid.TextMatrix(i, C_ENTIDAD) = vFld(Rs("Nombre"))
      Grid.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
      Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
      Grid.TextMatrix(i, C_FEMISION) = IIf(vFld(Rs("FEmisionOri")) > 0, Format(vFld(Rs("FEmisionOri")), EDATEFMT), "")
      Grid.TextMatrix(i, C_FVENC) = IIf(vFld(Rs("FVenc")) > 0, Format(vFld(Rs("FVenc")), EDATEFMT), "")
      Grid.TextMatrix(i, C_GLOSAMOV) = vFld(Rs("GlosaMov"))
      Grid.TextMatrix(i, C_AREANEG) = vFld(Rs("AreaNeg"))
      Grid.TextMatrix(i, C_CCOSTO) = vFld(Rs("CCosto"))
      Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("Debe")), NUMFMT)
      Grid.TextMatrix(i, C_HABER) = Format(vFld(Rs("Haber")), NUMFMT)
      
      IdComp = vFld(Rs("IdComp"))
      
      Rs.MoveNext

      i = i + 1
      
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid, 1)
   
   Grid.TopRow = Grid.FixedRows
   
   Grid.Col = C_CORRCOMP
   Grid.RowSel = Grid.Row
   Grid.ColSel = Grid.Col

   Grid.Redraw = True
   
   Call EnableFrm(False)
End Sub
Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - 600
   
   Grid.Width = Me.Width - Grid.Left - 200
      
   Call FGrVRows(Grid)

End Sub

Private Sub Grid_Click()
   Dim Col As Integer
   Dim Row As Integer
         
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Row >= Grid.FixedRows Then
      Exit Sub
   End If
   
End Sub

Private Sub Grid_DblClick()
   Dim Col As Integer
   Dim Row As Integer
   Dim i As Integer
         
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   Call ViewDetComp(Row, Col)
         
End Sub
Private Sub ViewDetComp(ByVal Row As Integer, ByVal Col As Integer)
   Dim IdComp As Long
   Dim Frm As FrmComprobante

   If Row < Grid.FixedRows Then
      Exit Sub
   End If
      
   IdComp = Val(Grid.TextMatrix(Row, C_IDCOMP))

   If IdComp <> 0 Then
      Set Frm = New FrmComprobante
      Call Frm.FView(IdComp, False)
      Set Frm = Nothing
      
   ElseIf Grid.TextMatrix(Row, C_IDCOMP) = "0" Then
      MsgBox1 "Este comprobante ha sido eliminado.", vbExclamation + vbOKOnly
      
   End If
            
End Sub
Public Sub FView()
   Me.Show vbModal
End Sub

Private Sub EnableFrm(bool As Boolean)

   Bt_Search.Enabled = bool
   
End Sub
