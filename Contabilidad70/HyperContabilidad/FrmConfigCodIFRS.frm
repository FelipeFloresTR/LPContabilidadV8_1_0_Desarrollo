VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmConfigCodIFRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Códigos IFRS"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   Icon            =   "FrmConfigCodIFRS.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   11685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_DelCodIFRS 
      Caption         =   "Eliminar Código IFRS"
      Height          =   315
      Left            =   9840
      TabIndex        =   10
      ToolTipText     =   "Eliminar código IFRS de la celda seleccionada"
      Top             =   8580
      Width           =   1755
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   7695
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   13573
      Cols            =   8
      Rows            =   20
      FixedCols       =   2
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
      Height          =   615
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   11595
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   9120
         TabIndex        =   5
         Top             =   180
         Width           =   1155
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
         Left            =   960
         Picture         =   "FrmConfigCodIFRS.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_View 
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
         Picture         =   "FrmConfigCodIFRS.frx":0451
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.ComboBox Cb_Informe 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   4695
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
         Picture         =   "FrmConfigCodIFRS.frx":08F8
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   10320
         TabIndex        =   6
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Informe IFRS"
         Height          =   195
         Index           =   0
         Left            =   2460
         TabIndex        =   8
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Label Lb_Nota2 
      AutoSize        =   -1  'True
      Caption         =   "Las cuentas que no tienen código IFRS asociado se muestran en azul"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   780
      TabIndex        =   11
      Top             =   8880
      Width           =   4965
   End
   Begin VB.Label Lb_Nota 
      Caption         =   "Notas:    Para agregar o modificar el código IFRS asociado a una cuenta, haga doble-click en la celda correspondiente"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   8640
      Width           =   8355
   End
End
Attribute VB_Name = "FrmConfigCodIFRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CUENTA = 0
Const C_NOMBCORTO = 1
Const C_DESC = 2
Const C_IDCUENTA = 3
Const C_NIVEL = 4
Const C_CODIFRS = 5
Const C_DESCIFRS = 6
Const C_FMT = 7
Const C_UPDATE = 8


Const NCOLS = C_UPDATE

Dim lEditCodIFRS As Boolean
Dim lCuentasNoConfiguradas As Boolean

Public Sub FConfig()
   lCuentasNoConfiguradas = False
   
   Me.Show vbModal

End Sub

Public Sub FCuentasNoConfiguradas()
   lCuentasNoConfiguradas = True
   
   Me.Show vbModal

End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
  Grid.AllowUserResizing = flexResizeColumns
  
  Grid.Cols = NCOLS + 1
  Call FGrSetup(Grid)
   
   Grid.ColWidth(C_CUENTA) = 1200
   Grid.ColWidth(C_NOMBCORTO) = 0
   Grid.ColWidth(C_DESC) = 4500
   Grid.ColWidth(C_CODIFRS) = 1000
   Grid.ColWidth(C_DESCIFRS) = 4500
   Grid.ColWidth(C_NIVEL) = 0
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_UPDATE) = 0
   Grid.ColWidth(C_FMT) = 0
         
   Grid.TextMatrix(0, C_CUENTA) = "Cuenta"
   Grid.TextMatrix(0, C_NOMBCORTO) = ""    '"Nombre"
   Grid.TextMatrix(0, C_DESC) = "Descripción"
   Grid.TextMatrix(0, C_CODIFRS) = "Cód. IFRS"
   Grid.TextMatrix(0, C_DESCIFRS) = "Descripción Registro IFRS"
   
End Sub


Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, "")
End Sub


Private Sub bt_OK_Click()
   Call SaveAll
   Unload Me
End Sub

Private Sub Bt_Print_Click()
   Dim PrtOrient As Integer
      
   Call SetUpPrtGrid
   
'   PrtOrient = Printer.Orientation
'   Printer.Orientation = cdlLandscape
   
   Me.MousePointer = vbHourglass
   
   Call gPrtReportes.PrtFlexGrid(Printer)
   
   Me.MousePointer = vbDefault
   
'   Printer.Orientation = PrtOrient
   
   Call ResetPrtBas(gPrtReportes)
   
End Sub

Private Sub bt_View_Click()
   Dim Frm As FrmPrintPreview
   
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   Call ResetPrtBas(gPrtReportes)
   
End Sub

Private Sub Cb_Informe_Click()
   Call LoadAll
End Sub

Private Sub Bt_DelCodIFRS_Click()
   Dim Row As Integer
   
   Row = Grid.Row

   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If lEditCodIFRS = False Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_IDCUENTA) = "" Or Val(Grid.TextMatrix(Row, C_NIVEL)) <> gLastNivel Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Grid.Col = C_CODIFRS Or Grid.Col = C_DESCIFRS Then
      If MsgBox1("¿Está seguro que desea eliminar el código IFRS del registro seleccionado?" & vbCrLf & vbCrLf & "Recuerde que cualquier eliminación de cuenta podrá afectar los reportes IFRS.", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Exit Sub
      Else
         Grid.TextMatrix(Row, C_CODIFRS) = ""
         Grid.TextMatrix(Row, C_DESCIFRS) = ""
         Call FGrSetRowStyle(Grid, Row, "FC", vbBlue)

         Call FGrModRow(Grid, Row, FGR_U, C_IDCUENTA, C_UPDATE)
      End If
   End If
         
   
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   Call SetUpGrid
     
   Call CbAddItem(Cb_Informe, gInformeIFRS(IFRS_ESTFIN), IFRS_ESTFIN)
   Call CbAddItem(Cb_Informe, gInformeIFRS(IFRS_ESTRES), IFRS_ESTRES)
   Cb_Informe.ListIndex = 0
   
   Call SetupPriv
   
   MsgBox1 "Recuerde verificar los códigos IFRS para aquellas cuentas que no pertenecen a uno de los planes predefinidos por el sistema," & vbNewLine & "o que han sido modificadas a partir de la definición original en alguno de estos planes.", vbInformation + vbOKOnly

End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(0) As String
   Dim Encabezados(3) As String
   Dim FontTit(0) As FontDef_t
   Dim FontNom(0) As FontDef_t
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Me.Caption
   gPrtReportes.Titulos = Titulos
   
   FontTit(0).FontBold = True
   Call gPrtReportes.FntTitulos(FontTit())
      
   gPrtReportes.GrFontName = "Arial"
   gPrtReportes.GrFontSize = 8
   gPrtReportes.Encabezados = Encabezados
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_IDCUENTA
   gPrtReportes.FmtCol = C_FMT
   gPrtReportes.NTotLines = 0
      
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer, i As Integer
   Dim CurNiv As Integer
   Dim CodFather As String
   Dim RowEstadoRes As Integer
   
   RowEstadoRes = 0
   
   Grid.FlxGrid.Redraw = False
      
   Q1 = "SELECT DISTINCT Cuentas.idCuenta, Cuentas.Codigo, Cuentas.Nombre, Cuentas.Nivel "
   Q1 = Q1 & ", Cuentas.Descripcion As DescCta "
   Q1 = Q1 & ", CodIFRS, IFRS_PlanIFRS.Descripcion As DescIFRS "
   
   Q1 = Q1 & " FROM Cuentas "
   Q1 = Q1 & " LEFT JOIN IFRS_PlanIFRS ON Cuentas.CodIFRS = IFRS_PlanIFRS.Codigo"
   
   Q1 = Q1 & " WHERE Cuentas.Nivel <= " & gLastNivel
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
'   If CbItemData(Cb_Informe) = IFRS_ESTFIN Then
'      Q1 = Q1 & " AND (IFRS_PlanIFRS.Clasificacion IN (" & CLASCTA_ACTIVO & "," & CLASCTA_PASIVO & ") OR IFRS_PlanIFRS.Clasificacion IS NULL)"
'   Else
'      Q1 = Q1 & " AND (IFRS_PlanIFRS.Clasificacion = " & CLASCTA_RESULTADO & " OR IFRS_PlanIFRS.Clasificacion IS NULL)"
'   End If
   If CbItemData(Cb_Informe) = IFRS_ESTFIN Then
      Q1 = Q1 & " AND (Cuentas.Clasificacion IN (" & CLASCTA_ACTIVO & "," & CLASCTA_PASIVO & ") )"
   Else
      Q1 = Q1 & " AND (Cuentas.Clasificacion = " & CLASCTA_RESULTADO & ")"
   End If
   
   Q1 = Q1 & " ORDER BY Cuentas.Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   
   Row = Grid.FixedRows
   Grid.rows = Grid.FixedRows
   Do While Rs.EOF = False
      CurNiv = vFld(Rs("Nivel"))
      
      If CodFather <> Left(vFld(Rs("Codigo")), gNiveles.Largo(CurNiv)) Then
         If Row > 1 Then
            'Salto una línea
            Grid.rows = Row + 1
            Grid.TextMatrix(Row, C_IDCUENTA) = "*******"
            
            Row = Row + 1
            
         End If
         
         CodFather = Left(vFld(Rs("Codigo")), gNiveles.Largo(vFld(Rs("Nivel"))))
         'NomFather = vFld(Rs("Descripcion"))
         
      End If
      
      Grid.rows = Row + 1
      Grid.TextMatrix(Row, C_CUENTA) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
      Grid.TextMatrix(Row, C_DESC) = String(REP_INDENT * (CurNiv - 1), " ") & FCase(vFld(Rs("DescCta"), True))
      Grid.TextMatrix(Row, C_NOMBCORTO) = vFld(Rs("Nombre"))
      
      If LCase(vFld(Rs("DescCta"))) = "estado de resultados" Then
         RowEstadoRes = Row
      End If
      
      If vFld(Rs("Nivel")) = gLastNivel Then
         If vFld(Rs("CodIFRS")) <> "" Then
            Grid.TextMatrix(Row, C_CODIFRS) = FmtCodIFRS(vFld(Rs("CodIFRS")))
            Grid.TextMatrix(Row, C_DESCIFRS) = vFld(Rs("DescIFRS"))
            Call FGrSetRowStyle(Grid, Row, "FC", vbBlack)
         Else
            Call FGrSetRowStyle(Grid, Row, "FC", vbBlue)
         End If
         
      End If
      
      Grid.TextMatrix(Row, C_IDCUENTA) = vFld(Rs("idCuenta"))
      Grid.TextMatrix(Row, C_NIVEL) = vFld(Rs("Nivel"))
      
      If vFld(Rs("Nivel")) <= 2 Then
         Call FGrFontBold(Grid, Row, -1, True)
         Grid.TextMatrix(Row, C_FMT) = "B"
      End If
         
        
      Row = Row + 1
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   
   Grid.TopRow = Grid.FixedRows
   
   If CbItemData(Cb_Informe) = IFRS_ESTRES And RowEstadoRes >= Grid.FixedRows Then
      Grid.TopRow = RowEstadoRes
   End If
   
   Grid.FlxGrid.Redraw = True
   
End Sub
Private Sub Form_Resize()
    Dim d As Integer

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

'   d = Me.Width - 2 * (Grid.Left + W.xFrame)
'   If d > 1000 Then
'      Grid.Width = d
'   End If
 
   d = Me.Height - Grid.Top - W.YCaption * 2 + 80 - Lb_Nota.Height - 300
   If d > 1000 Then
      Grid.Height = d
   Else
      Me.Height = Grid.Top + 1000 + W.YCaption * 2
   End If
   
   Lb_Nota.Top = Grid.Top + Grid.Height + 100
   Lb_Nota2.Top = Grid.Top + Grid.Height + 330
   Bt_DelCodIFRS.Top = Grid.Top + Grid.Height + 100
   
   Call FGrVRows(Grid)
   
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)

'   If Col = C_CODIFRS Or Col = C_DESCIFRS Then
'      If Val(Value) < 0 Then
'         MsgBox1 "Código inválido.", vbExclamation
'         Action = vbCancel
'      Else
'         Call FGrModRow(Grid, Row, FGR_U, C_IDCUENTA, C_UPDATE)
'         Action = vbOK
'      End If
'   End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   Dim Frm As FrmLstInformeIFRS
   Dim CodIFRS As String
   Dim DescIFRS As String
   
   EdType = FEG_None
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If lEditCodIFRS = False Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_IDCUENTA) = "" Or Val(Grid.TextMatrix(Row, C_NIVEL)) <> gLastNivel Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Col = C_CODIFRS Or Col = C_DESCIFRS Then
      'EdType = FEG_None  'por ahora no dejamos editar códigos IFRS
      
      Set Frm = New FrmLstInformeIFRS
      
      If Frm.FSelect(CbItemData(Cb_Informe), CodIFRS, DescIFRS) = vbOK Then
         Grid.TextMatrix(Row, C_CODIFRS) = FmtCodIFRS(CodIFRS)
         Grid.TextMatrix(Row, C_DESCIFRS) = DescIFRS
         Call FGrSetRowStyle(Grid, Row, "FC", vbBlack)
         Call FGrModRow(Grid, Row, FGR_U, C_IDCUENTA, C_UPDATE)
      End If
      
      Set Frm = Nothing
   End If
   
End Sub

Private Sub SaveAll()
   Dim i As Integer, j As Integer
   Dim Q1 As String
   Dim Qcp As String
   Dim QCod As String
   
   For i = Grid.FixedRows To Grid.rows - 1
      
      If Grid.TextMatrix(i, C_IDCUENTA) = "" Then
         Exit For
      End If
      
      Q1 = ""
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
         
         QCod = " CodIFRS = '" & ParaSQL(VFmtCodigoIFRS(Grid.TextMatrix(i, C_CODIFRS))) & "'"
         
         Q1 = "UPDATE Cuentas SET " & QCod & " WHERE IdCuenta=" & Grid.TextMatrix(i, C_IDCUENTA)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
        
         Call ExecSQL(DbMain, Q1)
      End If
      
   Next i

End Sub

Private Function SetupPriv()
   
   If Not ChkPriv(PRV_ADM_CTAS) Then
      Bt_OK.Enabled = False
      lEditCodIFRS = False
      Bt_DelCodIFRS.Enabled = False
   Else
      lEditCodIFRS = True
   End If
   
End Function


