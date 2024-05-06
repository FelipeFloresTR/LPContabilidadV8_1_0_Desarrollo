VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmAFCompsFicha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activo Fijo - Detalle Financiero Componentes"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_ConfigComp 
      Cancel          =   -1  'True
      Caption         =   "Configurar componentes..."
      Height          =   315
      Left            =   6540
      TabIndex        =   7
      Top             =   7260
      Width           =   2235
   End
   Begin VB.Frame Frame3 
      Caption         =   "Activo Fijo"
      Height          =   1275
      Left            =   1620
      TabIndex        =   12
      Top             =   420
      Width           =   5595
      Begin VB.TextBox Tx_Grupo 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   18
         Top             =   300
         Width           =   2235
      End
      Begin VB.TextBox Tx_Cantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4620
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   17
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox Tx_Descrip 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   13
         Top             =   780
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   3780
         TabIndex        =   16
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Descrip:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   780
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle Componentes"
      Height          =   5115
      Left            =   480
      TabIndex        =   10
      Top             =   1980
      Width           =   8295
      Begin VB.CheckBox Ch_NoExisteValRazonable 
         Caption         =   "No Existe Valor Razonable"
         Height          =   195
         Left            =   5760
         TabIndex        =   2
         Top             =   4680
         Width           =   2235
      End
      Begin VB.CheckBox Ch_SinDetComps 
         Caption         =   "Sin detalle de Componentes"
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   900
         Width           =   2415
      End
      Begin VB.ListBox Ls_CompGrupo 
         Height          =   2400
         Left            =   5700
         TabIndex        =   19
         Top             =   900
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Bt_AddComp 
         Height          =   480
         Left            =   6240
         Picture         =   "FrmAFCompsFicha.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Agregar componente a este activo fijo"
         Top             =   420
         Width           =   540
      End
      Begin VB.CommandButton Bt_SaveComp 
         Height          =   480
         Left            =   6840
         Picture         =   "FrmAFCompsFicha.frx":0578
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Grabar cambios componente seleccionada"
         Top             =   420
         Width           =   540
      End
      Begin VB.CommandButton Bt_DelComp 
         Height          =   480
         Left            =   7380
         Picture         =   "FrmAFCompsFicha.frx":091B
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Eliminar componente seleccionada de este activo fijo"
         Top             =   420
         Width           =   540
      End
      Begin FlexEdGrid2.FEd2Grid Grid 
         Height          =   3255
         Left            =   300
         TabIndex        =   1
         Top             =   1260
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5741
         Cols            =   2
         Rows            =   2
         FixedCols       =   1
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
      Begin VB.ComboBox Cb_Componente 
         Height          =   315
         Left            =   1380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   4155
      End
      Begin VB.Label Label1 
         Caption         =   "Componente:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   540
         Width           =   1215
      End
   End
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   7560
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   480
      Picture         =   "FrmAFCompsFicha.frx":0F7D
      ScaleHeight     =   630
      ScaleWidth      =   825
      TabIndex        =   9
      Top             =   480
      Width           =   885
   End
End
Attribute VB_Name = "FrmAFCompsFicha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const C_FIELD = 0
Const C_EDIT = 1
Const C_DESC = 2
Const C_VALOR = 3
Const C_FORMAT = 4

Const NCOLS = C_FORMAT

Const D_PJEDIVCOMP = 1
Const D_VALORCOMPRA = 2
Const D_VALORCRESIDUAL = 3
Const D_PJEAMORTIZACION = 4
Const D_VIDAUTIL = 5
Const D_COSTOSADICIONALES = 6
Const D_TASADESC = 7
Const D_COSTODESMANT = 8
Const D_VALACTCOSTODESMANT = 9
Const D_VALORBIEN = 10
Const D_VALORRAZONABLE_31_12 = 11
Const D_OTRASDIFERENCIAS = 12

Const nRows = D_OTRASDIFERENCIAS + 1

Const L_IDCOMP = 2

Dim lCbComponente As ClsCombo

Dim lIdActFijo As Long
Dim lIdGrupo As Long
Dim lIdFicha As Long
Dim lIdCompFicha As Long
Dim lidComp As Long
Dim lInLoad As Boolean
Dim lModif As Boolean

Dim lFEditIdCompFicha As Long

Dim lSave As Boolean

Dim lPrecioFactura As Double
Dim lDerechosIntern As Double
Dim lTransporte As Double
Dim lObrasAdapt As Double
Dim lAFNoDepreciable As Boolean

Dim lFImported As Long
Dim lFromReport As Boolean


Public Function FEdit(ByVal IdActFijo As Long, Optional ByVal IdCompFicha As Long = 0, Optional ByVal FromReport As Boolean = False) As Integer

   lIdActFijo = IdActFijo
   lFEditIdCompFicha = IdCompFicha
   lFromReport = FromReport
   
   Me.Show vbModal
   
   FEdit = IIf(lSave, vbOK, vbCancel)
   
End Function

Private Sub bt_Cerrar_Click()
      
   If lModif Then
   
      If Ch_SinDetComps = 0 And lCbComponente.ListCount = 0 Then
         Call SaveFicha
      
      ElseIf MsgBox1("¿Desea guardar los cambios realizados a este detalle?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         If Not Valida() Then
            Exit Sub
         End If

         Call SaveCompAF
      End If
   End If
   
   Unload Me
End Sub


Private Sub Bt_ConfigComp_Click()
   Dim Frm As FrmConfigActFijoIFRS
   Dim Q1 As String
   
   Set Frm = New FrmConfigActFijoIFRS
   Frm.Show vbModal
   Set Frm = Nothing
   
   Q1 = "SELECT NombComp, IdComp "
   Q1 = Q1 & " FROM AFComponentes "
   Q1 = Q1 & " WHERE IdGrupo = " & lIdGrupo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " ORDER BY NombComp "
   
   Ls_CompGrupo.Clear
   
   Call FillCombo(Ls_CompGrupo, DbMain, Q1, 0)


End Sub

Private Sub Bt_DelComp_Click()
   Dim id As Long
   Dim Q1 As String
   
   If Ch_SinDetComps <> 0 Or lCbComponente.ListIndex < 0 Then
      Exit Sub
   End If
   
   id = lCbComponente.ItemData
   If id <= 0 Then
      MsgBox1 "Componente inválida.", vbExclamation
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea eliminar esta componente?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
      Exit Sub
   End If
   
   Call DeleteSQL(DbMain, "ActFijoCompsFicha", "WHERE IdCompFicha = " & id)
   
   lModif = False
   
   lCbComponente.RemoveItem (lCbComponente.ListIndex)
   If lCbComponente.ListCount > 0 Then
      lCbComponente.ListIndex = 0
   End If
   

End Sub

Private Sub Bt_SaveComp_Click()
   
   If Not Valida() Then
      Exit Sub
   End If

   Call SaveCompAF
End Sub

Private Sub Bt_AddComp_Click()
   Ls_CompGrupo.visible = Not Ls_CompGrupo.visible
   
End Sub

Private Sub Ch_NoExisteValRazonable_Click()
   
   If Ch_NoExisteValRazonable <> 0 Then
      Grid.TextMatrix(D_VALORRAZONABLE_31_12, C_VALOR) = ""
      Grid.TextMatrix(D_VALORRAZONABLE_31_12, C_EDIT) = "0"
      Call FGrSetRowStyle(Grid, D_VALORRAZONABLE_31_12, "BC", vbButtonFace, C_VALOR, C_VALOR)
   Else
      Grid.TextMatrix(D_VALORRAZONABLE_31_12, C_EDIT) = "1"
      Call FGrSetRowStyle(Grid, D_VALORRAZONABLE_31_12, "BC", vbWindowBackground, C_VALOR, C_VALOR)
   End If
      
   lModif = True
   Call CalcFicha
      
End Sub

Private Sub Ch_SinDetComps_Click()
   Dim Q1 As String
   Static InClick As Boolean
   
   If InClick = True Then
      Exit Sub
   End If
   
   InClick = True
   
   If Ch_SinDetComps <> 0 Then
      
      If lCbComponente.ListCount > 0 Then
         If MsgBox1("Si selecciona esta opción perderá la información de las componentes ya ingresadas." & vbCrLf & vbCrLf & "¿Desea continar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Ch_SinDetComps = 0
            InClick = False
            Exit Sub
         End If
            
         lCbComponente.Clear
         
         Call DeleteSQL(DbMain, "ActFijoCompsFicha", "WHERE IdActFijo = " & lIdActFijo)
         
         
      End If
      
      Bt_AddComp.Enabled = False
      Bt_DelComp.Enabled = False
      
      lidComp = -1
      lIdCompFicha = 0
      
      Call ClearGrid
      
      Grid.TextMatrix(D_PJEDIVCOMP, C_VALOR) = Format(1, Grid.TextMatrix(D_PJEDIVCOMP, C_FORMAT))
      Call CalcFicha
      lModif = True
      
   Else      'esto solo ocurre cuando tenía marcado SinDetComp y lo desmarca
      
      If lFImported = 0 And gEmpresa.FCierre = 0 Then
         Bt_AddComp.Enabled = True
         Bt_DelComp.Enabled = True
      End If
      lidComp = 0
            
      'eliminamos la componente única (sin detalle pero se crea una componente con Id -1)
      Call DeleteSQL(DbMain, "ActFijoCompsFicha", "WHERE IdActFijo = " & lIdActFijo)
            
      Call ClearGrid(True)
      
   End If
   
   InClick = False
   Call SaveFicha
   
End Sub
Private Sub ClearGrid(Optional ByVal NullVal As Boolean = False)
   Dim i As Integer

   For i = 1 To Grid.rows - 1
      If NullVal Then
         Grid.TextMatrix(i, C_VALOR) = ""
      Else
         Grid.TextMatrix(i, C_VALOR) = Format(0, Grid.TextMatrix(i, C_FORMAT))
      End If
   Next i

   
End Sub
Private Sub Form_Load()

   lInLoad = True
   
   Set lCbComponente = New ClsCombo
   Call lCbComponente.SetControl(Cb_Componente)
   
   Call SetUpGrid
   
   lModif = False
   lIdCompFicha = 0
   lidComp = 0
   
   Call LoadAll
   
   Call EnableForm(Me, lFImported = 0 And gEmpresa.FCierre = 0)
   Call SetRO(Tx_Grupo, True)
   Call SetRO(Tx_Descrip, True)
   Call SetRO(Tx_Cantidad, True)
   
   Cb_Componente.Locked = False
   If gEmpresa.FCierre = 0 Then
      Grid.Locked = False
      Bt_SaveComp.Enabled = True
   End If
   
   If lFromReport <> 0 Then   'sólo se permite que se modifique la componente seleccionada para recargar sólo esa en el reporte
      Cb_Componente.Locked = True
      Bt_AddComp.Enabled = False
      Bt_DelComp.Enabled = False
      Ch_SinDetComps.Enabled = False
   End If
   
   lInLoad = False

   
End Sub

Private Sub SetUpGrid()
   Dim i As Integer

   Grid.Cols = NCOLS + 1
   Grid.FixedCols = 3
   Grid.rows = nRows
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_FIELD) = 0
   Grid.ColWidth(C_EDIT) = 0
   Grid.ColWidth(C_DESC) = 5600
   Grid.ColWidth(C_VALOR) = 2000
   Grid.ColWidth(C_FORMAT) = 0
   
   Grid.ColAlignment(C_DESC) = flexAlignLeftCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_DESC) = "Item"
   Grid.TextMatrix(0, C_VALOR) = "Valor"
   
   Grid.TextMatrix(D_PJEDIVCOMP, C_DESC) = "% División Componenentes"
   Grid.TextMatrix(D_VALORCOMPRA, C_DESC) = "Valor de compra"
   Grid.TextMatrix(D_VALORCRESIDUAL, C_DESC) = "Valor residual"
   Grid.TextMatrix(D_PJEAMORTIZACION, C_DESC) = "% amortización del bien"
   Grid.TextMatrix(D_VIDAUTIL, C_DESC) = "Vida útil (meses)"
   Grid.TextMatrix(D_COSTOSADICIONALES, C_DESC) = "Costos adicionales"
   Grid.TextMatrix(D_TASADESC, C_DESC) = "Tasa de descuento"
   Grid.TextMatrix(D_COSTODESMANT, C_DESC) = "Costo de desmantelamiento"
   Grid.TextMatrix(D_VALACTCOSTODESMANT, C_DESC) = "Valor actual Costos de Desmantelamiento"
   Grid.TextMatrix(D_VALORBIEN, C_DESC) = "Valor del bien"
   Grid.TextMatrix(D_VALORRAZONABLE_31_12, C_DESC) = "Valor razonable al 31/12"
   Grid.TextMatrix(D_OTRASDIFERENCIAS, C_DESC) = "Otras diferencias"
  
   Grid.TextMatrix(D_PJEDIVCOMP, C_FIELD) = "PjeDivComp"
   Grid.TextMatrix(D_VALORCOMPRA, C_FIELD) = "ValorCompra"
   Grid.TextMatrix(D_VALORCRESIDUAL, C_FIELD) = "ValorResidual"
   Grid.TextMatrix(D_PJEAMORTIZACION, C_FIELD) = "PjeAmortizacion"
   Grid.TextMatrix(D_VIDAUTIL, C_FIELD) = "VidaUtil"
   Grid.TextMatrix(D_COSTOSADICIONALES, C_FIELD) = "CostosAdicionales"
   Grid.TextMatrix(D_TASADESC, C_FIELD) = "TasaDesc"
   Grid.TextMatrix(D_COSTODESMANT, C_FIELD) = "CostoDesmant"
   Grid.TextMatrix(D_VALACTCOSTODESMANT, C_FIELD) = "ValActCostoDesmant"
   Grid.TextMatrix(D_VALORBIEN, C_FIELD) = "ValorBien"
   Grid.TextMatrix(D_VALORRAZONABLE_31_12, C_FIELD) = "ValorRazonable_31_12"
   Grid.TextMatrix(D_OTRASDIFERENCIAS, C_FIELD) = "OtrasDiferencias"

   Grid.TextMatrix(D_PJEDIVCOMP, C_EDIT) = "1"
   Grid.TextMatrix(D_VALORCOMPRA, C_EDIT) = "0"
   Grid.TextMatrix(D_VALORCRESIDUAL, C_EDIT) = "1"
   Grid.TextMatrix(D_PJEAMORTIZACION, C_EDIT) = "1"
   Grid.TextMatrix(D_VIDAUTIL, C_EDIT) = "1"
   Grid.TextMatrix(D_COSTOSADICIONALES, C_EDIT) = "0"
   Grid.TextMatrix(D_TASADESC, C_EDIT) = "1"
   Grid.TextMatrix(D_COSTODESMANT, C_EDIT) = "1"
   Grid.TextMatrix(D_VALACTCOSTODESMANT, C_EDIT) = "0"
   Grid.TextMatrix(D_VALORBIEN, C_EDIT) = "0"
   Grid.TextMatrix(D_VALORRAZONABLE_31_12, C_EDIT) = "1"
   Grid.TextMatrix(D_OTRASDIFERENCIAS, C_EDIT) = "1"
   
   Grid.TextMatrix(D_PJEDIVCOMP, C_FORMAT) = DBLFMT1 & "%"
   Grid.TextMatrix(D_VALORCOMPRA, C_FORMAT) = NUMFMT
   Grid.TextMatrix(D_VALORCRESIDUAL, C_FORMAT) = NUMFMT
   Grid.TextMatrix(D_PJEAMORTIZACION, C_FORMAT) = DBLFMT1 & "%"
   Grid.TextMatrix(D_VIDAUTIL, C_FORMAT) = NUMFMT
   Grid.TextMatrix(D_COSTOSADICIONALES, C_FORMAT) = NUMFMT
   Grid.TextMatrix(D_TASADESC, C_FORMAT) = DBLFMT1 & "%"
   Grid.TextMatrix(D_COSTODESMANT, C_FORMAT) = NUMFMT
   Grid.TextMatrix(D_VALACTCOSTODESMANT, C_FORMAT) = NUMFMT
   Grid.TextMatrix(D_VALORBIEN, C_FORMAT) = NUMFMT
   Grid.TextMatrix(D_VALORRAZONABLE_31_12, C_FORMAT) = NUMFMT
   Grid.TextMatrix(D_OTRASDIFERENCIAS, C_FORMAT) = NUMFMT
   
   For i = 1 To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_EDIT)) = 0 Then
         Call FGrSetRowStyle(Grid, i, "B", 0, C_DESC, C_VALOR)
         Call FGrSetRowStyle(Grid, i, "BC", vbButtonFace, C_VALOR, C_VALOR)
      End If
   Next i
         
         
   
End Sub

Private Sub LoadAll()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   
   Q1 = "SELECT ActFijoFicha.IdFicha, ActFijoFicha.IdGrupo, NombGrupo, MovActivoFijo.Descrip, MovActivoFijo.Cantidad, MovActivoFijo.NoDepreciable, SinDetComps "
   Q1 = Q1 & ", PrecioFactura, DerechosIntern, Transporte, ObrasAdapt, FImported"
   Q1 = Q1 & " FROM (MovActivoFijo LEFT JOIN ActFijoFicha ON ActFijoFicha.IdActfijo =  MovActivoFijo.IdActFijo )"
   Q1 = Q1 & " LEFT JOIN AFGrupos ON ActFijoFicha.IdGrupo =  AFGrupos.IdGrupo AND AFGrupos.IdEmpresa = ActFijoFicha.IdEmpresa "
   Q1 = Q1 & " WHERE MovActivoFijo.IdActFijo = " & lIdActFijo
   Q1 = Q1 & " AND MovActivoFijo.IdEmpresa = " & gEmpresa.id & " AND MovActivoFijo.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      
      lFImported = vFld(Rs("FImported"))    'tiene que estar primero para evitar CalcFicha
      
      
      Tx_Grupo = vFld(Rs("NombGrupo"))
      Tx_Descrip = vFld(Rs("Descrip"))
      Tx_Cantidad = Format(vFld(Rs("Cantidad")), NUMFMT)
      lIdFicha = vFld(Rs("IdFicha"))
      lIdGrupo = vFld(Rs("IdGrupo"))
      
      lPrecioFactura = vFld(Rs("PrecioFactura"))
      lDerechosIntern = vFld(Rs("DerechosIntern"))
      lTransporte = vFld(Rs("Transporte"))
      lObrasAdapt = vFld(Rs("ObrasAdapt"))
      
      lAFNoDepreciable = vFld(Rs("NoDepreciable"))
      
      Ch_SinDetComps = IIf(vFld(Rs("SinDetComps")) <> 0, 1, 0)
      
      
      lModif = False
      

   End If
   
   Call CloseRs(Rs)
   
   If Ch_SinDetComps <> 0 Then
      Call LoadDetFin(-1)

   Else
      Call LoadComps
   End If
      
   Q1 = "SELECT NombComp, IdComp "
   Q1 = Q1 & " FROM AFComponentes "
   Q1 = Q1 & " WHERE IdGrupo = " & lIdGrupo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " ORDER BY NombComp "
   
   Call FillCombo(Ls_CompGrupo, DbMain, Q1, 0)
   
   If lFImported <> 0 Then   'viene del año anterior
   
      Grid.TextMatrix(D_PJEDIVCOMP, C_EDIT) = "0"
      Grid.TextMatrix(D_VALORCOMPRA, C_EDIT) = "0"
      Grid.TextMatrix(D_VALORCRESIDUAL, C_EDIT) = "1"
      Grid.TextMatrix(D_PJEAMORTIZACION, C_EDIT) = "0"
      Grid.TextMatrix(D_VIDAUTIL, C_EDIT) = "1"
      Grid.TextMatrix(D_COSTOSADICIONALES, C_EDIT) = "0"
      Grid.TextMatrix(D_TASADESC, C_EDIT) = "0"
      Grid.TextMatrix(D_COSTODESMANT, C_EDIT) = "0"
      Grid.TextMatrix(D_VALACTCOSTODESMANT, C_EDIT) = "0"
      Grid.TextMatrix(D_VALORBIEN, C_EDIT) = "0"
      Grid.TextMatrix(D_VALORRAZONABLE_31_12, C_EDIT) = "1"
      Grid.TextMatrix(D_OTRASDIFERENCIAS, C_EDIT) = "0"
      
      For i = 1 To Grid.rows - 1
         If Val(Grid.TextMatrix(i, C_EDIT)) = 0 Then
            Call FGrSetRowStyle(Grid, i, "B", 0, C_DESC, C_VALOR)
            Call FGrSetRowStyle(Grid, i, "BC", vbButtonFace, C_VALOR, C_VALOR)
         End If
      Next i
      
      Grid.TextMatrix(D_VALORBIEN, C_DESC) = "Valor del bien actualizado"
   Else
      Call SetUpGrid
   End If
      
End Sub
Private Sub LoadComps()
   Dim Rs As Recordset
   Dim Q1 As String
   
   lCbComponente.Clear
   
   Q1 = "SELECT NombComp, IdCompFicha, ActFijoCompsFicha.IdComp "
   Q1 = Q1 & " FROM ActFijoCompsFicha INNER JOIN AFComponentes ON ActFijoCompsFicha.IdComp = AFComponentes.IdComp AND ActFijoCompsFicha.IdEmpresa = AFComponentes.IdEmpresa "
   Q1 = Q1 & " WHERE IdActFijo = " & lIdActFijo
   Q1 = Q1 & " AND ActFijoCompsFicha.IdEmpresa = " & gEmpresa.id & " AND ActFijoCompsFicha.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY NombComp "
      
   If lInLoad And lFEditIdCompFicha > 0 Then
      Call lCbComponente.FillCombo(DbMain, Q1, lFEditIdCompFicha)
   Else
      Call lCbComponente.FillCombo(DbMain, Q1, -1)
   End If

End Sub

Private Function Valida(Optional ByVal Oper As Integer = O_EDIT) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   Dim TotPje As Double, PjeComps As Double
   
   Valida = False
 
 
   If Not Oper = O_NEW Then
      
      If vFmt(Grid.TextMatrix(D_VALORCOMPRA, C_VALOR)) > 0 And vFmt(Grid.TextMatrix(D_VALORCRESIDUAL, C_VALOR)) > vFmt(Grid.TextMatrix(D_VALORCOMPRA, C_VALOR)) Then
         MsgBox1 "El valor residual debe ser inferior al valor de compra del bien.", vbExclamation
         Exit Function
      End If
     
      If vFmt(Grid.TextMatrix(D_VALORRAZONABLE_31_12, C_VALOR)) < 0 And Ch_NoExisteValRazonable = 0 Then
         MsgBox1 "El valor razonable debe ser mayor o igual que cero.", vbExclamation
         Exit Function
      End If
      
      'vemos si el total del porcentaje de división de componentes no supera el 100%
      
      If Ch_SinDetComps <> 0 Then
         If vFmt(Grid.TextMatrix(D_PJEDIVCOMP, C_VALOR)) <> 1 Then
            MsgBox1 "El porcentaje de División Componenentes debe ser 100% dado que no hay detalle de componentes.", vbExclamation
            Exit Function
         End If
      
      Else
                  
         'obtenemos la suma de porcentajes de las componentes distintas a la actual
         Q1 = "SELECT Sum(PjeDivComp) FROM ActFijoCompsFicha WHERE IdActFijo = " & lIdActFijo & " AND IdComp <> " & lidComp
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         
         If Not Rs.EOF Then
            PjeComps = Round(vFld(Rs(0)), 4)
         End If
         
         Call CloseRs(Rs)
         
         TotPje = PjeComps + vFmt(Grid.TextMatrix(D_PJEDIVCOMP, C_VALOR))
         
         If TotPje > 1 Then
            MsgBox1 "El porcentaje de División Componenentes supera el 100% si se consideran todas las componentes." & vbCrLf & vbCrLf & "El valor sugerido para este porcentaje es " & Format(1 - PjeComps, Grid.TextMatrix(D_PJEDIVCOMP, C_FORMAT)), vbExclamation
            Exit Function
         ElseIf TotPje < 1 Then
            MsgBox1 "ATENCIÖN: El porcentaje de División Componenentes es inferior al 100% si se consideran todas las componentes." & vbCrLf & vbCrLf & "El valor sugerido para este porcentaje es " & Format(1 - TotPje + vFmt(Grid.TextMatrix(D_PJEDIVCOMP, C_VALOR)), Grid.TextMatrix(D_PJEDIVCOMP, C_FORMAT)), vbInformation
            'Exit Function      'no se pone esto porque si no no puede salir de la componente
         End If
               
      End If
      
   End If
   
   Valida = True
   
End Function
Private Sub SaveCompAF()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim id As Long, IdComp As Long
   Dim Fld As String, ValFld As String
   Dim i As Integer
   Dim FldArray(4) As AdvTbAddNew_t
   
   If Ch_SinDetComps = 0 And lCbComponente.ListIndex < 0 Then
      Exit Sub
   End If
      
   'Id = lCbComponente.ItemData    'no sirve porque tenemos que guardar la anterior, si el usuario hizo click en la lista de componentes
   id = lIdCompFicha
   
   If id = 0 Then      'es nueva
   
      For i = 1 To Grid.rows - 1
'         Fld = Fld & "," & Grid.TextMatrix(i, C_FIELD)
'         ValFld = ValFld & "," & Str(vFmt(Grid.TextMatrix(i, C_VALOR)))
         Fld = Fld & ", " & Grid.TextMatrix(i, C_FIELD) & " = " & str(vFmt(Grid.TextMatrix(i, C_VALOR)))
      Next i
      
      If Ch_SinDetComps <> 0 Then
         IdComp = -1
      Else
         IdComp = lCbComponente.Matrix(L_IDCOMP)
      End If
      

      FldArray(0).FldName = "IdActFijo"
      FldArray(0).FldValue = lIdActFijo
      FldArray(0).FldIsNum = True
      
      FldArray(1).FldName = "IdGrupo"
      FldArray(1).FldValue = lIdGrupo
      FldArray(1).FldIsNum = True
            
      FldArray(2).FldName = "IdComp"
      FldArray(2).FldValue = IdComp
      FldArray(2).FldIsNum = True
            
      FldArray(3).FldName = "IdEmpresa"
      FldArray(3).FldValue = gEmpresa.id
      FldArray(3).FldIsNum = True
                  
      FldArray(4).FldName = "Ano"
      FldArray(4).FldValue = gEmpresa.Ano
      FldArray(4).FldIsNum = True
      
      lIdCompFicha = AdvTbAddNewMult(DbMain, "ActFijoCompsFicha", "IdCompFicha", FldArray)
      
'      If lIdCompFicha > 0 Then
'         Q1 = "UPDATE ActFijoCompsFicha SET"
'         Q1 = Q1 & "  IdActFijo = " & lIdActFijo
'         Q1 = Q1 & ", IdGrupo = " & lIdGrupo
'         Q1 = Q1 & ", IdComp = " & IdComp
'         Q1 = Q1 & ", IdEmpresa = " & gEmpresa.id
'         Q1 = Q1 & ", Ano = " & gEmpresa.Ano
'         Q1 = Q1 & Fld
'         Q1 = Q1 & " WHERE IdCompFicha = " & lIdCompFicha
'
'         Call ExecSQL(DbMain, Q1)
'      End If
      
      Call LoadComps
      If Ch_SinDetComps <> 0 Then
         LoadDetFin (-1)
      Else
         id = lIdCompFicha
         Call lCbComponente.SelItem(id)
      End If
      
      
'      Q1 = "INSERT INTO ActFijoCompsFicha ( IdActFijo, IdGrupo, IdComp, IdEmpresa, Ano " & Fld & ")"
'      Q1 = Q1 & " VALUES (" & lIdActFijo & "," & lIdGrupo & "," & IdComp & "," & gEmpresa.id & "," & gEmpresa.Ano & ValFld & ")"
'
'      Call ExecSQL(DbMain, Q1)
      
'      Call LoadComps
'
'      Q1 = "SELECT IdCompFicha FROM ActFijoCompsFicha WHERE IdActFijo = " & lIdActFijo & " AND IdGrupo = " & lIdGrupo & " AND IdComp = " & IdComp
'
'      Set Rs = OpenRs(DbMain, Q1)
'      If Not Rs.EOF Then
'         If Ch_SinDetComps <> 0 Then
'            LoadDetFin (-1)
'         Else
'            id = vFld(Rs("IdCompFicha"))
'            Call lCbComponente.SelItem(id)
'         End If
'      End If
'
'      Call CloseRs(Rs)
      
   Else   'ya existe
      Q1 = "UPDATE ActFijoCompsFicha SET "
      
      For i = Grid.FixedRows To Grid.rows - 1
         If lFImported = 0 Or (lFImported <> 0 And Grid.TextMatrix(i, C_EDIT) <> 0) Then
            Q1 = Q1 & Grid.TextMatrix(i, C_FIELD) & " = " & str(vFmt(Grid.TextMatrix(i, C_VALOR))) & ","
         End If
      Next i
      
      If Right(Q1, 1) = "," Then
         Q1 = Left(Q1, Len(Q1) - 1)
         Q1 = Q1 & " WHERE IdCompFicha = " & id
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   'actualizamos NoExisteValRazonable
   Q1 = "UPDATE ActFijoCompsFicha SET "
   Q1 = Q1 & "  NoExisteValRazonable = " & IIf(Ch_NoExisteValRazonable <> 0, 1, 0)
   Q1 = Q1 & ", ValorRazonable_31_12 = " & IIf(Ch_NoExisteValRazonable <> 0, "NULL", str(vFmt(Grid.TextMatrix(D_VALORRAZONABLE_31_12, C_VALOR))))
   Q1 = Q1 & " WHERE IdCompFicha = " & id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
      
      
   lModif = False
   Call SaveFicha
End Sub

Private Sub SaveFicha()
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "UPDATE ActFijoFicha SET SinDetComps = " & Int(Ch_SinDetComps <> 0) & " WHERE IdFicha = " & lIdFicha
   Call ExecSQL(DbMain, Q1)
   
   lSave = True
   
End Sub
Private Sub Cb_Componente_Click()
   Dim id As Long
   Dim i As Integer
   Static InClick As Boolean
   
   If InClick Then
      Exit Sub
   End If
   
   If lCbComponente.ListIndex < 0 Then
      Exit Sub
   End If

   If Not lInLoad And lModif Then
      If MsgBox1("¿Desea guardar los cambios realizados a este detalle?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         
         If Not Valida() Then
            InClick = True
            Call lCbComponente.SelItem(lIdCompFicha)
            DoEvents
            InClick = False
            Exit Sub
         End If

         Call SaveCompAF
         
      End If
   End If
   
   For i = 1 To Grid.rows - 1
      Grid.TextMatrix(i, C_VALOR) = 0
   Next i
   
   id = lCbComponente.ItemData
   
   If id > 0 Then
   
      Call LoadDetFin(id)
      
   End If
   
   lIdCompFicha = id
   lidComp = lCbComponente.Matrix(L_IDCOMP)
   
   lModif = False

End Sub
Private Sub LoadDetFin(ByVal IdCompFicha As Long)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Fld As String
   Dim i As Integer
   Dim Wh As String

   If IdCompFicha = 0 Then
      Exit Sub
   End If

   For i = 1 To Grid.rows - 1
      Fld = Fld & "," & Grid.TextMatrix(i, C_FIELD)
   Next i
   
   If Len(Fld) > 0 Then
      Fld = Mid(Fld, 2)
   End If
   
   If IdCompFicha > 0 Then    'si IdCompFicha = -1 es sin detalle de componentes, por lo tanto es un solo registro
      Wh = " WHERE IdCompFicha = " & IdCompFicha
   Else
      Wh = " WHERE IdActFijo = " & lIdActFijo
   End If
   
   Q1 = "SELECT IdCompFicha, " & Fld & ", NoExisteValRazonable FROM ActFijoCompsFicha " & Wh
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
   
      For i = 1 To Grid.rows - 1
         Grid.TextMatrix(i, C_VALOR) = Format(vFld(Rs(Grid.TextMatrix(i, C_FIELD))), Grid.TextMatrix(i, C_FORMAT))
      Next i
      
      If IdCompFicha < 0 Then    'si IdCompFicha = -1 es sin detalle de componentes, por lo tanto es un solo registro
         lIdCompFicha = vFld(Rs("IdCompFicha"))
         lidComp = -1
      End If
      
      Ch_NoExisteValRazonable = IIf(vFld(Rs("NoExisteValRazonable")) <> 0, 1, 0)
         
      
   ElseIf IdCompFicha < 0 Then
      lModif = True     'si es sin detalle de componentes, se asigna porcentaje y se calcula en la carga inicial, de esta manera se valida y se graba
   
   End If
   
   Call CloseRs(Rs)
   
   If lAFNoDepreciable <> 0 Then    'si es no depreciable se bloquea Valor Residual
      Grid.TextMatrix(D_VALORCRESIDUAL, C_EDIT) = 0
      
      If vFmt(Grid.TextMatrix(D_VALORCRESIDUAL, C_VALOR)) <> 0 Then
         Grid.TextMatrix(D_VALORCRESIDUAL, C_VALOR) = 0
         lModif = True
      End If
      
      Call FGrSetRowStyle(Grid, D_VALORCRESIDUAL, "B", 0, C_DESC, C_VALOR)
      Call FGrSetRowStyle(Grid, D_VALORCRESIDUAL, "BC", vbButtonFace, C_VALOR, C_VALOR)
   End If
   
   Call CalcFicha

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   
   If Ch_SinDetComps = 0 And lCbComponente.ListIndex < 0 Then
      Exit Sub
   End If
   
   If Col <> C_VALOR Or Val(Grid.TextMatrix(Row, C_EDIT)) = 0 Then
      Exit Sub
   End If
   
   EdType = FEG_Edit
      
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)

   If Row <> D_OTRASDIFERENCIAS And vFmt(Value) < 0 Then
      Action = vbCancel
      
   Else
      If InStr(Grid.TextMatrix(Row, C_FORMAT), "%") > 0 Then
         If InStr(Value, "%") > 0 Then
            Value = vFmt(Value)
         Else
            Value = vFmt(Value) / 100
         End If
         
         
         If vFmt(Value) > 1 Or vFmt(Value) < 0 Then
            MsgBox1 "Valor de porcentaje inválido.", vbExclamation
            Action = vbCancel
            Exit Sub
         End If
           
      End If
      Value = Format(vFmt(Value), Grid.TextMatrix(Row, C_FORMAT))
      Grid.TextMatrix(Row, C_VALOR) = Value

      Action = vbOK
      lModif = True
      Call CalcFicha
   End If

End Sub


Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Row = D_PJEDIVCOMP Or Row = D_PJEAMORTIZACION Or Row = D_TASADESC Then
      Call KeyDec(KeyAscii)
   Else
      Call KeyNum(KeyAscii)
   End If
   
End Sub

Private Sub Ls_CompGrupo_DblClick()
   Dim id As Long, i As Integer
   
   id = CbItemData(Ls_CompGrupo)
   
   If id <= 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   'primero vemos si hay una componente modificada y grabamos
   If Not lInLoad And lModif Then
      Ls_CompGrupo.visible = False
      If MsgBox1("¿Desea guardar los cambios realizados a este detalle?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         If Not Valida() Then
            Exit Sub
         End If

         Call SaveCompAF
      End If
   End If

   'vemos si ya está
   For i = 0 To lCbComponente.ListCount - 1
      If Val(lCbComponente.Matrix(L_IDCOMP, i)) = id Then     'ya existe
         MsgBox1 "Esta componente ya está asociada al Activo Fijo.", vbExclamation
         Ls_CompGrupo.visible = False
         Exit Sub
      End If
   Next i
   
   'no está, la agregamos
   
   Call lCbComponente.AddItem(Ls_CompGrupo, 0, id, "", True)
   Ls_CompGrupo.visible = False
   
   If Not Valida(O_NEW) Then
      Exit Sub
   Else
      Call SaveCompAF
   End If
   
End Sub

Private Sub Ls_CompGrupo_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call Ls_CompGrupo_DblClick
   End If
   
End Sub

Private Sub Ls_CompGrupo_LostFocus()
   Ls_CompGrupo.visible = False
End Sub

Private Function CalcFicha()
   Dim PjeDivComp As Double, TasaDesc As Double
   Dim CostoDesmant As Double, ValActCostoDesmant As Double
   Dim VidaUtil As Single
   Dim kk As Double
   
   If lFImported <> 0 Then  'viene de años anteriores, no se recalcula porque el vaor del bien viene del año anterior con factor
      Exit Function
   End If
   
   PjeDivComp = vFmt(Grid.TextMatrix(D_PJEDIVCOMP, C_VALOR))
   Grid.TextMatrix(D_VALORCOMPRA, C_VALOR) = Format(lPrecioFactura * PjeDivComp, NUMFMT)
   Grid.TextMatrix(D_COSTOSADICIONALES, C_VALOR) = Format((lDerechosIntern + lTransporte + lObrasAdapt) * PjeDivComp, NUMFMT)
   
   CostoDesmant = vFmt(Grid.TextMatrix(D_COSTODESMANT, C_VALOR))
   TasaDesc = vFmt(Grid.TextMatrix(D_TASADESC, C_VALOR))
   VidaUtil = vFmt(Grid.TextMatrix(D_VIDAUTIL, C_VALOR))
   ValActCostoDesmant = CostoDesmant * ((1 / (1 + TasaDesc)) ^ (VidaUtil / 12))
   Grid.TextMatrix(D_VALACTCOSTODESMANT, C_VALOR) = Format(ValActCostoDesmant, NUMFMT)
   
'   Grid.TextMatrix(D_VALORBIEN, C_VALOR) = Format((vFmt(Grid.TextMatrix(D_VALORCOMPRA, C_VALOR)) - vFmt(Grid.TextMatrix(D_VALORCRESIDUAL, C_VALOR)) + vFmt(Grid.TextMatrix(D_COSTOSADICIONALES, C_VALOR)) + vFmt(Grid.TextMatrix(D_VALACTCOSTODESMANT, C_VALOR))) * vFmt(Tx_Cantidad), NUMFMT)
   Grid.TextMatrix(D_VALORBIEN, C_VALOR) = Format((vFmt(Grid.TextMatrix(D_VALORCOMPRA, C_VALOR)) + vFmt(Grid.TextMatrix(D_COSTOSADICIONALES, C_VALOR)) + vFmt(Grid.TextMatrix(D_VALACTCOSTODESMANT, C_VALOR))) * vFmt(Tx_Cantidad), NUMFMT)   'Thomson Reuters Joshua 4 jul 2018
   

End Function

