VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmEmpresas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empresas"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   Icon            =   "FrmEmpresas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fr_Sort 
      Caption         =   "Ordenar por"
      Height          =   975
      Left            =   7260
      TabIndex        =   8
      Top             =   3660
      Width           =   1275
      Begin VB.OptionButton Op_SortRUT 
         Caption         =   "RUT"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   795
      End
      Begin VB.OptionButton Op_SortNombre 
         Caption         =   "Nombre"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton Bt_Del 
      Caption         =   "&Eliminar"
      Height          =   735
      Left            =   7260
      MousePointer    =   99  'Custom
      Picture         =   "FrmEmpresas.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Eliminar empresa"
      Top             =   2520
      Width           =   1200
   End
   Begin VB.CommandButton Bt_Ren 
      Caption         =   "&Modificar"
      Height          =   735
      Left            =   7260
      MousePointer    =   99  'Custom
      Picture         =   "FrmEmpresas.frx":066E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Modificar empresa"
      Top             =   1740
      Width           =   1200
   End
   Begin VB.CommandButton Bt_New 
      Caption         =   "&Nueva"
      Height          =   735
      Left            =   7260
      MousePointer    =   99  'Custom
      Picture         =   "FrmEmpresas.frx":0C41
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nueva empresa"
      Top             =   960
      Width           =   1200
   End
   Begin VB.CommandButton bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   7260
      TabIndex        =   6
      Top             =   540
      Width           =   1200
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   4635
      Left            =   1500
      TabIndex        =   0
      Top             =   540
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8176
      Cols            =   4
      Rows            =   20
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
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   570
      Index           =   0
      Left            =   420
      Picture         =   "FrmEmpresas.frx":11D3
      Top             =   540
      Width           =   570
   End
   Begin VB.Label La_demo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "DEMO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002A01A6&
      Height          =   330
      Left            =   7200
      TabIndex        =   7
      Top             =   4860
      Visible         =   0   'False
      Width           =   885
   End
End
Attribute VB_Name = "FrmEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_RUT = 0
Const C_NOMBRECORTO = 1
Const C_ID = 2
Const C_ESTADO = 3    'no Activo

Private Sub bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Del_Click()
   Dim RUT As String, NCorto As String
   Dim id As Long
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Ano As Integer
   Dim DbName As String
   Dim SeDel As Boolean
   Dim sWhere As String
      
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Trim(Grid.TextMatrix(Row, C_RUT)) = "" Then
      Exit Sub
   End If
   
   id = Grid.TextMatrix(Row, C_ID)
   RUT = Grid.TextMatrix(Row, C_RUT)
   NCorto = Grid.TextMatrix(Row, C_NOMBRECORTO)

   If MsgBox1("¡ ATENCION !" & vbNewLine & vbNewLine & " A continuación se eliminaran todos los datos asociados a la empresa " & NCorto & " y no podrá recuperarlos. Para esto asegurese que ningún usuario este trabajando con esta empresa." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   SeDel = True
   Me.MousePointer = vbHourglass
   
#If DATACON = 1 Then
   
   Q1 = "SELECT Ano FROM EmpresasAno WHERE IdEmpresa=" & id
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      Ano = vFld(Rs("Ano"))
      If Ano > 0 Then
      
         DbName = gDbPath & "\Empresas\" & Ano & "\" & vFmtCID(RUT) & ".mdb"
         If ExistFile(DbName) Then
            If KillFile(DbName) = False Then
               MsgBox1 "¡ ATENCION !" & vbNewLine & vbNewLine & "No puede eliminar el año " & Ano & " de la empresa " & NCorto & ", porque esta abierta por algún usuario.", vbExclamation
               SeDel = False
               Exit Do
            End If
            
         End If
         
         DbName = gDbPath & "\Empresas\" & Ano & "\B_" & vFmtCID(RUT) & ".mdb"
         If ExistFile(DbName) Then
            Call KillFile(DbName)
         End If
         
'         Call ExecSQL(DbMain, "DELETE * FROM EmpresasAno WHERE IdEmpresa=" & id & " AND Ano=" & Ano)
         Call DeleteSQL(DbMain, "EmpresasAno", " WHERE IdEmpresa=" & id)
      End If
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
#Else
   sWhere = " WHERE IdEmpresa = " & id

   Call DeleteSQL(DbMain, "ActFijoCompsFicha", sWhere)
   Call DeleteSQL(DbMain, "ActFijoFicha", sWhere)
   Call DeleteSQL(DbMain, "AjustesExtLibCaja", sWhere)
   Call DeleteSQL(DbMain, "AsistImpPrimCat", sWhere)
   Call DeleteSQL(DbMain, "BaseImponible14Ter", sWhere)
   Call DeleteSQL(DbMain, "Cartola", sWhere)
   Call DeleteSQL(DbMain, "Comprobante", sWhere)
   Call DeleteSQL(DbMain, "ControlEmpresa", sWhere)
   Call DeleteSQL(DbMain, "CtasAjustesExCont", sWhere)
   Call DeleteSQL(DbMain, "Cuentas", sWhere)
   Call DeleteSQL(DbMain, "CuentasBasicas", sWhere)
   Call DeleteSQL(DbMain, "DetCartola", sWhere)
   Call DeleteSQL(DbMain, "DetSaldosAp", sWhere)
   Call DeleteSQL(DbMain, "DocCuotas", sWhere)
   Call DeleteSQL(DbMain, "DetCartola", sWhere)
   Call DeleteSQL(DbMain, "Documento", sWhere)
   Call DeleteSQL(DbMain, "Empresa", " WHERE Id= " & id)
   Call DeleteSQL(DbMain, "EstadoMes", sWhere)
   Call DeleteSQL(DbMain, "EmpresasAno", sWhere)
   Call DeleteSQL(DbMain, "ImpAdic", sWhere)
   Call DeleteSQL(DbMain, "InfoAnualDJ1847", sWhere)
   Call DeleteSQL(DbMain, "LibroCaja", sWhere)
   Call DeleteSQL(DbMain, "LogComprobantes", sWhere)
   Call DeleteSQL(DbMain, "LogImpreso", sWhere)
   Call DeleteSQL(DbMain, "MovActivoFijo", sWhere)
   Call DeleteSQL(DbMain, "MovComprobante", sWhere)
   Call DeleteSQL(DbMain, "MovDocumento", sWhere)
   Call DeleteSQL(DbMain, "ParamEmpresa", sWhere)
   Call DeleteSQL(DbMain, "PropIVA_TotMensual", sWhere)
   Call DeleteSQL(DbMain, "Socios", sWhere)
   
   Call DeleteSQL(DbMain, "AFComponentes", sWhere)
   Call DeleteSQL(DbMain, "AFGrupos", sWhere)
   Call DeleteSQL(DbMain, "AreaNegocio", sWhere)
   Call DeleteSQL(DbMain, "CentroCosto", sWhere)
   Call DeleteSQL(DbMain, "CT_Comprobante", sWhere)
   Call DeleteSQL(DbMain, "CT_MovComprobante", sWhere)
   Call DeleteSQL(DbMain, "CuentasRazon", sWhere)
   Call DeleteSQL(DbMain, "Entidades", sWhere)
   Call DeleteSQL(DbMain, "Glosas", sWhere)
   Call DeleteSQL(DbMain, "Notas", sWhere)
   Call DeleteSQL(DbMain, "Notas", sWhere)
   Call DeleteSQL(DbMain, "ParamEmpresa", sWhere)
   Call DeleteSQL(DbMain, "ParamRazon", sWhere)
   Call DeleteSQL(DbMain, "Sucursales", sWhere)


#End If

   
   If SeDel Then
      Call DeleteSQL(DbMain, "Empresas", " WHERE IdEmpresa=" & id)
      Call DeleteSQL(DbMain, "UsuarioEmpresa", " WHERE IdEmpresa=" & id)
      
      'Franca 23/06/2005   (se elimina la línea en vez de dejarla tamaño cero, porque si se agrega una empresa inmediatamente después, la pone en la línea de tamaño cero y no se ve en la lista
      
'      Grid.RowHeight(Row) = 0
'      Grid.Rows = Grid.Rows + 1
'
'      Grid.TextMatrix(Row, C_ID) = ""
'      Grid.TextMatrix(Row, C_RUT) = ""
'      Grid.TextMatrix(Row, C_NOMBRECORTO) = ""

      Grid.RemoveItem (Row)
      
      Call FGrVRows(Grid)
      
      MsgBox1 "La empresa ha sido eliminada.", vbInformation
      
   End If
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Bt_New_Click()
   Dim Frm As FrmMantEmpresa
   Dim RUT As String, NCorto As String
   Dim id As Long
   Dim Row As Integer, r As Integer, n As Integer
   
   ' pam: 13 dic 2010
   n = 0
   If gAppCode.NivProd <> VER_ILIM Then
      For r = Grid.FixedRows To Grid.Rows - 1
         If Grid.TextMatrix(r, C_ID) <> "" Then
            n = n + 1
         End If
      Next r
   
      If (gAppCode.NivProd = VER_DEMO And n >= 3) Then
         MsgBox1 "Usted ya tiene la cantidad de empresas permitidas para la demo de este sistema.", vbInformation
         Exit Sub
      
      ElseIf n >= gMaxEmpLicencia Then
         MsgBox1 "Usted ya tiene la cantidad de empresas permitidas por el tipo de licencia que ha adquirido (" & gMaxEmpLicencia & " empresas).", vbInformation
         Exit Sub
      End If
   End If
   
   Set Frm = New FrmMantEmpresa
   If Frm.FNew(id, RUT, NCorto) = vbOK Then
      Row = FGrAddRow(Grid)
      Grid.TextMatrix(Row, C_NOMBRECORTO) = NCorto
      Grid.TextMatrix(Row, C_ID) = id
      Grid.TextMatrix(Row, C_RUT) = FmtCID(RUT)
      Grid.TextMatrix(Row, C_ESTADO) = "Si"
      
   End If
   
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Ren_Click()
   Dim Frm As FrmMantEmpresa
   Dim RUT As String, NCorto As String
   Dim id As Long
   Dim Row As Integer
   Dim Estado As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Trim(Grid.TextMatrix(Row, C_RUT)) = "" Then
      Exit Sub
   End If
   
   id = Grid.TextMatrix(Row, C_ID)
   RUT = Grid.TextMatrix(Row, C_RUT)
   NCorto = Grid.TextMatrix(Row, C_NOMBRECORTO)
   
   Set Frm = New FrmMantEmpresa
   Estado = IIf(LCase(Grid.TextMatrix(Row, C_ESTADO)) = "si", 0, 1)
   
   If Frm.FEdit(id, RUT, NCorto, Estado) = vbOK Then
      Grid.TextMatrix(Row, C_NOMBRECORTO) = NCorto
      Grid.TextMatrix(Row, C_ID) = id
      Grid.TextMatrix(Row, C_RUT) = FmtCID(RUT)
      Grid.TextMatrix(Row, C_ESTADO) = IIf(Estado = 0, "Si", "No")
      
   End If
   
   Set Frm = Nothing
End Sub

Private Sub Form_Load()
   Call SetUpGrid
   Call LoadAll
   
'   If gAppCode.Demo Then
'      Bt_New.Enabled = False
'   End If
   
  ' Bt_Del.Enabled = ChkVMant(VMANT_2005) se dejo para todos =
   
   La_demo.Visible = gAppCode.Demo
   
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_RUT) = 1500
   Grid.ColWidth(C_NOMBRECORTO) = 2500
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_ESTADO) = 1000
      
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_NOMBRECORTO) = flexAlignLeftCenter
   Grid.ColAlignment(C_ESTADO) = flexAlignCenterCenter
   
   Grid.TextMatrix(0, C_RUT) = "RUT"
   Grid.TextMatrix(0, C_NOMBRECORTO) = "Nombre Corto"
   Grid.TextMatrix(0, C_ESTADO) = "Activa"
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   
   Q1 = "SELECT IdEmpresa, Rut, NombreCorto, Estado FROM Empresas"
   If gAppCode.Demo Then
      Q1 = Q1 & " WHERE RUT IN ('1','2','3')"
   End If
   
   If Op_SortNombre Then
      Q1 = Q1 & " ORDER BY NombreCorto"
   Else
      Q1 = Q1 & " ORDER BY " & SqlVal("Rut")
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Row = 1
   Grid.Rows = Row
   Do While Rs.EOF = False
      Grid.Rows = Row + 1
      
      Grid.TextMatrix(Row, C_RUT) = FmtCID(vFld(Rs("Rut")))
      Grid.TextMatrix(Row, C_NOMBRECORTO) = vFld(Rs("NombreCorto"))
      Grid.TextMatrix(Row, C_ESTADO) = IIf(vFld(Rs("Estado")) = 0, "Si", "No")
      
      Grid.Row = Row
      Grid.TextMatrix(Row, C_ID) = vFld(Rs("IdEmpresa"))
      
      Row = Row + 1
      
      If gAppCode.NivProd = VER_5EMP And Row > 5 Then
         Exit Do
      End If
      
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
      
End Sub
Private Sub Grid_DblClick()
   Call Bt_Ren_Click
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Grid, Me.Caption)
   End If
End Sub

Private Sub Op_SortNombre_Click()

   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault

End Sub

Private Sub Op_SortRUT_Click()
   
   Me.MousePointer = vbHourglass
   Call LoadAll
   Me.MousePointer = vbDefault

End Sub
