VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmResetEmprAno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminar Empresa-Año"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "FrmResetEmprAno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "FrmResetEmprAno.frx":000C
      Top             =   3600
      Width           =   6495
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   6660
      TabIndex        =   2
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Remove 
      Caption         =   "Eliminar"
      Height          =   315
      Left            =   6660
      TabIndex        =   1
      Top             =   480
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3015
      Left            =   1380
      TabIndex        =   0
      Top             =   420
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   25
      Cols            =   4
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   360
      Picture         =   "FrmResetEmprAno.frx":00A8
      Top             =   480
      Width           =   720
   End
End
Attribute VB_Name = "FrmResetEmprAno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_RUT = 0
Private Const C_EMP = 1
Private Const C_ANO = 2
Private Const C_IDEMP = 3


Private Sub LoadAll()
   Dim Q1 As String, Rs As Recordset, r As Integer
   
   Q1 = "SELECT EmpresasAno.idEmpresa, Empresas.NombreCorto, Empresas.Rut, Max(EmpresasAno.Ano) AS Ano"
   Q1 = Q1 & " FROM EmpresasAno INNER JOIN Empresas ON EmpresasAno.idEmpresa = Empresas.IdEmpresa"
   Q1 = Q1 & " WHERE EmpresasAno.Ano > 0"
   If gAppCode.Demo Then
      Q1 = Q1 & " AND RUT IN ('1','2','3')"
   End If
   Q1 = Q1 & " GROUP BY EmpresasAno.idEmpresa, Empresas.NombreCorto, Empresas.Rut"
   Q1 = Q1 & " ORDER BY Empresas.NombreCorto"

   r = 0
   Set Rs = OpenRs(DbMain, Q1)
   Do Until Rs.EOF
      r = r + 1
      
      Grid.rows = r + 1
      Grid.TextMatrix(r, C_RUT) = FmtRut(vFld(Rs("RUT")))
      Grid.TextMatrix(r, C_EMP) = vFld(Rs("NombreCorto"))
      Grid.TextMatrix(r, C_ANO) = vFld(Rs("Ano"))
      Grid.TextMatrix(r, C_IDEMP) = vFld(Rs("idEmpresa"))
      
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
      
   Call FGrVRows(Grid)
   
End Sub

Private Sub SetUpGrid()

   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_RUT) = 1200
   Grid.ColWidth(C_EMP) = 3000
   Grid.ColWidth(C_ANO) = 600
   Grid.ColWidth(C_IDEMP) = 0
   
   Grid.TextMatrix(0, C_RUT) = "R.U.T."
   Grid.TextMatrix(0, C_EMP) = "Nombre"
   Grid.TextMatrix(0, C_ANO) = "Año"
   Grid.TextMatrix(0, C_IDEMP) = ""


End Sub

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_Remove_Click()
   Dim idEmp As Long, Q1 As String, Rc As Long, Ano As Integer, Rut As Long
   Dim TbDoc As String
   Dim sWhere As String
   Dim Rs As Recordset
   Dim InitAno As String
   Dim fname As String
   
   idEmp = Val(Grid.TextMatrix(Grid.Row, C_IDEMP))
   If idEmp <= 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If

   On Error Resume Next

   Ano = Val(Grid.TextMatrix(Grid.Row, C_ANO))
   Rut = vFmtRut(Grid.TextMatrix(Grid.Row, C_RUT))

   Q1 = "¡ATENCIÓN!" & vbCrLf & "¿Está seguro que desea eliminar el año " & Ano & " de la empresa " & Grid.TextMatrix(Grid.Row, C_EMP) & " ?"
   If MsgBox1(Q1, vbYesNo Or vbDefaultButton2 Or vbQuestion) <> vbYes Then
      Exit Sub
   End If

   sWhere = " WHERE IdEmpresa = " & idEmp & " AND Ano = " & Ano

#If DATACON = 1 Then
   Q1 = gDbPath & "\Empresas\" & Ano - 1 & "\" & Rut & ".mdb"
   If ExistFile(Q1) Then

      TbDoc = "tmp_Doc" & idEmp
      
      Q1 = "PWD=" & PASSW_PREFIX & Rut & ";"
      
      Rc = LinkMdbTable(DbMain, gDbPath & "\Empresas\" & Ano - 1 & "\" & Rut & ".mdb", "Documento", TbDoc, , , Q1)
      If Rc = 0 Then
         Exit Sub
      End If
   
      Q1 = "UPDATE " & TbDoc & " SET FExported=0 WHERE FExported<>0"
      Rc = ExecSQL(DbMain, Q1)
      
   End If
   
   fname = gDbPath & "\Empresas\" & Ano & "\" & Rut & ".mdb"
   If ExistFile(fname) Then
      Call KillFile(fname)
      If Err Then
         MsgErr fname
      End If
   End If
   
#Else
   
   Q1 = "SELECT * FROM ParamEmpresa " & sWhere & " AND Tipo = 'INITAÑO'"
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      InitAno = vFld(Rs("Valor"))
   End If
   Call CloseRs(Rs)


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
   Call DeleteSQL(DbMain, "Empresa", "WHERE Id= " & idEmp & " AND Ano=" & Ano)
   Call DeleteSQL(DbMain, "EstadoMes", sWhere)
   Call DeleteSQL(DbMain, "ImpAdic", sWhere)
   Call DeleteSQL(DbMain, "InfoAnualDJ1847", sWhere)
   Call DeleteSQL(DbMain, "LibroCaja", sWhere)
   Call DeleteSQL(DbMain, "LogComprobantes", sWhere)
   Call DeleteSQL(DbMain, "LogImpreso", sWhere)
   Call DeleteSQL(DbMain, "MovActivoFijo", sWhere)
   Call DeleteSQL(DbMain, "MovComprobante", sWhere)
   Call DeleteSQL(DbMain, "ParamEmpresa", sWhere)
   Call DeleteSQL(DbMain, "PropIVA_TotMensual", sWhere)
   Call DeleteSQL(DbMain, "Socios", sWhere)
    
   If InitAno = "EMPHISTACC" Then   'viene de una base de datos Access, hay que eliminar todos los documentos de la empresa, incluso de años anteriores
      Call DeleteSQL(DbMain, "Documento", "WHERE IdEmpresa = " & idEmp)
      Call DeleteSQL(DbMain, "MovDocumento", "WHERE IdEmpresa = " & idEmp)
    
   Else
      Call DeleteSQL(DbMain, "Documento", sWhere)
      Call DeleteSQL(DbMain, "MovDocumento", sWhere)
      
   End If
   
#End If
    
'   Q1 = "DELETE * FROM EmpresasAno WHERE idEmpresa=" & idEmp & " AND Ano=" & Ano
   Rc = DeleteSQL(DbMain, "EmpresasAno", sWhere)
   
   
   MsgBox1 "El año ha sido eliminado.", vbInformation
   
   Grid.RemoveItem Grid.Row
   
End Sub

Private Sub Form_Load()

   Call SetUpGrid
   
   Call LoadAll

End Sub
