VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmUsuarioPriv 
   Caption         =   "Privilegios de Usuarios por Empresa"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   Icon            =   "FrmUsuarioPriv.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_Ok 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   7620
      TabIndex        =   5
      Top             =   540
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7620
      TabIndex        =   4
      Top             =   900
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3315
      Left            =   1560
      TabIndex        =   3
      Top             =   1380
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5847
      _Version        =   393216
      Rows            =   20
      Cols            =   6
      FixedRows       =   2
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   1560
      TabIndex        =   0
      Top             =   420
      Width           =   4875
      Begin VB.ComboBox Cb_Usuarios 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   3315
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   420
      Picture         =   "FrmUsuarioPriv.frx":000C
      Top             =   480
      Width           =   765
   End
End
Attribute VB_Name = "FrmUsuarioPriv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_EMPRESA = 0
Const C_IDPERFIL = 1
Const C_UPDATE = 2
Const C_ESTADO = 3
Const C_IDEMPRESA = 4
Const C_INIPERF = 5

Const R_IDPERF = 1

Dim lModUsr As Boolean
Dim lIdUsr As Long
Dim lUsrName As String

Dim lcbUsuarios As ClsCombo

Private Sub SetUpGrid()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim ColPerf As Integer, ColWid As Integer, c As Integer
   
   Call FGrSetup(Grid)
   
   Q1 = "SELECT Nombre,idPerfil FROM Perfiles ORDER BY Nombre"
   Set Rs = OpenRs(DbMain, Q1)
   
   ColPerf = C_INIPERF
   Do While Rs.EOF = False
      Grid.Cols = ColPerf + 1
      
      c = Me.TextWidth(vFld(Rs("Nombre"))) + Screen.TwipsPerPixelX * 6
      If c > ColWid Then
         ColWid = c
      End If
      
      Grid.ColWidth(ColPerf) = ColWid
      Grid.FixedAlignment(ColPerf) = flexAlignCenterCenter
      Grid.ColAlignment(ColPerf) = flexAlignCenterCenter
      Grid.TextMatrix(0, ColPerf) = vFld(Rs("Nombre"))
      Grid.TextMatrix(R_IDPERF, ColPerf) = vFld(Rs("idPerfil"))
        
      ColPerf = ColPerf + 1
   
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   
   For c = C_INIPERF To ColPerf - 1
      Grid.ColWidth(c) = ColWid
   Next c
   
   Grid.RowHeight(R_IDPERF) = 0
   Grid.ColWidth(C_EMPRESA) = 2700
   Grid.FixedAlignment(C_EMPRESA) = flexAlignCenterCenter
   Grid.ColAlignment(C_EMPRESA) = flexAlignLeftCenter
   Grid.TextMatrix(0, C_EMPRESA) = "Empresa"
   
   Grid.ColWidth(C_ESTADO) = 0
   Grid.ColWidth(C_IDEMPRESA) = 0
   Grid.ColWidth(C_IDPERFIL) = 0
   Grid.ColWidth(C_UPDATE) = 0
   
End Sub

Private Sub bt_Cancelar_Click()
   Unload Me
End Sub
Private Sub bt_OK_Click()

   Call SaveUsr
   
   Unload Me
End Sub

Private Sub Cb_Usuarios_Click()
   If lModUsr Then
      If MsgBox1("¿Desea grabar las modificaciones realizadas al usuario " & lUsrName & "?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call SaveUsr
      End If
   End If

   Call LoadAll
End Sub

Private Sub Form_Load()
   Dim Q1 As String, Rs As Recordset
   
   'LLeno combo usuario
   Set lcbUsuarios = New ClsCombo
   Call lcbUsuarios.SetControl(Cb_Usuarios)
   
   Call SetUpGrid
   
   Q1 = "SELECT Usuario, idUsuario FROM Usuarios WHERE Usuario <> '" & gAdmUser & "' ORDER BY Usuario"
   Call lcbUsuarios.FillCombo(DbMain, Q1, -1)
   
   'Call LoadAll
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer, n As Integer
   Dim Col As Integer, idPerf As Integer, idEmp As Integer
   
   Q1 = "SELECT idEmpresa, NombreCorto FROM Empresas"
   If gAppCode.Demo Then
      Q1 = Q1 & " WHERE RUT IN ('1','2','3')"
   End If
   Q1 = Q1 & " ORDER BY NombreCorto"
   Set Rs = OpenRs(DbMain, Q1)
   
   Row = 2
   Grid.Rows = 2
   n = 0
   Do While Rs.EOF = False
      Grid.Rows = Row + 1
      
      n = n + 1
      Grid.TextMatrix(Row, C_EMPRESA) = vFld(Rs("NombreCorto"), True)
      Grid.TextMatrix(Row, C_IDEMPRESA) = vFld(Rs("idEmpresa"))
            
      If gAppCode.NivProd = VER_5EMP And n >= 5 Then
         Exit Do
      End If
            
      Rs.MoveNext
      Row = Row + 1
   Loop
   Call CloseRs(Rs)
   
   Q1 = "SELECT idEmpresa, idPerfil FROM UsuarioEmpresa WHERE idUsuario=" & lcbUsuarios.ItemData
   Set Rs = OpenRs(DbMain, Q1)
   Do Until Rs.EOF
      idPerf = vFld(Rs("idPerfil"))
      idEmp = vFld(Rs("idEmpresa"))
      For Row = R_IDPERF + 1 To Grid.Rows - 1
         If Val(Grid.TextMatrix(Row, C_IDEMPRESA)) = idEmp Then
            For Col = C_INIPERF To Grid.Cols - 1
               If idPerf = Val(Grid.TextMatrix(R_IDPERF, Col)) Then
                  Grid.TextMatrix(Row, Col) = "x"
                  Grid.TextMatrix(Row, C_IDPERFIL) = idPerf
                  Grid.TextMatrix(Row, C_UPDATE) = idPerf
                  Exit For
               End If
            Next Col
         End If

      Next Row
   
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   
   lModUsr = False
   lIdUsr = lcbUsuarios.ItemData
   lUsrName = lcbUsuarios
   
   Call FGrVRows(Grid)
   
End Sub

Private Sub Form_Resize()
   Dim d As Integer

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   d = Me.Width - (Bt_OK.Left + Bt_OK.Width)
   If d > 1000 Then
      Grid.Width = d
   End If
   
   d = Me.Height - Grid.Top - W.YCaption * 3
   If d > 1000 Then
      Grid.Height = d
   Else
      Me.Height = Grid.Top + 1000 + W.YCaption * 2
   End If
   
   d = FGrVRows(Grid) + 1
   If Grid.Rows < d Then
      Grid.Rows = d
   End If
End Sub

Private Sub Grid_DblClick()
   Dim Col As Integer, c As Integer
   Dim Row As Integer
   
   If Grid.Col < C_INIPERF Then
      Exit Sub
   End If
   
   Row = Grid.Row
   If Grid.TextMatrix(Row, C_IDEMPRESA) = "" Then
      Exit Sub
   End If
      
   c = 0
   For Col = C_INIPERF To Grid.Cols - 1
      If Grid.TextMatrix(Row, Col) = "x" Then
         Grid.TextMatrix(Row, Col) = ""
         c = Col
      End If
   Next Col

   If c <> Grid.Col Then ' seleccionó otro perfil

      If Grid.TextMatrix(Row, Grid.Col) = "x" Then
         Grid.TextMatrix(Row, Grid.Col) = ""
      Else
         Call FGrForeColor(Grid, Row, Grid.Col, vbBlue)
         Grid.TextMatrix(Row, Grid.Col) = "x"
      End If
      
      Grid.TextMatrix(Row, C_IDPERFIL) = Grid.TextMatrix(R_IDPERF, Grid.Col)
      Call FGrModRow(Grid, Row, FGR_U, C_UPDATE, C_ESTADO)
   
   Else  ' lo dejó sin perfil
      Call FGrModRow(Grid, Row, FGR_D, C_UPDATE, C_ESTADO, False)
      Grid.TextMatrix(Row, C_IDPERFIL) = ""
   End If
   
   lModUsr = True
  
   'Cb_Usuarios.Enabled = False
   
End Sub

Private Sub SaveUsr()
   Dim Q1 As String
   Dim Row As Integer
   
   For Row = R_IDPERF + 1 To Grid.Rows - 1
      
      If Trim(Grid.TextMatrix(Row, C_EMPRESA)) <> "" And Grid.TextMatrix(Row, C_ESTADO) <> "" Then
         
         If Grid.TextMatrix(Row, C_ESTADO) = FGR_I Then
            
            Q1 = "INSERT INTO UsuarioEmpresa (idUsuario,idEmpresa,idPerfil) VALUES("
            Q1 = Q1 & lIdUsr & "," & Grid.TextMatrix(Row, C_IDEMPRESA)
            Q1 = Q1 & "," & Grid.TextMatrix(Row, C_IDPERFIL) & ")"
            
            Call ExecSQL(DbMain, Q1)
            
         ElseIf Grid.TextMatrix(Row, C_ESTADO) = FGR_U Then
            
            Q1 = "UPDATE UsuarioEmpresa SET idPerfil=" & Grid.TextMatrix(Row, C_IDPERFIL)
            Q1 = Q1 & " WHERE idUsuario=" & lIdUsr
            Q1 = Q1 & " AND idEmpresa=" & Grid.TextMatrix(Row, C_IDEMPRESA)
            
            Call ExecSQL(DbMain, Q1)
            
         ElseIf Grid.TextMatrix(Row, C_ESTADO) = FGR_D Then
            
            Q1 = " WHERE idUsuario=" & lIdUsr
            Q1 = Q1 & " AND idEmpresa=" & Grid.TextMatrix(Row, C_IDEMPRESA)
            
            Call DeleteSQL(DbMain, "UsuarioEmpresa", Q1)
            
         End If
      End If
   Next Row

End Sub
