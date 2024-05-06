VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmEstadoMeses 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado Meses"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "FrmEstadoMeses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   300
      Picture         =   "FrmEstadoMeses.frx":000C
      ScaleHeight     =   630
      ScaleWidth      =   690
      TabIndex        =   6
      Top             =   300
      Width           =   690
   End
   Begin VB.CommandButton Bt_Close 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   315
      Left            =   4500
      TabIndex        =   3
      Top             =   360
      Width           =   1035
   End
   Begin VB.CommandButton Bt_CerrarMes 
      Caption         =   "Cerrar Mes"
      Height          =   735
      Left            =   4500
      Picture         =   "FrmEstadoMeses.frx":0697
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1035
   End
   Begin VB.CommandButton Bt_AbrirMes 
      Caption         =   "Abrir Mes"
      Height          =   735
      Left            =   4500
      Picture         =   "FrmEstadoMeses.frx":0C6B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   900
      Width           =   1035
   End
   Begin VB.PictureBox Pc_RedArrow 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      Picture         =   "FrmEstadoMeses.frx":127F
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   3720
      Width           =   225
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3315
      Left            =   1320
      TabIndex        =   0
      Top             =   300
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   5847
      _Version        =   393216
      Cols            =   5
      FixedCols       =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Último mes con datos"
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   3720
      Width           =   1635
   End
End
Attribute VB_Name = "FrmEstadoMeses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_NUMMES = 0
Const C_ULTIMOMES = 1
Const C_NOMBMES = 2
Const C_ESTADO = 3
Const C_IDESTADO = 4

Private Sub Bt_AbrirMes_Click()
   Dim Mes As Integer
   Dim i As Integer
   Dim Q1 As String

   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Mes = Val(Grid.TextMatrix(Grid.Row, C_NUMMES))
   
   If Mes = 0 Then
      Exit Sub
   End If

   If Not gAbrirMesesParalelo Then

      For i = 1 To Grid.rows - 1
         If Val(Grid.TextMatrix(i, C_IDESTADO)) = EM_ABIERTO Then
            MsgBox1 "Para abrir este mes, debe antes cerrar el mes de " & gNomMes(i) & ".", vbExclamation + vbOKOnly
            Exit Sub
         End If
      Next i
   
   End If
   
   If AbrirMes(Mes) = True Then
      Grid.TextMatrix(Grid.Row, C_IDESTADO) = EM_ABIERTO
      Grid.TextMatrix(Grid.Row, C_ESTADO) = gEstadoMes(EM_ABIERTO)
      MsgBox1 "El mes de " & gNomMes(Mes) & " fue abierto con éxito.", vbInformation + vbOKOnly
   End If
   
End Sub

Private Sub Bt_CerrarMes_Click()
   Dim Mes As Integer

   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Mes = Val(Grid.TextMatrix(Grid.Row, C_NUMMES))
   
   If Mes = 0 Then
      Exit Sub
   End If

   If CerrarMes(Mes) = True Then
      Grid.TextMatrix(Grid.Row, C_IDESTADO) = EM_CERRADO
      Grid.TextMatrix(Grid.Row, C_ESTADO) = gEstadoMes(EM_CERRADO)
      MsgBox1 "El mes de " & gNomMes(Mes) & " fue cerrado con éxito.", vbInformation + vbOKOnly
   End If
   
End Sub

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   
   Call SetUpGrid
   
   Call LoadMeses
   
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   
   Call SetupPriv
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_NUMMES) = 0
   Grid.ColWidth(C_ULTIMOMES) = 200
   Grid.ColWidth(C_NOMBMES) = 1200
   Grid.ColWidth(C_ESTADO) = 1200
   Grid.ColWidth(C_IDESTADO) = 0
   
   For i = 0 To Grid.Cols - 1
      Grid.FixedAlignment(i) = flexAlignCenterCenter
   Next i
   
   Grid.ColAlignment(C_ULTIMOMES) = flexAlignCenterCenter
   Grid.ColAlignment(C_NOMBMES) = flexAlignLeftCenter
   Grid.ColAlignment(C_ESTADO) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_NOMBMES) = "Mes"
   Grid.TextMatrix(0, C_ESTADO) = "Estado"

   Grid.rows = Grid.FixedRows + 12
   
End Sub

Private Sub LoadMeses()
   Dim i As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim MesAbierto As Integer
   Dim UltMes As Integer
   
   AddLog ("Estamos el LoadMeses")
      
   UltMes = GetUltimoMesConMovs(True)
   
   AddLog ("UltMes=" & UltMes)
   
   For i = 1 To 12
      Grid.TextMatrix(i, C_NUMMES) = i
      Grid.TextMatrix(i, C_NOMBMES) = gNomMes(i)
      
      AddLog ("Agrgamos Mes =" & gNomMes(i))
      
      If i = UltMes Then
         Call FGrSetPicture(Grid, i, C_ULTIMOMES, Pc_RedArrow)
         AddLog ("Marcamos Mes =" & gNomMes(i))
      End If

   Next i

   Q1 = "SELECT Mes, Estado FROM EstadoMes "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Mes"
   Set Rs = OpenRs(DbMain, Q1)
   AddLog ("Consulta estado meses")
   
   i = Grid.FixedRows
   
   Do While Rs.EOF = False
      Grid.TextMatrix(i, C_ESTADO) = gEstadoMes(vFld(Rs("Estado")))
      Grid.TextMatrix(i, C_IDESTADO) = vFld(Rs("Estado"))
      AddLog ("Estado mes " & i & " " & gEstadoMes(vFld(Rs("Estado"))))
      
      If vFld(Rs("Estado")) = EM_ABIERTO Then
         MesAbierto = i
         AddLog ("Mes Abierto = " & i)
      End If
      
      i = i + 1
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   AddLog ("Cerramos Rs")

   If MesAbierto > 0 Then
      Call FGrSelRow(Grid, MesAbierto, False)
   Else
      Call FGrSelRow(Grid, Grid.FixedRows, False)
   End If
   
   AddLog ("Nos vamos")
      
End Sub
Private Function SetupPriv()
   
   If Not ChkPriv(PRV_ADM_EMPRESA) Then
      Call EnableForm(Me, False)
   End If
   
End Function

