VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFactoresAct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porcentajes de Actualización Corrección Monetaria"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10305
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   8700
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8700
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Bt_GetFactoresSII 
      Caption         =   "Obtener Factores SII"
      Height          =   315
      Left            =   8040
      TabIndex        =   5
      ToolTipText     =   "Obtener Punto IPC del mes seleccionado"
      Top             =   5040
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   480
      Picture         =   "FrmFactoresAct.frx":0000
      ScaleHeight     =   630
      ScaleWidth      =   660
      TabIndex        =   3
      Top             =   480
      Width           =   660
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   5955
      Begin VB.TextBox Tx_Ano 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   345
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Año:"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   1380
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6165
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmFactoresAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_ID = 0
Const C_IDMES = 1
Const C_MES = 2
Const C_MES1 = 3

Const NCOLS = C_MES1 + 12 - 1


Dim lAno As Integer
Dim lFactores(12, 12) As SII_Fact_t

Public Function FEdit(ByVal Ano As Integer) As Integer

   lAno = Ano
   
   Me.Show vbModal

End Function

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub


Private Sub Bt_GetFactoresSII_Click()
   Dim Rc As Integer

   If lAno < 2013 Then
      MsgBox1 "No es posible obtener esta información desde el sitio del SII para años anteriores al 2013", vbInformation
      Bt_OK.Enabled = False
      Exit Sub
   End If
   
   Rc = SII_CorrMonetAnual(lAno, lFactores)
   If Rc <> 0 Then
      MsgBox1 "Error al obtener factores desde el sitio del SII (" & Rc & ").", vbExclamation
      Exit Sub
   End If
   
   Call LoadFactoresSII
   
End Sub

Private Sub bt_OK_Click()
   Call SaveAll
   Unload Me
End Sub

Private Sub Form_Load()

   Call SetUpGrid
   
   Tx_Ano = lAno

   Call Load_Factores
   
End Sub

Private Sub SetUpGrid()
   Dim i As Integer

   Grid.Cols = NCOLS + 1
   Grid.FixedRows = 1
   Grid.FixedCols = C_MES1
   Grid.rows = Grid.FixedRows + 1 + 12
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_IDMES) = 0
   Grid.ColWidth(C_MES) = 2000
      
   Grid.TextMatrix(Grid.FixedRows, C_MES) = "Capital Inicial"
   For i = Grid.FixedRows + 1 To Grid.rows - 1
      Grid.TextMatrix(i, C_IDMES) = i - Grid.FixedRows
      Grid.TextMatrix(i, C_MES) = gNomMes(Val(Grid.TextMatrix(i, C_IDMES)))
   Next i
      
   For i = C_MES1 To Grid.Cols - 1
      Grid.ColWidth(i) = 600
      Grid.ColAlignment(i) = flexAlignRightCenter
      Grid.TextMatrix(0, i) = Left(gNomMes(i - 2), 3)
   Next i
   
   
   
End Sub
Private Sub LoadFactoresSII()
   Dim i As Integer, j As Integer
   
   For i = 0 To 12
      For j = 1 To 12
         Grid.TextMatrix(i + Grid.FixedRows, j + C_MES1 - 1) = IIf(lFactores(i, j).bFact, Format(lFactores(i, j).Fact, DBLFMT3), "")
      Next j
   Next i
      

End Sub
Private Sub Load_Factores()
   Dim i As Integer, j As Integer
   Dim Fact(12, 12) As SII_Fact_t
   
   If lAno <> gEmpresa.Ano Then
      Call ReadFactorActAnual(lAno, Fact)
   End If
      
   For i = 0 To 12
      For j = 1 To 12
         If lAno = gEmpresa.Ano Then
            Grid.TextMatrix(i + Grid.FixedRows, j + C_MES1 - 1) = IIf(gFactorActAnual(i, j).bFact, Format(gFactorActAnual(i, j).Fact, DBLFMT3), "")
         Else
            Grid.TextMatrix(i + Grid.FixedRows, j + C_MES1 - 1) = IIf(Fact(i, j).bFact, Format(Fact(i, j).Fact, DBLFMT3), "")
         End If
      Next j
   Next i

End Sub

Private Sub LoadFactorActAnual(ByVal Ano As Integer)
   Dim Fact(12, 12) As SII_Fact_t
   
   Call ReadFactorActAnual(Ano, Fact)
   
   
End Sub
Private Sub SaveAll()
   Dim i As Integer, j As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   
   For i = 0 To 12
      For j = 1 To 12
      
         Q1 = "SELECT Factor FROM FactorActAnual WHERE Ano = " & lAno & " AND MesRow = " & i & " AND MesCol = " & j
         Set Rs = OpenRs(DbMain, Q1)
                  
         If Grid.TextMatrix(i + Grid.FixedRows, j + C_MES1 - 1) <> "" Then
         
            If Not Rs.EOF Then
            
               Q1 = "UPDATE FactorActAnual SET Factor = " & IIf(Grid.TextMatrix(i + Grid.FixedRows, j + C_MES1 - 1) = "", "NULL", str(vFmt(Grid.TextMatrix(i + Grid.FixedRows, j + C_MES1 - 1))))
               Q1 = Q1 & " WHERE Ano = " & lAno & " AND MesRow = " & i & " AND MesCol = " & j
                        
            Else
               Q1 = "INSERT INTO FactorActAnual (Ano, MesRow, MesCol, Factor) "
               Q1 = Q1 & "VALUES (" & lAno & ", " & i & ", " & j & ", "
               Q1 = Q1 & IIf(Grid.TextMatrix(i + Grid.FixedRows, j + C_MES1 - 1) = "", "NULL", str(vFmt(Grid.TextMatrix(i + Grid.FixedRows, j + C_MES1 - 1)))) & ")"
            
            End If
            
            Call ExecSQL(DbMain, Q1)
            
         ElseIf Not Rs.EOF Then
            Q1 = " WHERE Ano = " & lAno & " AND MesRow = " & i & " AND MesCol = " & j
            Call DeleteSQL(DbMain, "FactorActAnual", Q1)
         
         End If
         Call CloseRs(Rs)

      Next j
   Next i
   
   If lAno = gEmpresa.Ano Then
      Call ReadFactorActAnual(gEmpresa.Ano, gFactorActAnual)
   End If

End Sub
