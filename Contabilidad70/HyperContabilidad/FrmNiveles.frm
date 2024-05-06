VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmNiveles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Niveles para Plan de Cuenta"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "FrmNiveles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   480
      Picture         =   "FrmNiveles.frx":000C
      ScaleHeight     =   630
      ScaleWidth      =   600
      TabIndex        =   7
      Top             =   480
      Width           =   600
   End
   Begin VB.CommandButton bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6120
      TabIndex        =   4
      Top             =   900
      Width           =   1035
   End
   Begin VB.CommandButton bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   6120
      TabIndex        =   3
      Top             =   540
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   4455
      Begin VB.CommandButton bt_Definir 
         Caption         =   "&Definir"
         Height          =   315
         Left            =   2940
         TabIndex        =   2
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox tx_Niveles 
         Height          =   315
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "5"
         Top             =   360
         Width           =   315
      End
      Begin FlexEdGrid2.FEd2Grid Grid 
         Height          =   1515
         Left            =   720
         TabIndex        =   1
         Top             =   960
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   2672
         Cols            =   2
         Rows            =   6
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
      Begin VB.Label Label1 
         Caption         =   "Total de Niveles : (1-5)"
         Height          =   195
         Left            =   720
         TabIndex        =   6
         Top             =   420
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_NIVELES = 0
Const C_DIGITOS = 1

Private Sub SetUpGrid()
   Dim Row As Integer
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_NIVELES) = 700
   Grid.ColWidth(C_DIGITOS) = 1200
   
   Grid.FixedAlignment(C_NIVELES) = flexAlignCenterCenter
   Grid.FixedAlignment(C_DIGITOS) = flexAlignCenterCenter
   
   Grid.ColAlignment(C_NIVELES) = flexAlignRightCenter
   Grid.ColAlignment(C_DIGITOS) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_NIVELES) = "Nivel"
   Grid.TextMatrix(0, C_DIGITOS) = "Largo (1-5)"
  
   
End Sub

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_Definir_Click()
   Dim Row As Integer
   
   If vFmt(tx_Niveles) = 0 Then
      MsgBox1 "Total de niveles debe ser mayor a cero", vbExclamation
      Exit Sub
   End If
   
   If vFmt(tx_Niveles) > 5 Then
      MsgBox1 "Total de niveles menor o igual a cinco", vbExclamation
      Exit Sub
   End If
      
   If vFmt(tx_Niveles) < MAX_NIVELES Then
      For Row = vFmt(tx_Niveles) + 1 To MAX_NIVELES
         Grid.TextMatrix(Row, C_NIVELES) = ""
         Grid.TextMatrix(Row, C_DIGITOS) = ""
      Next Row
   End If
   
   For Row = 1 To vFmt(tx_Niveles)
      Grid.TextMatrix(Row, C_NIVELES) = Row
   Next Row
   
   'Call EnableFrm(False)      'Franca (26 Ene 2006) Si se cambian los niveles no deja cambiar los dígitos por nivel
      
End Sub

Private Sub Bt_OK_Click()
   Dim Q1 As String
   Dim Row As Integer
   
   If vFmt(tx_Niveles) = 0 Then
      MsgBox1 "Debe ingresar número de niveles", vbExclamation
      Exit Sub
   End If
   
   'Franca (26 Ene 2006): vemos si la cantidad de niveles corresponde al detalle de los dígitos por nivel
   For Row = 1 To MAX_NIVELES
      If Grid.TextMatrix(Row, C_NIVELES) = "" Then
         If vFmt(tx_Niveles) <> Row - 1 Then
            MsgBox1 "Presione el botón Definir para ajustar los niveles y los dígitos por nivel.", vbExclamation
            Exit Sub
         Else
            Exit For
         End If
      End If
   Next Row
   
   For Row = 1 To tx_Niveles
   
      If Trim(Grid.TextMatrix(Row, C_NIVELES)) <> "" And vFmt(Grid.TextMatrix(Row, C_DIGITOS)) <= 0 Then
         MsgBox1 "Debe ingresar digitos para el nivel " & Row & ".", vbExclamation
         Exit Sub
      End If
                  
   Next Row
   
   For Row = 1 To MAX_NIVELES
      Q1 = "UPDATE ParamEmpresa SET Valor=" & vFmt(Grid.TextMatrix(Row, C_DIGITOS))
      Q1 = Q1 & " WHERE Tipo='DIGNIV" & Row & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
      
   Next Row
   
   Q1 = "UPDATE ParamEmpresa SET Valor=" & vFmt(tx_Niveles) & " WHERE Tipo='NIVELES'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   Call ConfigNiveles
   Unload Me
   
End Sub

Private Sub Form_Load()
   Call SetUpGrid
   Call LoadAll
   Call LockForm
   Call SetupPriv
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   If vFmt(Value) > 5 Then
      MsgBox1 "El total de dígitos por nivel debe ser menor o igual a cinco.", vbExclamation
      Value = ""
      Exit Sub
   End If
   
   
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FEG2_EdType)
   If Col = C_NIVELES Then
      Exit Sub
   End If
   
   If Trim(Grid.TextMatrix(Row, C_NIVELES)) = "" Then
      Exit Sub
   End If
   
   EdType = FEG_Edit
      Grid.TxBox.MaxLength = 1
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub

Private Sub tx_Niveles_Change()
   Call EnableFrm(True)
End Sub

Private Sub tx_Niveles_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
   
End Sub
Private Sub EnableFrm(bool As Boolean)
   Grid.Locked = Not bool
   Bt_OK.Enabled = bool
   bt_Definir.Enabled = bool
   
End Sub
Private Sub LockForm()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim bool As Boolean
   
   bool = True
   
   'Cheque que no tenga plan de cuentas, de lo contrario no puede modificar los niveles
   Q1 = "SELECT Count(*) as n FROM Cuentas "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If vFld(Rs("n")) <> 0 Then
      MsgBox1 "No puede cambiar la cantidad de niveles porque ya existen cuentas creadas.", vbExclamation
      bool = False
   End If
   Call CloseRs(Rs)

   Bt_OK.Enabled = bool
   bt_Definir.Enabled = bool
   
   Call SetTxRO(tx_Niveles, Not bool)
   Grid.Locked = Not bool
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT Codigo, Tipo, Valor FROM ParamEmpresa WHERE (Tipo='NIVELES' OR " & GenLike(DbMain, "DIGNIV", "Tipo") & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
      If vFld(Rs("Tipo")) = "NIVELES" Then
         tx_Niveles = vFld(Rs("Valor"))
      ElseIf vFld(Rs("Valor")) <> 0 Then
         Grid.TextMatrix(vFld(Rs("Codigo")), C_DIGITOS) = vFld(Rs("Valor"))
         Grid.TextMatrix(vFld(Rs("Codigo")), C_NIVELES) = vFld(Rs("Codigo"))
      End If
      
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
      
End Sub

Private Function SetupPriv()
   
   If Not ChkPriv(PRV_ADM_CTAS) Then
      Call EnableFrm(False)
   End If
   
End Function



