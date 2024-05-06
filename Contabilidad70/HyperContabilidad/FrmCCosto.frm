VERSION 5.00
Begin VB.Form FrmCCosto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Centro de Gestión"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "FrmCCosto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   1500
      TabIndex        =   5
      Top             =   420
      Width           =   6315
      Begin VB.CheckBox Ch_Vigente 
         Caption         =   "Vigente"
         Height          =   195
         Left            =   1380
         TabIndex        =   2
         Top             =   1260
         Width           =   2235
      End
      Begin VB.TextBox Tx_Descripcion 
         Height          =   315
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   1
         Top             =   780
         Width           =   4455
      End
      Begin VB.TextBox Tx_Codigo 
         Height          =   315
         Left            =   1380
         MaxLength       =   15
         TabIndex        =   0
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   7
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   6
         Top             =   840
         Width           =   885
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8100
      TabIndex        =   4
      Top             =   840
      Width           =   1155
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   8100
      TabIndex        =   3
      Top             =   480
      Width           =   1155
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   780
      Left            =   300
      Picture         =   "FrmCCosto.frx":000C
      Top             =   540
      Width           =   750
   End
End
Attribute VB_Name = "FrmCCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lRc As Integer
Dim Oper As Integer
Dim lCCosto As CCosto_t

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub bt_OK_Click()
   Dim Rs As Recordset
   Dim Q1 As String
      
   If Valida() = False Then
      Exit Sub
   End If
   
   If Oper = O_NEW Then
        
      lCCosto.Id = AdvTbAddNew(DbMain, "CentroCosto", "idCCosto", "IdEmpresa", gEmpresa.Id)
   
   End If
   
   Q1 = "UPDATE CentroCosto SET Codigo='" & ParaSQL(Tx_Codigo) & "'"
   Q1 = Q1 & ", Descripcion='" & ParaSQL(Tx_Descripcion) & "'"
   Q1 = Q1 & ", Vigente = " & IIf(Ch_Vigente <> 0, -1, 0)
   Q1 = Q1 & " WHERE idCCosto=" & lCCosto.Id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.Id
   
   
   Call ExecSQL(DbMain, Q1)
   
   lCCosto.Codigo = Tx_Codigo
   lCCosto.Descrip = Tx_Descripcion
   
   lRc = vbOK
   Unload Me
End Sub
Friend Function FView(CCosto As CCosto_t) As Integer
   lCCosto = CCosto
   Oper = O_VIEW
   Me.Show vbModal
   
   FView = lRc
   CCosto = lCCosto
   
End Function
Friend Function FEdit(CCosto As CCosto_t) As Integer
   lCCosto = CCosto
   Oper = O_EDIT
   Me.Show vbModal
   
   CCosto = lCCosto
   FEdit = lRc
End Function

Friend Function FNew(CCosto As CCosto_t) As Integer
   Oper = O_NEW
   Me.Show vbModal
   
   CCosto = lCCosto
   FNew = lRc
End Function

Private Sub Form_Load()
   lRc = vbCancel
   
   If Oper = O_NEW Then
      Caption = "Nuevo Centro de Gestión"
      Ch_Vigente = 1
   ElseIf Oper = O_EDIT Then
      Caption = "Modificar Centro de Gestión"
      Call LoadAll
   ElseIf Oper = O_VIEW Then
      Caption = "Ver Centro de Gestión"
      Call LoadAll
      
   End If
   
   Call SetupPriv
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT Codigo, Descripcion, Vigente FROM CentroCosto"
   Q1 = Q1 & " WHERE idCCosto=" & lCCosto.Id
   
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      Tx_Codigo = vFld(Rs("Codigo"), True)
      Tx_Descripcion = vFld(Rs("Descripcion"), True)
      Ch_Vigente = IIf(vFld(Rs("Vigente")) <> 0, 1, 0)
      
   End If
   Call CloseRs(Rs)
   
End Sub
Private Function Valida() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   Valida = False
   
   If Tx_Codigo = "" Then
      MsgBox1 "Debe ingresar el código", vbExclamation
      Exit Function
      
   End If
   
   Q1 = "SELECT idCCosto FROM CentroCosto WHERE Codigo='" & Tx_Codigo & "'"
   Q1 = Q1 & " AND idCCosto<>" & lCCosto.Id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.Id

   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      MsgBox1 "Código ya existe", vbExclamation
      Tx_Codigo.SetFocus
      Call CloseRs(Rs)
      Exit Function
      
   End If
   Call CloseRs(Rs)
   
   Valida = True
End Function
Private Sub Tx_Codigo_KeyPress(KeyAscii As Integer)
   Call KeyUpper(KeyAscii)
   
End Sub
Private Function SetupPriv()
   
   If Not ChkPriv(PRV_ADM_DEF) Then
      Call EnableForm(Me, False)
   End If
   
End Function
