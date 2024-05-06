VERSION 5.00
Begin VB.Form FrmSucursal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sucursal"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "FrmSucursal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   8055
      Begin VB.CheckBox Ch_Vigente 
         Caption         =   "Vigente"
         Height          =   195
         Left            =   660
         TabIndex        =   2
         Top             =   1560
         Width           =   1875
      End
      Begin VB.CommandButton Bt_Cancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   6300
         TabIndex        =   4
         Top             =   960
         Width           =   1155
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   6300
         TabIndex        =   3
         Top             =   600
         Width           =   1155
      End
      Begin VB.TextBox Tx_Codigo 
         Height          =   315
         Left            =   1620
         MaxLength       =   15
         TabIndex        =   0
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Tx_Descripcion 
         Height          =   315
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   1
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   7
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   6
         Top             =   660
         Width           =   540
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   420
      Picture         =   "FrmSucursal.frx":000C
      Top             =   600
      Width           =   660
   End
End
Attribute VB_Name = "FrmSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lRc As Integer
Dim Oper As Integer
Dim lSucursal As Sucursal_t

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
'      Set Rs = DbMain.OpenRecordset("Sucursales", dbOpenTable)
'      Rs.AddNew
'
'      lSucursal.id = Rs("idSucursal")
'      Rs("Codigo") = ParaSQL(Tx_Codigo)
'      Rs("Descripcion") = ParaSQL(Tx_Descripcion)
'
'      Rs.Update
'      Rs.Close
        
      lSucursal.id = AdvTbAddNew(DbMain, "Sucursales", "idSucursal", "IdEmpresa", gEmpresa.id)
            
   End If
   
   Q1 = "UPDATE Sucursales SET Codigo='" & ParaSQL(Tx_Codigo)
   Q1 = Q1 & "', Descripcion='" & ParaSQL(Tx_Descripcion) & "'"
   Q1 = Q1 & ", Vigente = " & IIf(Ch_Vigente <> 0, -1, 0)
   Q1 = Q1 & " WHERE idSucursal=" & lSucursal.id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Call ExecSQL(DbMain, Q1)
   
   lSucursal.Codigo = Tx_Codigo
   lSucursal.Descrip = Tx_Descripcion
   
   lRc = vbOK
   Unload Me
End Sub
Friend Function FView(Sucursal As Sucursal_t) As Integer
   lSucursal = Sucursal
   Oper = O_VIEW
   Me.Show vbModal
   
   FView = lRc
   
End Function
Friend Function FEdit(Sucursal As Sucursal_t) As Integer
   lSucursal = Sucursal
   Oper = O_EDIT
   Me.Show vbModal
   
   Sucursal = lSucursal
   FEdit = lRc
End Function

Friend Function FNew(Sucursal As Sucursal_t) As Integer
   Oper = O_NEW
   Me.Show vbModal
   
   Sucursal = lSucursal
   FNew = lRc
End Function

Private Sub Form_Load()
   lRc = vbCancel
   
   If Oper = O_NEW Then
      Caption = "Nueva Sucursal"
      Ch_Vigente = 1
   ElseIf Oper = O_EDIT Then
      Caption = "Modificar Sucursal"
      Call LoadAll
   ElseIf Oper = O_VIEW Then
      Caption = "Ver Sucursal"
      Call LoadAll
      
   End If
   
   Call SetupPriv
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT Codigo, Descripcion, Vigente FROM Sucursales"
   Q1 = Q1 & " WHERE idSucursal=" & lSucursal.id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
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
   
   Q1 = "SELECT idSucursal FROM Sucursales WHERE Codigo='" & Tx_Codigo & "'"
   Q1 = Q1 & " AND idSucursal<>" & lSucursal.id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      MsgBox1 "El código ya existe.", vbExclamation
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

