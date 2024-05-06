VERSION 5.00
Begin VB.Form FrmANeg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Area de negocio"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "FrmANeg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   1320
      TabIndex        =   5
      Top             =   480
      Width           =   6195
      Begin VB.CheckBox Ch_Vigente 
         Caption         =   "Vigente"
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Top             =   1440
         Width           =   2235
      End
      Begin VB.TextBox Tx_Codigo 
         Height          =   315
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   0
         Top             =   540
         Width           =   2115
      End
      Begin VB.TextBox Tx_Descripcion 
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   900
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   540
      End
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   7920
      TabIndex        =   3
      Top             =   540
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7920
      TabIndex        =   4
      Top             =   960
      Width           =   1155
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   360
      Picture         =   "FrmANeg.frx":000C
      Top             =   600
      Width           =   750
   End
End
Attribute VB_Name = "FrmANeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lAreaNeg As AreaNeg_t
Dim Oper As Integer
Dim lRc As Integer

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
'      Set Rs = DbMain.OpenRecordset("AreaNegocio", dbOpenTable)
'      Rs.AddNew
'
'      lAreaNeg.id = Rs("idAreaNegocio")
'      Rs("Codigo") = ParaSQL(Tx_Codigo)
'      Rs("Descripcion") = ParaSQL(Tx_Descripcion)
'
'      Rs.Update
'      Rs.Close

      lAreaNeg.Id = AdvTbAddNew(DbMain, "AreaNegocio", "idAreaNegocio", "IdEmpresa", gEmpresa.Id)
        
   End If
   
   Q1 = "UPDATE AreaNegocio SET Codigo='" & ParaSQL(Tx_Codigo) & "'"
   Q1 = Q1 & ", Descripcion='" & ParaSQL(Tx_Descripcion) & "'"
   Q1 = Q1 & ", Vigente = " & IIf(Ch_Vigente <> 0, -1, 0)
   Q1 = Q1 & " WHERE idAreaNegocio=" & lAreaNeg.Id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.Id
   Call ExecSQL(DbMain, Q1)
      
   
   lAreaNeg.Codigo = Tx_Codigo
   lAreaNeg.Descrip = Tx_Descripcion
   
   lRc = vbOK
   Unload Me
End Sub
Friend Function FView(AreaNeg As AreaNeg_t) As Integer
   lAreaNeg = AreaNeg
   Oper = O_VIEW
   Me.Show vbModal
   
   FView = lRc
   AreaNeg = lAreaNeg
   
End Function
Friend Function FEdit(AreaNeg As AreaNeg_t) As Integer
   lAreaNeg = AreaNeg
   Oper = O_EDIT
   Me.Show vbModal
   
   AreaNeg = lAreaNeg
   FEdit = lRc
End Function

Friend Function FNew(AreaNeg As AreaNeg_t) As Integer
   Oper = O_NEW
   Me.Show vbModal
      
   AreaNeg = lAreaNeg
   FNew = lRc
End Function

Private Sub Form_Load()
   lRc = vbCancel
   
   If Oper = O_NEW Then
      Caption = "Nueva Area de negocio"
      Ch_Vigente = 1
   ElseIf Oper = O_EDIT Then
      Caption = "Modificar Area de negocio"
      Call LoadAll
   ElseIf Oper = O_VIEW Then
      Caption = "Ver Area de negocio"
      Call LoadAll
   End If
   
   
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   
   Call SetupPriv
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   
   Q1 = "SELECT Codigo, Descripcion, Vigente FROM AreaNegocio"
   Q1 = Q1 & " WHERE idAreaNegocio=" & lAreaNeg.Id
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.Id
   
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
   
   Q1 = "SELECT idAreaNegocio FROM AreaNegocio WHERE Codigo='" & Tx_Codigo & "'"
   Q1 = Q1 & " AND idAreaNegocio<>" & lAreaNeg.Id
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

