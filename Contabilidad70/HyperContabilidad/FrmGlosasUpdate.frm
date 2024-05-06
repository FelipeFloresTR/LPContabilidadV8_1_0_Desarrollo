VERSION 5.00
Begin VB.Form FrmGlosasUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Glosa"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "FrmGlosasUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7620
      TabIndex        =   3
      Top             =   780
      Width           =   1095
   End
   Begin VB.CommandButton bt_Ok 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   7620
      TabIndex        =   2
      Top             =   420
      Width           =   1095
   End
   Begin VB.TextBox tx_Glosa 
      Height          =   495
      Left            =   1620
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   5835
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   240
      Picture         =   "FrmGlosasUpdate.frx":000C
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Glosa:"
      Height          =   255
      Left            =   1020
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "FrmGlosasUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lGlosa As String
Dim lidGlosa As Long
Dim lRc As Integer
Dim lOper As Integer

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub bt_OK_Click()
   Dim Q1 As String
   Dim Rs As Recordset
      
   If Trim(Tx_Glosa) = "" Then
      MsgBox1 "Debe ingresar una glosa", vbExclamation
      Exit Sub
   End If
   
   If lOper = O_NEW Then
'      Set Rs = DbMain.OpenRecordset("Glosas", dbOpenTable)
'      Rs.AddNew
'
'      lidGlosa = Rs("idGlosa")
'      Rs("Glosa") = ParaSQL(Tx_Glosa)
'
'      Rs.Update
'      Rs.Close
      
      lidGlosa = AdvTbAddNew(DbMain, "Glosas", "idGlosa", "IdEmpresa", gEmpresa.id)
      Q1 = "UPDATE Glosas SET"
      Q1 = Q1 & " Glosa = '" & ParaSQL(Tx_Glosa) & "'"
      Q1 = Q1 & " WHERE idGlosa=" & lidGlosa
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      Call ExecSQL(DbMain, Q1)
            
   Else
      Q1 = "UPDATE Glosas SET Glosa='" & ParaSQL(Tx_Glosa) & "'"
      Q1 = Q1 & " WHERE idGlosa=" & lidGlosa
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      Call ExecSQL(DbMain, Q1)
      
   End If
   
   lGlosa = Tx_Glosa
   lRc = vbOK
   Unload Me
   
End Sub
Public Function FEdit(Glosa As String, idGlosa As Long, Oper As Integer) As Integer
   lOper = Oper
   lidGlosa = idGlosa
   lGlosa = Glosa
   Me.Show vbModal
   
   FEdit = lRc
   Glosa = lGlosa
   idGlosa = lidGlosa
   
End Function

Private Sub Form_Load()
   lRc = vbCancel
   Tx_Glosa = lGlosa
   
End Sub
