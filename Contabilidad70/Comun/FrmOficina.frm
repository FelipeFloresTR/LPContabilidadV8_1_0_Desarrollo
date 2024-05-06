VERSION 5.00
Begin VB.Form FrmOficina 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos Oficina Contabilidad"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   Icon            =   "FrmOficina.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   1440
      TabIndex        =   4
      Top             =   360
      Width           =   4695
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   1020
         MaxLength       =   30
         TabIndex        =   0
         Top             =   420
         Width           =   3255
      End
      Begin VB.TextBox Tx_RUT 
         Height          =   315
         Left            =   1020
         MaxLength       =   16
         TabIndex        =   1
         Top             =   840
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   900
         Width           =   390
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6480
      TabIndex        =   3
      Top             =   840
      Width           =   1155
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   6480
      TabIndex        =   2
      Top             =   420
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   420
      Picture         =   "FrmOficina.frx":000C
      Top             =   420
      Width           =   750
   End
End
Attribute VB_Name = "FrmOficina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Rut As String

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub bt_OK_Click()
   Dim Q1 As String, Rc As Long

   If Trim(Tx_Nombre) = "" Then
      MsgBox1 "Ingrese el nombre de la oficina.", vbExclamation
      Exit Sub
   End If

   If ValidRut(Tx_RUT) = False Then
      MsgBox1 "RUT Inválido.", vbExclamation
      Exit Sub
   End If

   gOficina.Rut = vFmtRut(Tx_RUT) & "-" & DV_Rut(vFmtRut(Tx_RUT))
   gOficina.Nombre = Trim(Tx_Nombre)
   
   gAppCode.Rut = gOficina.Rut

   Q1 = "UPDATE Param SET Valor='" & ParaSQL(gOficina.Nombre) & "' WHERE Tipo='OFICINA' AND Codigo=" & TOF_NOMBRE
   Rc = ExecSQL(DbMain, Q1)

   Q1 = "UPDATE Param SET Valor='" & ParaSQL(gOficina.Rut) & "' WHERE Tipo='OFICINA' AND Codigo=" & TOF_RUT
   Rc = ExecSQL(DbMain, Q1)

   Unload Me
   
End Sub

Private Sub Form_Load()

   Call LoadAll

End Sub

Private Sub LoadAll()
   Dim Q1 As String, Rc As Long, Rs As Recordset

   Q1 = "SELECT Codigo, Valor FROM Param WHERE Tipo='OFICINA'"
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF Then
     Call CloseRs(Rs)
     
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('OFICINA', " & TOF_NOMBRE & ", ' ' )"
      Rc = ExecSQL(DbMain, Q1)
   
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('OFICINA', " & TOF_RUT & ", '" & gAppCode.Rut & "' )"
      Rc = ExecSQL(DbMain, Q1)
      
   Else
   
      Do Until Rs.EOF
      
         Select Case vFld(Rs("Codigo"))
         
            Case TOF_NOMBRE:
               Tx_Nombre = vFld(Rs("Valor"))
            Case TOF_RUT:
               Rut = vFld(Rs("Valor"))
               Rc = vFmtRut(Rut)
               If Rc > 0 Then
                  Tx_RUT = FmtRut(vFmtRut(Rut))
               Else
                  Tx_RUT = ""
               End If
         End Select
   
         Rs.MoveNext
      Loop
      Call CloseRs(Rs)
   End If
   
   If Tx_RUT = "" And gAppCode.Rut <> "" Then
      Tx_RUT = gAppCode.Rut
   End If
   
End Sub

Private Sub Tx_Rut_GotFocus()
   Tx_RUT = RutGotFocus(Tx_RUT)
End Sub

Private Sub Tx_Rut_LostFocus()
    Tx_RUT = RutLostFocus(Tx_RUT)
End Sub
