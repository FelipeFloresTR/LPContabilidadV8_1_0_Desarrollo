VERSION 5.00
Begin VB.Form FrmConfigActFijo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar Activo Fijo"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8040
      TabIndex        =   4
      Top             =   960
      Width           =   1155
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   8040
      TabIndex        =   3
      Top             =   540
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   " Reporte de Control de Activo Fijo Financiero"
      Height          =   1095
      Left            =   1680
      TabIndex        =   1
      Top             =   420
      Width           =   5955
      Begin VB.CheckBox Ch_AFMesCompleto 
         Caption         =   "Considerar Mes Completo indistintamente la fecha de inicio de utilización"
         Height          =   435
         Left            =   300
         TabIndex        =   2
         Top             =   360
         Width           =   5595
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   480
      Picture         =   "FrmConfigActFijo.frx":0000
      ScaleHeight     =   630
      ScaleWidth      =   825
      TabIndex        =   0
      Top             =   540
      Width           =   885
   End
End
Attribute VB_Name = "FrmConfigActFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_OK_Click()
   Call SaveAll
   Unload Me
End Sub

Private Sub Form_Load()
   
   Call EnableForm(Me, gEmpresa.FCierre = 0)
      
   Call LoadAll
   
   Call SetupPriv

End Sub

Private Function SetupPriv()
   
   If Not ChkPriv(PRV_CFG_EMP) Then
      Call EnableForm(Me, False)
   End If
   
End Function

Private Sub LoadAll()

   Ch_AFMesCompleto = Abs(gAFMesCompleto)

End Sub
Private Sub SaveAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim AFMesCompleto As Integer
   
   Set Rs = OpenRs(DbMain, "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'AFMESCOMPT'")
   
   If Ch_AFMesCompleto <> Abs(gAFMesCompleto = True) Then 'cambió

      AFMesCompleto = False
      If Ch_AFMesCompleto <> 0 Then
         AFMesCompleto = True
      End If

      If Rs.EOF = False Then
         'actualizamos
         Call ExecSQL(DbMain, "UPDATE ParamEmpresa SET Codigo = 0, Valor = '" & AFMesCompleto & "' WHERE Tipo = 'AFMESCOMPT'")
      Else
         'insertamos
         Call ExecSQL(DbMain, "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('AFMESCOMPT', 0, '" & AFMesCompleto & "')")
      End If

      gAFMesCompleto = AFMesCompleto

   End If
   
   Call CloseRs(Rs)

End Sub
