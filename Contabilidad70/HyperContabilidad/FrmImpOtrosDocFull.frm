VERSION 5.00
Begin VB.Form FrmImpOtrosDocFull 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Otros Documentos Full desde Archivo de Texto"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   540
      Picture         =   "FrmImpOtrosDocFull.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   600
      Width           =   555
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   8460
      TabIndex        =   4
      Top             =   780
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Importar"
      Height          =   315
      Left            =   8460
      TabIndex        =   3
      Top             =   420
      Width           =   1275
   End
   Begin VB.CommandButton Bt_VerFormato 
      Caption         =   "Ver Formato Archivo..."
      Height          =   315
      Left            =   420
      TabIndex        =   2
      Top             =   3000
      Width           =   1995
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Archivo"
      Height          =   1095
      Left            =   420
      TabIndex        =   5
      Top             =   1620
      Width           =   9315
      Begin VB.TextBox Tx_FName 
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   420
         Width           =   7455
      End
      Begin VB.CommandButton Bt_Browse 
         Height          =   495
         Left            =   7800
         Picture         =   "FrmImpOtrosDocFull.frx":058D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmImpOtrosDocFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFName As String

Private Sub Bt_Browse_Click()

   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = ""
   FrmMain.Cm_ComDlg.InitDir = gImportPath
   FrmMain.Cm_ComDlg.Filter = "Archivos de Texto (*.txt)|*.txt"
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
   If ERR = cdlCancel Then
      Exit Sub
   ElseIf ERR Then
      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   ERR.Clear
   
   lFName = FrmMain.Cm_ComDlg.Filename
   
   Tx_FName = lFName
   
   DoEvents
      
End Sub

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Import_Click()
   Dim Rc As Integer
   
   If gEmpresa.FCierre <> 0 Then
      MsgBox1 "Este periodo está cerrado.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   If lFName = "" Then
      MsgBox1 "Debe seleccionar el archivo.", vbExclamation + vbOKOnly
      Exit Sub
   End If

   Rc = MsgBox1("Atención:" & vbNewLine & vbNewLine & "Se importará el archivo:" & vbNewLine & vbNewLine & lFName & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
   If Rc = vbNo Then
      Exit Sub
   End If
   
   Call ImportOtrosDocFull(Me, lFName)

End Sub

Private Sub Bt_VerFormato_Click()
   Dim Frm As FrmFmtImpEnt
   
   Set Frm = New FrmFmtImpEnt
   Call Frm.FViewOtrosDocFull
   Set Frm = Nothing
   
End Sub

