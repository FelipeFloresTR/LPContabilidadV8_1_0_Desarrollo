VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImpActFijoFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Activos Fijos desde Archivo de Texto"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   420
      Picture         =   "FrmImpActFijoFile.frx":0000
      ScaleHeight     =   630
      ScaleWidth      =   825
      TabIndex        =   9
      Top             =   540
      Width           =   885
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   9480
      TabIndex        =   4
      Top             =   1740
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Importar"
      Height          =   315
      Left            =   8100
      TabIndex        =   3
      Top             =   1740
      Width           =   1275
   End
   Begin VB.CommandButton Bt_VerFormato 
      Caption         =   "Ver Formato Archivo..."
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1740
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proceso"
      Height          =   1395
      Left            =   1440
      TabIndex        =   6
      Top             =   2280
      Width           =   9315
      Begin MSComctlLib.ProgressBar Pb_Proceso 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   780
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label Lb_Proceso 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   420
         Width           =   2235
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Archivo"
      Height          =   1095
      Left            =   1500
      TabIndex        =   5
      Top             =   480
      Width           =   9255
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
         Picture         =   "FrmImpActFijoFile.frx":05E8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmImpActFijoFile"
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
   
   If lFName = "" Then
      MsgBox1 "Debe seleccionar el archivo.", vbExclamation + vbOKOnly
      Exit Sub
   End If

   Rc = MsgBox1("Atención:" & vbNewLine & vbNewLine & "Se importará el archivo:" & vbNewLine & vbNewLine & lFName & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
   If Rc = vbNo Then
      Exit Sub
   End If
   
   'revisamos el archivo primero
   If ImportActFijoFile(Me, lFName, True) = True Then
      
      'ahora importamos
      Call ImportActFijoFile(Me, lFName, False)
       
      Tx_FName = ""    'para que el usuario no lo importe de nuevo
      lFName = ""
      
      MsgBox1 "ADVERTENCIA: si vuelve a importar este mismo archivo, se generarán activos fijos duplicados.", vbInformation + vbOKOnly
      
   End If

End Sub

Private Sub Bt_VerFormato_Click()
   Dim Frm As FrmFmtImpEnt
   
   Set Frm = New FrmFmtImpEnt
   Call Frm.FViewActivoFijo
   Set Frm = Nothing
   
End Sub

