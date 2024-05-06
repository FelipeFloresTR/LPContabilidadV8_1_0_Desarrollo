VERSION 5.00
Begin VB.Form FrmSelRuta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar la Ruta del Archivo"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   540
      Picture         =   "FrmSelRuta.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   600
      Width           =   555
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   9360
      TabIndex        =   3
      Top             =   1740
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "Seleccionar"
      Height          =   315
      Left            =   7800
      TabIndex        =   2
      Top             =   1740
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Archivo"
      Height          =   1095
      Left            =   1260
      TabIndex        =   4
      Top             =   480
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
         Picture         =   "FrmSelRuta.frx":058D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmSelRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lSelFile As String
Dim lWinTitle As String
Dim lFileFilter As String
Dim lRc As Integer
Dim lExpectedFName As String

Public Function FSelFile(ByVal WinTitle As String, ByVal FileFilter As String, ExpectedFName As String, SelFile As String) As Integer

   lWinTitle = WinTitle
   lFileFilter = FileFilter
   lExpectedFName = ExpectedFName
   
   Me.Show vbModal
   
   If lRc = vbOK Then
      SelFile = lSelFile
   End If
   
   FSelFile = lRc
End Function


Private Sub Bt_Browse_Click()

   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = ""
   FrmMain.Cm_ComDlg.InitDir = gAppPath
   If lFileFilter = "" Then
      FrmMain.Cm_ComDlg.Filter = "Archivos CSV (*.csv)|*.csv"
   Else
      FrmMain.Cm_ComDlg.Filter = lFileFilter
   End If
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
   If Err = cdlCancel Then
      Exit Sub
   ElseIf Err Then
      MsgBox1 "Error " & Err & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   Err.Clear
   
   If lExpectedFName <> "" Then
      If FrmMain.Cm_ComDlg.FileTitle <> lExpectedFName Then
         MsgBox1 "Nombre de archivo inválido." & vbCrLf & vbCrLf & "Nombre esperado: " & lExpectedFName, vbExclamation
         Exit Sub
      End If
   End If
   
   lSelFile = FrmMain.Cm_ComDlg.Filename
   
   Tx_FName = lSelFile
   
   DoEvents
      
End Sub

Private Sub Bt_Cancelar_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_Sel_Click()
   Dim Rc As Integer
   

   If lSelFile = "" Then
      MsgBox1 "Debe seleccionar el archivo.", vbExclamation + vbOKOnly
      Exit Sub
   End If

   Rc = MsgBox1("ATENCIÓN" & vbNewLine & vbNewLine & "Se utilizará el archivo:" & vbNewLine & vbNewLine & lSelFile & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
   If Rc = vbNo Then
      Exit Sub
   End If
   
   lRc = vbOK
   
   Unload Me
         
End Sub

Private Sub Form_Load()

   Me.Caption = lWinTitle
   
End Sub
