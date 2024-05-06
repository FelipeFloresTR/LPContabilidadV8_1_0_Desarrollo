VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImpComp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Comprobantes desde Archivo de Texto"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   540
      Picture         =   "FrmImpComp.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   12
      Top             =   600
      Width           =   555
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   8460
      TabIndex        =   5
      Top             =   780
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Importar"
      Height          =   315
      Left            =   8460
      TabIndex        =   4
      Top             =   420
      Width           =   1275
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Período"
      Height          =   975
      Left            =   2160
      TabIndex        =   8
      Top             =   360
      Width           =   4395
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox Tx_Ano 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   11
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   10
         Top             =   480
         Width           =   330
      End
   End
   Begin VB.CommandButton Bt_VerFormato 
      Caption         =   "Ver Formato Archivo..."
      Height          =   315
      Left            =   420
      TabIndex        =   3
      Top             =   3000
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proceso"
      Height          =   1395
      Left            =   360
      TabIndex        =   7
      Top             =   3780
      Width           =   9315
      Begin MSComctlLib.ProgressBar Pb_Proceso 
         Height          =   255
         Left            =   240
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   420
         Width           =   2235
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Archivo"
      Height          =   1095
      Left            =   420
      TabIndex        =   6
      Top             =   1620
      Width           =   9315
      Begin VB.TextBox Tx_FName 
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   420
         Width           =   7455
      End
      Begin VB.CommandButton Bt_Browse 
         Height          =   495
         Left            =   7800
         Picture         =   "FrmImpComp.frx":058D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmImpComp"
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
   
   If GetEstadoMes(CbItemData(Cb_Mes)) <> EM_ABIERTO Then
      MsgBox1 "El mes seleccionado no está abierto.", vbExclamation
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
   
   'revisamos el archivo primero
   If ImportComprobantes(Me, lFName, Val(Tx_Ano), CbItemData(Cb_Mes), True) = True Then
      
      'ahora importamos
      Call ImportComprobantes(Me, lFName, Val(Tx_Ano), CbItemData(Cb_Mes), False)
       
      Tx_FName = ""    'para que el usuario nbo lo importe de nuevo
      lFName = ""
      
      MsgBox1 "ADVERTENCIA: si vuelve a importar este mismo archivo, se generarán comprobantes duplicados," & vbCrLf & "dado que el sistema asigna automáticamente el correlativo.", vbInformation + vbOKOnly
      
   End If

End Sub

Private Sub Bt_VerFormato_Click()
   Dim Frm As FrmFmtImpEnt
   
   Set Frm = New FrmFmtImpEnt
   Call Frm.FViewComprobantes
   Set Frm = Nothing
   
End Sub

Private Sub Form_Load()
   Dim MesActual As Integer

   MesActual = GetMesActual()
   
   Call FillMes(Cb_Mes)
   If MesActual > 0 Then
      Cb_Mes.ListIndex = MesActual - 1
   Else
      Cb_Mes.ListIndex = GetUltimoMesConMovs() - 1
   End If
   
   Tx_Ano = gEmpresa.Ano

End Sub
