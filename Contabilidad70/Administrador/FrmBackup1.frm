VERSION 5.00
Begin VB.Form FrmBackup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Respaldos"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   Icon            =   "FrmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tx_Dir 
      BackColor       =   &H8000000F&
      Height          =   555
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4740
      Width           =   5235
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   1680
      TabIndex        =   5
      Top             =   5460
      Width           =   5235
      Begin VB.Label Label2 
         Caption         =   $"FrmBackup.frx":000C
         Height          =   735
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   4830
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5580
      TabIndex        =   4
      Top             =   1140
      Width           =   1335
   End
   Begin VB.CommandButton Bt_Backup 
      Caption         =   "Respaldar"
      Height          =   315
      Left            =   5580
      TabIndex        =   3
      Top             =   780
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   780
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   1680
      TabIndex        =   0
      Top             =   1140
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   540
      Picture         =   "FrmBackup.frx":00CF
      Top             =   780
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Carpeta de Destino:"
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   8
      Top             =   4500
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione Carpeta de Destino:"
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   540
      Width           =   2250
   End
End
Attribute VB_Name = "FrmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Backup_Click()
   Dim DbPath As String, fname As String
   Dim Rc As Long
   Dim BckPath As String
   
   If MsgBox1("¡ATENCIÓN!" & vbLf & "Antes de respaldar debe verificar que no hayan usuarios conectados al sistema." & vbLf & "¿Desea continuar?", vbYesNo Or vbDefaultButton2 Or vbQuestion) <> vbYes Then
      Exit Sub
   End If
   
   MousePointer = vbHourglass
   DoEvents
   
   DbPath = DbMain.Name
   BckPath = Tx_Dir
   
   On Error Resume Next
   MkDir BckPath
         
   Call CloseDb(DbMain)
   
   Rc = GenZip(BckPath)

'   If Rc Then
'      Unload Me
'      Exit Sub
'   End If
   
   If OpenDbAdm() = False Then
      End
   End If

   MousePointer = vbDefault

End Sub

Private Sub bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Dir1_Change()
   Tx_Dir = Dir1.Path
End Sub

Private Sub Drive1_Change()

   Dir1.Path = Drive1.Drive

End Sub

Private Sub Form_Load()
   Dim BckPath As String
   
   On Error Resume Next
   BckPath = W.AppPath & "\Backup"

   MkDir BckPath

   Drive1.Drive = Dir1.Path
   Dir1.Path = BckPath
   Tx_Dir = BckPath

End Sub

Private Function GenZip(ByVal ZipPath As String) As Long
   Dim ZipFile As String
   Dim zOpt As ZipOPT_t
   Dim zFiles As ZIPnames_t
   Dim zFnc As ZIPUSERFUNCTIONS_t
   Dim Rc As Long
   Dim SubDir As String, BDir As String
   
   Rc = rInStr(gDbPath, "\")
   SubDir = Mid(gDbPath, Rc + 1)
   BDir = Left(gDbPath, Rc - 1)

   zFiles.zFiles(0) = SubDir
   zFiles.zFiles(1) = SubDir & "\*.mdb"

   zOpt.Date = vbNullString
   zOpt.flevel = Asc(9)  ' Compression Level (0 - 9)
   zOpt.szRootDir = BDir ' gDbPath
   zOpt.fRecurse = 1 ' -r

   ZipFile = ZipPath & "\DbLPContab" & Format(Now, "yymmdd") & ".zip"

   Rc = VBZip32(ZipFile, 2, zFiles, zOpt, zFnc)
   If Rc Then
      MsgBox1 "Error " & Rc & " al generar el archivo:" & vbLf & ZipFile
   Else
      MsgBox1 "Se generó el archivo" & vbLf & ZipFile & vbLf & "que contiene todas las bases de datos.", vbInformation
   End If

   GenZip = Rc
   
End Function
