VERSION 5.00
Begin VB.Form FrmMsgConBreak 
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   8220
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   1800
      Width           =   1275
   End
   Begin VB.CheckBox Ch_NoMostrarMas 
      Caption         =   "No mostrar nuevamente este mensaje"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1860
      Width           =   3015
   End
   Begin VB.TextBox Tx_Msg 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FrmMsgConBreak.frx":0000
      Top             =   360
      Width           =   6855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   300
      Picture         =   "FrmMsgConBreak.frx":0006
      ScaleHeight     =   525
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   420
      Width           =   510
   End
End
Attribute VB_Name = "FrmMsgConBreak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lMsg As String
Dim lOptIni As String

Private Sub Bt_OK_Click()

   If Ch_NoMostrarMas <> 0 Then
      Call SetIniString(gIniFile, "Opciones", lOptIni, "1")
   End If
   
   Unload Me
      
End Sub

Private Sub Form_Load()

   Me.Caption = App.Title
   Tx_Msg = lMsg
   
End Sub

Public Function FView(ByVal Msg As String, ByVal OptIni As String)
   Dim NoDispMsg As Integer
   Dim Buf As String

   lMsg = Msg
   lOptIni = OptIni
   
   Buf = GetIniString(gIniFile, "Opciones", lOptIni, "0")
   NoDispMsg = Val(Buf)
   
   If NoDispMsg = 0 Then
      Me.Show vbModal
   Else
      Unload Me
   End If
  
End Function


