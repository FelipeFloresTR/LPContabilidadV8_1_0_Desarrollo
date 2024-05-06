VERSION 5.00
Begin VB.Form FrmDemo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro Demo"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6360
   Icon            =   "FrmDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2775
      Index           =   1
      Left            =   3300
      TabIndex        =   6
      Top             =   2280
      Width           =   2775
      Begin VB.CommandButton Bt_Regist 
         Caption         =   "Inscribir..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   300
         TabIndex        =   3
         Top             =   2040
         Width           =   2235
      End
      Begin VB.CheckBox Ck_Lic 
         Caption         =   "Acepto las Condiciones de la Licencia de Uso"
         Height          =   495
         Left            =   420
         TabIndex        =   2
         Top             =   1380
         Width           =   1995
      End
      Begin VB.CommandButton Bt_Lic 
         Caption         =   "Leer Licencia de Uso y Garantía Limitada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   300
         TabIndex        =   1
         Top             =   300
         Width           =   2235
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
      Begin VB.CommandButton Bt_Demo 
         Caption         =   "DEMO"
         Default         =   -1  'True
         Height          =   975
         Left            =   360
         TabIndex        =   0
         Top             =   1500
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Presione el botón DEMO para probar la aplicación en modo demostración."
         Height          =   1095
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1875
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "FrmDemo.frx":000C
      Top             =   240
      Width           =   5835
   End
End
Attribute VB_Name = "FrmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Demo_Click()

   gRc.Rc = vbYes

   Unload Me
   
End Sub

Private Sub Bt_Lic_Click()
   Dim Rc As Long
   Dim Buf As String
   
   MousePointer = vbHourglass
   DoEvents
   
   Buf = gAppPath & "\Licencia1.wri"
   Rc = ExistFile(Buf)
      
   If Rc = 0 Then
      Buf = gAppPath & "\Licencia1.rtf"
      Rc = ExistFile(Buf)
   End If
   
   If Rc = 0 Then
      Buf = gAppPath & "\Licencia1.pdf"
      Rc = ExistFile(Buf)
   End If
      
   If Rc = 0 Then
      Buf = gAppPath & "\Licencia1.htm"
      Rc = ExistFile(Buf)
   End If

   If Rc = 0 Then
      MsgBox1 "No se encontró el archivo que contiene la licencia de uso de la aplicación, por favor contáctese con su proveedor para conseguirlo.", vbExclamation
   Else

      Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
      If Rc < 32 Then
         MsgBox1 "Error " & Rc & " al abrir el archivo '" & Buf & "' que contiene la licencia de uso y garantía de la aplicación." & vbLf & "Trate de abrir este archivo con otro programa.", vbExclamation
      End If
   End If

   MousePointer = vbDefault
   
End Sub

Private Sub Bt_Regist_Click()
   gRc.Rc = vbOK
   
   Unload Me
End Sub

Private Sub Ck_Lic_Click()
   Bt_Regist.Enabled = Ck_Lic.Value
End Sub

Private Sub Form_Load()

   If Trim(gAppCode.Title) = "" Then
      gAppCode.Title = App.Title
   End If

   Me.Caption = "Inscripción de " & gAppCode.Title
      
   gRc.Rc = vbCancel

   Call ToTaskBar(Me)

End Sub
