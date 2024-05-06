VERSION 5.00
Begin VB.Form FrmImportF29 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar desde HR-IVA"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "FrmImportF29.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Import 
      Caption         =   "Importar"
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Top             =   960
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   480
      Picture         =   "FrmImportF29.frx":000C
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   600
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Index           =   0
      Left            =   1440
      TabIndex        =   10
      Top             =   480
      Width           =   2595
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Retenciones"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   2
         Top             =   840
         Width           =   2355
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Ventas"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   2235
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Compras"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   2355
      End
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Período"
      Height          =   975
      Left            =   1440
      TabIndex        =   6
      Top             =   1920
      Width           =   4395
      Begin VB.TextBox Tx_Ano 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   420
         Width           =   855
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   8
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   7
         Top             =   480
         Width           =   345
      End
   End
End
Attribute VB_Name = "FrmImportF29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lMsgAdv As Boolean

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Import_Click()
   Dim i As Integer
   
   If lMsgAdv = False Then    'este mensaje se muestra sólo una vez
   
      If MsgBox1("Para realizar la importación desde HR-IVA, nadie debe estar trabajando en esta empresa en HR-IVA." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   
   End If
   
   lMsgAdv = True
   
   Me.MousePointer = vbHourglass
   
   For i = LIB_COMPRAS To LIB_RETEN
   
      If Op_Libros(i).Value = True Then
   
         Call ImportLibF29(ItemData(Cb_Mes), i)
         
         Exit For
         
      End If
   Next i
   
   Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
   Dim MesActual As Integer
   
   lMsgAdv = False
   
   MesActual = GetMesActual()
   
   Call FillMes(Cb_Mes)
   If MesActual > 0 Then
      Cb_Mes.ListIndex = MesActual - 1
   Else
      Cb_Mes.ListIndex = GetUltimoMesConMovs() - 1
   End If
   
   Tx_Ano = gEmpresa.Ano
   
End Sub
