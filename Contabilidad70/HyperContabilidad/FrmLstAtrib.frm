VERSION 5.00
Begin VB.Form FrmLstAtrib 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista de Claves para Atributos"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox Ls_Atrib 
      BackColor       =   &H80000018&
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Atributo"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clave"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "FrmLstAtrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim i As Integer

   For i = 1 To MAX_ATRIB
      Ls_Atrib.AddItem gAtribCuentas(i).NombreCorto & vbTab & gAtribCuentas(i).Nombre
      Ls_Atrib.ItemData(Ls_Atrib.NewIndex) = i
   Next i
   
   Ls_Atrib.ListIndex = -1

End Sub
