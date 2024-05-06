VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmHelpImpCartola 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formato Importación Cartolas Bancarias"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   Icon            =   "FrmHelpImpCartola.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton Bt_CopyExcel 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "FrmHelpImpCartola.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   7800
         TabIndex        =   1
         Top             =   180
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1875
      Left            =   480
      TabIndex        =   3
      Top             =   1260
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   3307
      _Version        =   393216
      Rows            =   7
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmHelpImpCartola.frx":0451
      Height          =   615
      Left            =   1200
      TabIndex        =   6
      Top             =   3180
      Width           =   7635
   End
   Begin VB.Label Label2 
      Caption         =   "Columnas o campos del archivo:"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   2475
   End
   Begin VB.Label Label3 
      Caption         =   "NOTA:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   615
   End
End
Attribute VB_Name = "FrmHelpImpCartola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CAMPO = 0
Const C_FORMATO = 1
Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Me.Caption)
End Sub

Private Sub Form_Load()
  

   Call SetUpGrid
   Call LoadGrid

End Sub

Private Sub SetUpGrid()

   Call FGrSetup(Grid)

   Grid.ColWidth(C_CAMPO) = 2400
   Grid.ColWidth(C_FORMATO) = 6200
   
   Grid.TextMatrix(0, C_CAMPO) = "Campo de Información"
   Grid.TextMatrix(0, C_FORMATO) = "Formato"
   
End Sub

Private Sub LoadGrid()
   Dim Row As Integer

   Grid.rows = Grid.FixedRows
   Row = 1
   
   Row = AddGrid("Fecha", "dd/mm/yyyy, por ejemplo: 16/07/2006.", Row)
   Row = AddGrid("Detalle", "Texto de descripción.", Row)
   Row = AddGrid("Nro. Doc.", "Número del documento, sin puntos.", Row)
   Row = AddGrid("Cargo", "Valor numérico.", Row)
   Row = AddGrid("Abono", "Valor numérico.", Row)
   
   Call FGrVRows(Grid)
End Sub

Private Function AddGrid(Campo As String, Formato As String, Row As Integer) As Integer

   Grid.rows = Row + 1
   
   Grid.TextMatrix(Row, C_CAMPO) = Campo
   Grid.TextMatrix(Row, C_FORMATO) = Formato
   
   Row = Row + 1
   AddGrid = Row
   
End Function

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call Bt_CopyExcel_Click
   End If
End Sub
