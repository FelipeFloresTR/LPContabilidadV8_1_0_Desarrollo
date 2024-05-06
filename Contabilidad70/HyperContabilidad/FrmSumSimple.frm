VERSION 5.00
Begin VB.Form FrmSumSimple 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Suma de Movimientos"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   Icon            =   "FrmSumSimple.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   6120
      TabIndex        =   7
      Top             =   600
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Suma y Promedio"
      ForeColor       =   &H00FF0000&
      Height          =   1515
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   4635
      Begin VB.TextBox Tx_Prom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   900
         Width           =   1515
      End
      Begin VB.TextBox Tx_NVal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   555
      End
      Begin VB.TextBox Tx_Suma 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Promedio:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   10
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valores <> 0: "
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   9
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Suma:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   420
      Width           =   4635
      Begin VB.TextBox Tx_NMov 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad de Movimientos: "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   420
      Picture         =   "FrmSumSimple.frx":000C
      Top             =   540
      Width           =   615
   End
End
Attribute VB_Name = "FrmSumSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lNMov As Integer
Dim lNVal As Integer
Dim lSum As Double
Dim lProm As Double

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Tx_NMov = lNMov
   Tx_NVal = lNVal
   
   Tx_Suma = Format(lSum, NEGNUMFMT)
   
   Tx_Prom = Format(lProm, NEGNUMFMT)
   
End Sub

Public Sub FView(ByVal NMov As Integer, ByVal NVal As Integer, ByVal Suma As Double, ByVal Prom As Double)
   lNMov = NMov
   lNVal = NVal
   
   lSum = Suma
   lProm = Prom
   
   Me.Show vbModal
End Sub

Public Sub FViewSum(Grid As Control, Optional ByVal FirstRow As Integer = -1, Optional ByVal LastRow As Integer = -1, Optional ByVal FirstCol As Integer = -1, Optional ByVal LASTCOL As Integer = -1)
   Dim Suma As Double
   Dim Prom As Double
   Dim ResDiv As Integer
   Dim Row As Integer
   Dim Col As Integer
   Dim NLin As Integer
   Dim AuxRow As Integer
   Dim AuxCol As Integer
      
   Suma = 0
   Prom = 0
   
   If FirstRow = -1 Then
      FirstRow = Grid.Row
   End If
   If LastRow = -1 Then
      LastRow = Grid.RowSel
   End If
       
   If FirstCol = -1 Then
      FirstCol = Grid.Col
   End If
   If LASTCOL = -1 Then
      LASTCOL = Grid.ColSel
   End If
   
   If LastRow < FirstRow Then  'swap
      AuxRow = FirstRow
      FirstRow = LastRow
      LastRow = AuxRow
   End If
   
   If LASTCOL < FirstCol Then  'swap
      AuxCol = FirstCol
      FirstCol = LASTCOL
      LASTCOL = AuxCol
   End If
   
   For Row = FirstRow To LastRow
      For Col = FirstCol To LASTCOL
         If Grid.RowHeight(Row) > 0 Then
         '2814013
          If Grid.ColWidth(Col) > 0 Then
            Suma = Suma + vFmt(Grid.TextMatrix(Row, Col))
            ResDiv = IIf(vFmt(Grid.TextMatrix(Row, Col)) <> 0, ResDiv + 1, ResDiv)
            NLin = NLin + 1
          End If
          ' fin 2814013
         End If
      Next Col
   Next Row
   
   If ResDiv <> 0 Then
      Prom = Suma / ResDiv
   End If
   
   Call FView(NLin, ResDiv, Suma, Prom)

End Sub

