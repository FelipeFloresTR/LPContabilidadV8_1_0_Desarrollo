VERSION 5.00
Begin VB.Form FrmSumMov 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Suma de Movimientos"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   Icon            =   "FrmSumMov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   1380
      TabIndex        =   19
      Top             =   420
      Width           =   5715
      Begin VB.TextBox Tx_NMov 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Promedio de Movimientos"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Index           =   1
      Left            =   1380
      TabIndex        =   10
      Top             =   2520
      Width           =   5715
      Begin VB.TextBox Tx_NValHaber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3060
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   1020
         Width           =   555
      End
      Begin VB.TextBox Tx_NValDebe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   1020
         Width           =   555
      End
      Begin VB.TextBox Tx_PromDif 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4020
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox Tx_PromHaber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox Tx_PromDebe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valores <> 0: "
         Height          =   195
         Index           =   2
         Left            =   2100
         TabIndex        =   24
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valores <> 0: "
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Index           =   5
         Left            =   4020
         TabIndex        =   18
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Haber:"
         Height          =   195
         Index           =   4
         Left            =   2100
         TabIndex        =   17
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Debe:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   12
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Suma de Movimientos"
      ForeColor       =   &H00FF0000&
      Height          =   1155
      Index           =   0
      Left            =   1380
      TabIndex        =   7
      Top             =   1200
      Width           =   5715
      Begin VB.TextBox Tx_Debe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox Tx_Haber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2100
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox Tx_Diff 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4020
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Index           =   2
         Left            =   4020
         TabIndex        =   15
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Haber:"
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   14
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Debe:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   3720
         TabIndex        =   8
         Top             =   600
         Width           =   195
      End
   End
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   7320
      TabIndex        =   6
      Top             =   540
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   420
      Picture         =   "FrmSumMov.frx":000C
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "FrmSumMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lNMov As Integer
Dim lNValDebe As Integer
Dim lNValHaber As Integer
Dim lSumDebe As Double
Dim lSumHaber As Double
Dim lPromDebe As Double
Dim lPromHaber As Double

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Tx_NMov = lNMov
   Tx_NValDebe = lNValDebe
   Tx_NValHaber = lNValHaber
   
   Tx_Debe = Format(lSumDebe, NEGNUMFMT)
   Tx_Haber = Format(lSumHaber, NEGNUMFMT)

   Tx_Diff = Format(lSumDebe - lSumHaber, NEGNUMFMT)
   
   Tx_PromDebe = Format(lPromDebe, NEGNUMFMT)
   Tx_PromHaber = Format(lPromHaber, NEGNUMFMT)
   Tx_PromDif = Format(lPromDebe - lPromHaber, NEGNUMFMT)
   
   
End Sub

Public Sub FView(ByVal NMov As Integer, ByVal NValDebe As Integer, ByVal NValHaber As Integer, ByVal SumDebe As Double, ByVal SumHaber As Double, ByVal PromDebe As Double, ByVal PromHaber As Double)
   lNMov = NMov
   lNValDebe = NValDebe
   lNValHaber = NValHaber
   
   lSumDebe = SumDebe
   lSumHaber = SumHaber
   lPromDebe = PromDebe
   lPromHaber = PromHaber
   
   Me.Show vbModal
End Sub

Public Sub FViewSum(Grid As Control, ByVal ColDebe As Integer, ColHaber As Integer, Optional ByVal FirstRow As Integer = -1, Optional ByVal LastRow As Integer = -1)
   Dim SumDebe As Double
   Dim SumHaber As Double
   Dim PromDebe As Double
   Dim PromHaber As Double
   Dim DivDebe As Integer
   Dim DivHaber As Integer
   Dim i As Integer
   Dim NLin As Integer
   Dim AuxRow As Integer
      
   SumDebe = 0
   SumHaber = 0
   PromDebe = 0
   PromHaber = 0
   
   If ColDebe < 0 Then
      If ColHaber < 0 Then
         Exit Sub
      Else
         ColDebe = ColHaber
      End If
   ElseIf ColHaber < 0 Then
      ColHaber = ColDebe
   End If
   
   If FirstRow = -1 Then
      FirstRow = Grid.Row
   End If
   If LastRow = -1 Then
      LastRow = Grid.RowSel
   End If
   
   If LastRow < FirstRow Then  'swap
      AuxRow = FirstRow
      FirstRow = LastRow
      LastRow = AuxRow
   End If
   
   For i = FirstRow To LastRow
      If Grid.RowHeight(i) > 0 Then
         SumDebe = SumDebe + vFmt(Grid.TextMatrix(i, ColDebe))
         SumHaber = SumHaber + vFmt(Grid.TextMatrix(i, ColHaber))
         DivDebe = IIf(vFmt(Grid.TextMatrix(i, ColDebe)) <> 0, DivDebe + 1, DivDebe)
         DivHaber = IIf(vFmt(Grid.TextMatrix(i, ColHaber)) <> 0, DivHaber + 1, DivHaber)
         NLin = NLin + 1
      End If
   Next i
   
   If DivDebe <> 0 Then
      PromDebe = SumDebe / DivDebe
   End If
   
   If DivHaber <> 0 Then
      PromHaber = SumHaber / DivHaber
   End If
   
   Call FView(NLin, DivDebe, DivHaber, SumDebe, SumHaber, PromDebe, PromHaber)

End Sub
