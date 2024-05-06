VERSION 5.00
Begin VB.Form FrmSelInfIFRS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes IFRS"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "FrmSelInfIFRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Estado de Resultados por Función"
      Height          =   975
      Left            =   1440
      TabIndex        =   9
      Top             =   2580
      Width           =   4035
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Clásico"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   3
         Top             =   420
         Width           =   3555
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   420
      Picture         =   "FrmSelInfIFRS.frx":000C
      ScaleHeight     =   330
      ScaleWidth      =   525
      TabIndex        =   7
      Top             =   1260
      Width           =   525
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5940
      TabIndex        =   5
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "Seleccionar..."
      Height          =   315
      Left            =   5940
      TabIndex        =   4
      Top             =   540
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado de Situación Financiera Clasificado"
      Height          =   1935
      Index           =   0
      Left            =   1440
      TabIndex        =   6
      Top             =   420
      Width           =   4035
      Begin VB.OptionButton Op_Informes 
         Caption         =   "8 Columnas"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   2
         Top             =   1380
         Width           =   3555
      End
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Ejecutivo"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   1
         Top             =   900
         Width           =   3555
      End
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Clásico"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   0
         Top             =   420
         Width           =   3615
      End
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   360
      Picture         =   "FrmSelInfIFRS.frx":04F6
      Top             =   480
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmSelInfIFRS.frx":0B96
      ForeColor       =   &H00FF0000&
      Height          =   435
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   3840
      Width           =   5775
   End
End
Attribute VB_Name = "FrmSelInfIFRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Sel_Click()
   Dim Frm As Form
   
   
   Me.MousePointer = vbHourglass
   If Op_Informes(IFRS_ESTFIN).Value <> 0 Then
      Set Frm = New FrmLstInformeIFRS
      Call Frm.FView(IFRS_ESTFIN)
   
   ElseIf Op_Informes(IFRS_BALEJEC).Value <> 0 Then
      Set Frm = New FrmBalEjecIFRS
      Call Frm.FView
      
   ElseIf Op_Informes(IFRS_BAL8COL).Value <> 0 Then
      Set Frm = New FrmBalTributarioIFRS
      Call Frm.FView

   ElseIf Op_Informes(IFRS_ESTRES).Value <> 0 Then
      Set Frm = New FrmLstInformeIFRS
      Call Frm.FView(IFRS_ESTRES)
   
   End If
   
   Set Frm = Nothing
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()

'   Op_Informes(IFRS_ESTFIN).Caption = gInformeIFRS(IFRS_ESTFIN)
'   Op_Informes(IFRS_ESTRES).Caption = gInformeIFRS(IFRS_ESTRES)
'   Op_Informes(IFRS_BALEJEC).Caption = gInformeIFRS(IFRS_BALEJEC)
'   Op_Informes(IFRS_BAL8COL).Caption = gInformeIFRS(IFRS_BAL8COL)

   Op_Informes(IFRS_ESTFIN) = True
   
End Sub

Private Sub Op_Informes_Click(Index As Integer)

   If Index = IFRS_ESTRES Then
      Op_Informes(IFRS_ESTFIN).Value = 0
      Op_Informes(IFRS_BALEJEC).Value = 0
      Op_Informes(IFRS_BAL8COL).Value = 0
   Else
      Op_Informes(IFRS_ESTRES).Value = 0
   End If

End Sub

Private Sub Op_Informes_DblClick(Index As Integer)

   Call PostClick(Bt_Sel)
   
End Sub
