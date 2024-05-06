VERSION 5.00
Begin VB.Form FrmSelEstRes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de Resultado"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "FrmSelEstRes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2955
      Index           =   0
      Left            =   1500
      TabIndex        =   8
      Top             =   540
      Width           =   4215
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Estado de Resultado FULL (Formato IFRS)"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   4
         Top             =   2340
         Width           =   3495
      End
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Comparativo Periodo Anterior"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   3
         Top             =   1860
         Width           =   2535
      End
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Mensual"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   1
         Top             =   900
         Width           =   1335
      End
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Clasificado"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   0
         Top             =   420
         Width           =   1515
      End
      Begin VB.OptionButton Op_Informes 
         Caption         =   "Comparativo"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   2
         Top             =   1380
         Width           =   1395
      End
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "Seleccionar..."
      Default         =   -1  'True
      Height          =   315
      Left            =   6120
      TabIndex        =   5
      Top             =   660
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6120
      TabIndex        =   6
      Top             =   1020
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   480
      Picture         =   "FrmSelEstRes.frx":000C
      ScaleHeight     =   705
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   660
      Width           =   735
   End
End
Attribute VB_Name = "FrmSelEstRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INFO_CLASIFICADO = 1
Const INFO_MENSUAL = 2
Const INFO_COMPARATIVO = 3
Const INFO_COMPPERANT = 4
Const INFO_ESTRESIFRS = 5

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Sel_Click()
   Dim Frm As FrmBalClasif
   Dim FrmC As FrmBalClasifCompar
   Dim FrmIFRS As FrmLstInformeIFRS

         
   Me.MousePointer = vbHourglass

   If Op_Informes(INFO_CLASIFICADO).Value <> 0 Then
      Set Frm = New FrmBalClasif
      Call Frm.FViewEstResultClasif
      Set Frm = Nothing
      
   ElseIf Op_Informes(INFO_MENSUAL).Value <> 0 Then
      Set Frm = New FrmBalClasif
      Call Frm.FViewEstResultMensual
      Set Frm = Nothing

   ElseIf Op_Informes(INFO_COMPARATIVO).Value <> 0 Then
      Set Frm = New FrmBalClasif
      Call Frm.FViewEstResultComparativo
      Set Frm = Nothing
      
   ElseIf Op_Informes(INFO_COMPPERANT).Value <> 0 Then
   
      If Not gEmpresa.TieneAnoAnt Then
         MsgBox1 "Esta empresa no tiene año anterior en el sistema. No se puede generar el reporte.", vbExclamation
         Me.MousePointer = vbDefault
         Exit Sub
      End If
      
      Set FrmC = New FrmBalClasifCompar
      Call FrmC.FViewEstResultClasif
      Set FrmC = Nothing
      
   ElseIf Op_Informes(INFO_ESTRESIFRS).Value <> 0 Then
      Set FrmIFRS = New FrmLstInformeIFRS
      Call FrmIFRS.FView(IFRS_ESTRES)
      Set FrmIFRS = Nothing
   
   End If
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()

   Op_Informes(INFO_CLASIFICADO) = True
   
End Sub

Private Sub Op_Informes_DblClick(Index As Integer)

   Call PostClick(Bt_Sel)
   
End Sub

