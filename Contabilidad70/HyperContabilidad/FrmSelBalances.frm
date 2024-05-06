VERSION 5.00
Begin VB.Form FrmSelBalances 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balances (bajo norma antigua)"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "FrmSelBalances.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_View 
      Caption         =   "Seleccionar..."
      Default         =   -1  'True
      Height          =   315
      Left            =   6120
      TabIndex        =   5
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6120
      TabIndex        =   6
      Top             =   960
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Index           =   0
      Left            =   1560
      TabIndex        =   7
      Top             =   540
      Width           =   4275
      Begin VB.OptionButton Op_Balances 
         Caption         =   "Balance Clasificado Ejecutivo"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   4
         Top             =   2580
         Width           =   3795
      End
      Begin VB.OptionButton Op_Balances 
         Caption         =   "Balance Clasificado Comparativo Periodo Anterior"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   3
         Top             =   2040
         Width           =   3795
      End
      Begin VB.OptionButton Op_Balances 
         Caption         =   "Balance de Comprobación y Saldos"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   420
         Width           =   2895
      End
      Begin VB.OptionButton Op_Balances 
         Caption         =   "Balance General 8 Columnas"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   960
         Width           =   2835
      End
      Begin VB.OptionButton Op_Balances 
         Caption         =   "Balance Clasificado"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Top             =   1500
         Width           =   1755
      End
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   480
      Picture         =   "FrmSelBalances.frx":000C
      Top             =   660
      Width           =   750
   End
End
Attribute VB_Name = "FrmSelBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const BAL_COMPROB = 0
Const BAL_TRIBUTARIO = 1
Const BAL_CLASIF = 2
Const BAL_COMPANT = 3
Const BAL_EJECUTIVO = 4
Const BAL_ESTFINIFRS = 5

Dim lIdx As Integer

Private Sub bt_Cerrar_Click()
    Unload Me
End Sub

Private Sub bt_View_Click()
   Dim FrmBalComp As FrmBalComprobacion
   Dim FrmBalTrib As FrmBalTributario
   Dim FrmBalCla As FrmBalClasif
   Dim FrmBalClaComp As FrmBalClasifCompar
   Dim FrmBalEjec As FrmBalClasifEjec
   Dim FrmInv As FrmLibInvBal
   Dim StList As String
   Dim FrmIFRS As FrmLstInformeIFRS
   Dim FrmBalEjIFRS As FrmBalEjecIFRS

   
   Me.MousePointer = vbHourglass
   
   Select Case lIdx
      Case BAL_COMPROB
         Set FrmBalComp = New FrmBalComprobacion
         FrmBalComp.FView (0)
         Set FrmBalComp = Nothing
         
      Case BAL_TRIBUTARIO
         Set FrmBalTrib = New FrmBalTributario
         FrmBalTrib.FView (0)
         Set FrmBalTrib = Nothing
         
      Case BAL_CLASIF
         Set FrmBalCla = New FrmBalClasif
         FrmBalCla.FViewBalClasif
         Set FrmBalCla = Nothing
         
      Case BAL_COMPANT
         
         If Not gEmpresa.TieneAnoAnt Then
            MsgBox1 "Esta empresa no tiene año anterior en el sistema. No se puede generar el reporte.", vbExclamation
            Me.MousePointer = vbDefault
            Exit Sub
         End If

         Set FrmBalClaComp = New FrmBalClasifCompar
         FrmBalClaComp.FViewBalClasif
         Set FrmBalCla = Nothing
         
      Case BAL_EJECUTIVO
'         If gPlanCuentas = "IFRS" Then
'            Set FrmBalEjIFRS = New FrmBalEjecIFRS
'            FrmBalEjIFRS.FView
'            Set FrmBalEjecIFRS = Nothing
'         Else
            Set FrmBalEjec = New FrmBalClasifEjec
            FrmBalEjec.FViewBalClasif
            Set FrmBalEjec = Nothing
'         End If
         
      Case BAL_ESTFINIFRS
         Set FrmIFRS = New FrmLstInformeIFRS
         Call FrmIFRS.FView(IFRS_ESTFIN)
         Set FrmIFRS = Nothing
   
         
   End Select

   Me.MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()
   Op_Balances(BAL_COMPROB).Value = True
End Sub
Private Sub Op_Balances_Click(Index As Integer)
     lIdx = Index
End Sub

Private Sub Op_Balances_DblClick(Index As Integer)
   lIdx = Index
   Call PostClick(Bt_View)
End Sub
