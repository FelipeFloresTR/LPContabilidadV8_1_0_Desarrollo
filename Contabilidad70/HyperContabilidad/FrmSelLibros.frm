VERSION 5.00
Begin VB.Form FrmSelLibros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libros"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "FrmSelLibros.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   420
      Picture         =   "FrmSelLibros.frx":000C
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   17
      Top             =   480
      Width           =   585
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Periodo"
      Height          =   975
      Left            =   1440
      TabIndex        =   14
      Top             =   5460
      Width           =   4395
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   420
         Width           =   1335
      End
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   16
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   15
         Top             =   480
         Width           =   330
      End
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6360
      TabIndex        =   12
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton Bt_View 
      Caption         =   "Seleccionar..."
      Height          =   315
      Left            =   6360
      TabIndex        =   11
      Top             =   480
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   4770
      Index           =   0
      Left            =   1440
      TabIndex        =   13
      Top             =   420
      Width           =   4455
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro Especial de Compras"
         Height          =   195
         Index           =   8
         Left            =   720
         TabIndex        =   7
         Top             =   3660
         Width           =   2430
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Ventas con Boletas"
         Height          =   195
         Index           =   7
         Left            =   720
         TabIndex        =   6
         Top             =   3210
         Width           =   2430
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Ventas Exportación"
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   5
         Top             =   2775
         Width           =   2340
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Inventario y Balance"
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   2
         Top             =   1320
         Width           =   2475
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Retenciones"
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   8
         Top             =   4140
         Width           =   1935
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Ventas"
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   4
         Top             =   2340
         Width           =   1755
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Compras"
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   3
         Top             =   1920
         Width           =   1755
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro Mayor"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   1
         Top             =   900
         Width           =   1215
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro Diario"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmSelLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const OP_DIARIO = 0
Const OP_MAYOR = 1
Const OP_INVENTARIOBAL = 2
Const OP_COMPRAS = 3
Const OP_VENTAS = 4
Const OP_RETENCIONES = 5
Const OP_EXPORTA = 6
Const OP_BOLETA = 7
Const OP_FCV = 8     'facturas de compra del libro de ventas

Dim lIdx As Integer

Dim lPeriodo As Boolean

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub bt_View_Click()
   Dim FrmLDiario As FrmLibDiario
   Dim FrmLMayor As FrmLibMayor
   Dim FrmLCompVta As FrmCompraVenta
   Dim FrmLReten As FrmLibRetenciones
   Dim FrmInv As FrmLibInvBal
   Dim Mes As Integer
   Dim Ano As Integer
   Dim IdDoc As Long
   Dim Wh As String
   
   If lPeriodo Then
      Mes = ItemData(Cb_Mes)
      Ano = Val(Cb_Ano)
   End If
         
   Me.MousePointer = vbHourglass
   
   Select Case lIdx
      Case OP_DIARIO
         Set FrmLDiario = New FrmLibDiario
         FrmLDiario.FView (Mes)
         Set FrmLDiario = Nothing
         
      Case OP_MAYOR
         Set FrmLMayor = New FrmLibMayor
         FrmLMayor.FView (Mes)
         Set FrmLMayor = Nothing
         
      Case OP_VENTAS
         Set FrmLCompVta = New FrmCompraVenta
         Wh = " AND (DocImpExp = 0 AND DocBoletas = 0 AND TipoDocs.Diminutivo <> 'FCV')"
         Call FrmLCompVta.FViewLibroLeg(LIB_VENTAS, Mes, Ano, Wh)
         Set FrmLCompVta = Nothing
         
      Case OP_COMPRAS
         Set FrmLCompVta = New FrmCompraVenta
         Call FrmLCompVta.FViewLibroLeg(LIB_COMPRAS, Mes, Ano)
         Set FrmLCompVta = Nothing
         
      Case OP_RETENCIONES
         Set FrmLReten = New FrmLibRetenciones
         Call FrmLReten.FViewLibroLeg(Mes, Ano)
         Set FrmLReten = Nothing
         
      Case OP_INVENTARIOBAL
         Set FrmInv = New FrmLibInvBal
         FrmInv.FView (Mes)
         Set FrmInv = Nothing
         
      Case OP_EXPORTA
         Set FrmLCompVta = New FrmCompraVenta
         Call FrmLCompVta.FViewLibroLeg(LIB_VENTAS, Mes, Ano, " AND DocImpExp<>0", "Libro de Ventas Exportación")
         Set FrmLCompVta = Nothing
         
      Case OP_BOLETA
         Set FrmLCompVta = New FrmCompraVenta
         Call FrmLCompVta.FViewLibroLeg(LIB_VENTAS, Mes, Ano, " AND DocBoletas<>0", "Libro de Ventas con Boletas")
         Set FrmLCompVta = Nothing
      
       Case OP_FCV
         Set FrmLCompVta = New FrmCompraVenta
         Call FrmLCompVta.FViewLibroLeg(LIB_VENTAS, Mes, Ano, " AND TipoDocs.Diminutivo = 'FCV'", "Libro Especial de Compras")
         Set FrmLCompVta = Nothing
     
   End Select
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()
   Dim MesActual As Integer

   If lPeriodo = True Then
   
      MesActual = GetMesActual()
      
      Call FillMes(Cb_Mes)
      If MesActual > 0 Then
         Cb_Mes.ListIndex = MesActual - 1
      Else
         Cb_Mes.ListIndex = GetUltimoMesConMovs() - 1
      End If
      
      Cb_Ano.AddItem gEmpresa.Ano
      Cb_Ano.ListIndex = 0
  
   Else
      Fr_Periodo.Visible = False
      Me.Height = Me.Height - Fr_Periodo.Height - 300
   End If
   
End Sub

Private Sub Op_Libros_Click(Index As Integer)
   lIdx = Index
End Sub

Private Sub Op_Libros_DblClick(Index As Integer)
   lIdx = Index
   Call PostClick(Bt_View)
End Sub

Public Function FSelect() As Integer

   Me.Show vbModeless
     
End Function
Public Function FSelectMes() As Integer

   lPeriodo = True
   Me.Show vbModeless
   
End Function

