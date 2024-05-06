VERSION 5.00
Begin VB.Form FrmSelLibCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Libro de Caja para Edición"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   540
      Picture         =   "FrmSelLibCaja.frx":0000
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   10
      Top             =   480
      Width           =   585
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      Top             =   900
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "Seleccionar..."
      Height          =   315
      Left            =   5880
      TabIndex        =   4
      Top             =   540
      Width           =   1275
   End
   Begin VB.Frame Fr_LibCaja 
      Height          =   1515
      Left            =   1500
      TabIndex        =   9
      Top             =   420
      Width           =   4095
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Caja - Ingresos"
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   1
         Top             =   480
         Width           =   2235
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Caja - Egresos"
         Height          =   195
         Index           =   2
         Left            =   660
         TabIndex        =   0
         Top             =   900
         Width           =   2355
      End
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Periodo"
      Height          =   975
      Left            =   1500
      TabIndex        =   6
      Top             =   2280
      Width           =   4095
      Begin VB.ComboBox Cb_Mes 
         CausesValidation=   0   'False
         Height          =   315
         ItemData        =   "FrmSelLibCaja.frx":0545
         Left            =   600
         List            =   "FrmSelLibCaja.frx":0547
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   1335
      End
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   8
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   7
         Top             =   480
         Width           =   330
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Antes de ingresar recuerde calcular Proporcionalidad de IVA CF"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1500
      TabIndex        =   12
      Top             =   3780
      Width           =   4545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ATENCIÓN: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   11
      Top             =   3420
      Width           =   1065
   End
End
Attribute VB_Name = "FrmSelLibCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lTipoOper As Integer
Dim lRc As Integer
Dim lMes As Integer
Dim lAño As Integer
Dim lSelTipoOper As Boolean

Private Sub Bt_Cancelar_Click()

   lRc = vbCancel
   Unload Me

End Sub

Private Sub Bt_Sel_Click()
   Dim Lib As Integer
   Dim i As Integer
   Dim Frm As Form

   If lSelTipoOper Then
      If Op_Libros(TOPERCAJA_INGRESO).Value <> 0 Then
         lTipoOper = TOPERCAJA_INGRESO
      Else
         lTipoOper = TOPERCAJA_EGRESO
      End If
   End If
   
   lMes = ItemData(Cb_Mes)
   lAño = Val(Cb_Ano)
   
   lRc = vbOK
   
   If lSelTipoOper Then
      
      Me.MousePointer = vbHourglass

      Set Frm = New FrmLibCaja
      Call Frm.FEdit(lTipoOper, lMes, lAño)

      Me.MousePointer = vbDefault
         
   Else
      
      Me.MousePointer = vbHourglass
      
      Set Frm = New FrmLibCaja
      Call Frm.FView(lMes, lAño)
      
      Me.MousePointer = vbDefault

   End If
   
   Set Frm = Nothing
   
End Sub




Private Sub Cb_Ano_LostFocus()
vMontoBaseImpoIngreso = 0
   vMontoBaseImpoEgreso = 0
End Sub

Private Sub Cb_Mes_LostFocus()
vMontoBaseImpoIngreso = 0
   vMontoBaseImpoEgreso = 0
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim MesActual As Integer
   
   Op_Libros(1).Value = True
   
   If lSelTipoOper = False Then
      Fr_LibCaja.visible = False
      Fr_Periodo.Top = Fr_LibCaja.Top
      Me.Height = Me.Height - Fr_LibCaja.Height - 300
   End If
         
   MesActual = GetMesActual()
   
   For i = 1 To 12
      Cb_Mes.AddItem gNomMes(i)
      Cb_Mes.ItemData(Cb_Mes.NewIndex) = i
   Next i
   
   Cb_Mes.ListIndex = 0
   If MesActual > 0 Then
      Cb_Mes.ListIndex = MesActual - 1
   End If
   
   Cb_Ano.AddItem gEmpresa.Ano
   Cb_Ano.ListIndex = Cb_Ano.NewIndex
     
End Sub
Public Function FSelect() As Integer

   Me.Show vbModal
     
   FSelect = lRc
   
End Function

Public Function FSelectOper() As Integer

   lSelTipoOper = True
   Me.Show vbModal
      
   FSelectOper = lRc
   
End Function

Private Sub Op_Libros_DblClick(Index As Integer)
   Call PostClick(Bt_Sel)
End Sub
