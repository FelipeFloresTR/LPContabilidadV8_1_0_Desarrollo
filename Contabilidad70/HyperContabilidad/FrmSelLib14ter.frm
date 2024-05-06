VERSION 5.00
Begin VB.Form FrmSelLib14ter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contribuyentes 14 TER A)"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   1380
      TabIndex        =   8
      Top             =   480
      Width           =   4035
      Begin VB.OptionButton Op_AsistImp1Cat 
         Caption         =   "Asistente Impuesto 1a Categoría"
         Height          =   315
         Left            =   420
         TabIndex        =   11
         Top             =   2880
         Width           =   3075
      End
      Begin VB.OptionButton Op_BaseImp 
         Caption         =   "Base Imponible 14 TER A)"
         Height          =   315
         Left            =   420
         TabIndex        =   4
         Top             =   2400
         Width           =   3075
      End
      Begin VB.OptionButton Op_Ajustes 
         Caption         =   "Ajustes Extra Libro de Caja"
         Height          =   315
         Left            =   420
         TabIndex        =   3
         Top             =   1920
         Width           =   3075
      End
      Begin VB.OptionButton Op_LibIngEg 
         Caption         =   "Libro de Ingresos y Egresos"
         Height          =   315
         Left            =   420
         TabIndex        =   2
         Top             =   1440
         Width           =   3075
      End
      Begin VB.OptionButton Op_ListLibCaja 
         Caption         =   "Listar Libro de Caja Consolidado"
         Height          =   315
         Left            =   420
         TabIndex        =   1
         Top             =   960
         Width           =   3075
      End
      Begin VB.OptionButton Op_EditLibCaja 
         Caption         =   "Ingresar Libro de Caja"
         Height          =   315
         Left            =   420
         TabIndex        =   0
         Top             =   480
         Width           =   3075
      End
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "Seleccionar..."
      Height          =   315
      Left            =   5760
      TabIndex        =   5
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5760
      TabIndex        =   6
      Top             =   960
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   420
      Picture         =   "FrmSelLib14ter.frx":0000
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   7
      Top             =   540
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Antes de ingresar recuerde calcular Proporcionalidad de IVA CF"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   10
      Top             =   4860
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
      Index           =   0
      Left            =   1380
      TabIndex        =   9
      Top             =   4500
      Width           =   1065
   End
End
Attribute VB_Name = "FrmSelLib14ter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lTitulo As String

Public Function FView(Optional ByVal Regimen As String = "")

   lTitulo = Regimen
      
   Me.Show vbModal
   
   
End Function

Private Sub Form_Load()
   Op_EditLibCaja = True
   
   If Not gFunciones.AjustesExtraLibCaja Then
      Op_Ajustes.visible = False
   End If
   
   If lTitulo <> "" Then
      Me.Caption = lTitulo
   End If
   
   If gEmpresa.Ano >= 2020 Then
      Op_Ajustes.visible = False
      Op_AsistImp1Cat.visible = False
      Op_BaseImp.Top = Op_Ajustes.Top
      Op_BaseImp.Caption = "Base Imponible 14 D"
   End If
   
End Sub

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Sel_Click()
   Dim Frm As Form
   
   If gEmpresa.Franq14Ter = 0 And gEmpresa.Ano < 2020 Then
      MsgBox1 "Empresa no acogida a Franquicia Artículo 14 TER, no lleva Libro de Caja.", vbInformation
      Exit Sub
   End If
      
   Me.MousePointer = vbHourglass
    
   If Op_EditLibCaja <> 0 Then
   
      Set Frm = New FrmSelLibCaja
      Call Frm.FSelectOper
      Set Frm = Nothing
      
   ElseIf Op_ListLibCaja <> 0 Then

      Set Frm = New FrmLibCaja
      Call Frm.FView
      Set Frm = Nothing

   ElseIf Op_LibIngEg <> 0 Then
      
      If gEmpresa.ObligaLibComprasVentas Then
         MsgBox1 "Empresa acogida a Franquicia Artículo 14 TER y obligada a llevar Libro de Compras y Ventas según la Ley de IVA, no lleva Libro de Ingresos y Egresos.", vbInformation
         
         Me.MousePointer = vbDefault
         Exit Sub
      End If
        
      Set Frm = New FrmLibIngEg
      Frm.Show vbModal
      Set Frm = Nothing
   
   ElseIf Op_Ajustes <> 0 Then
           
      Set Frm = New FrmAjustesExtraLibCaja
      Frm.Show vbModal
      Set Frm = Nothing
      
   ElseIf Op_BaseImp <> 0 Then
           
      If gEmpresa.Ano >= 2020 Then
            
         Set Frm = New FrmBaseImponible14DFull
         Frm.Show vbModal
         Set Frm = Nothing

         Me.MousePointer = vbDefault
         
         Exit Sub
      End If
         
      Set Frm = New FrmBaseImponible
      Frm.Show vbModal
      Set Frm = Nothing
   
   ElseIf Op_AsistImp1Cat <> 0 Then
           
      Set Frm = New FrmAsistImpPrimCat
      Frm.Show vbModal
      Set Frm = Nothing
   
   End If
  
   Me.MousePointer = vbDefault
  
End Sub

Private Sub Op_Ajustes_DblClick()
   Call PostClick(Bt_Sel)

End Sub

Private Sub Op_AsistImp1Cat_DblClick()
   Call PostClick(Bt_Sel)

End Sub

Private Sub Op_BaseImp_DblClick()
   Call PostClick(Bt_Sel)

End Sub

Private Sub Op_EditLibCaja_DblClick()
   Call PostClick(Bt_Sel)
End Sub



Private Sub Op_LibIngEg_DblClick()
   Call PostClick(Bt_Sel)

End Sub

Private Sub Op_ListLibCaja_DblClick()
   Call PostClick(Bt_Sel)

End Sub
