VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmResOtrosImp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Otros Impuestos Compras - Ventas"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4035
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   7215
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   2835
         Left            =   240
         TabIndex        =   7
         Top             =   1020
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   3
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Mes:"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   6
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Año:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   5
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Lb_Ano 
         Caption         =   "2005"
         Height          =   195
         Left            =   1800
         TabIndex        =   4
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   7620
      TabIndex        =   1
      Top             =   300
      Width           =   1275
   End
   Begin VB.CommandButton Bt_ResLibAux 
      Caption         =   "Resumen Libros      Auxiliares..."
      Height          =   1035
      Left            =   7620
      Picture         =   "FrmResOtrosImp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3180
      Width           =   1275
   End
End
Attribute VB_Name = "FrmResOtrosImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_TIPOLIB = 0
Const C_CODVALLIB = 1
Const C_LIBRO = 2
Const C_DESC = 3
Const C_VALOR = 4

Dim lMes As Integer
Dim lAno As Integer
Dim lVerBotonRes As Boolean

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_ResLibAux_Click()
   Dim Frm As FrmResLibAux
   
   Me.MousePointer = vbHourglass
   
   Set Frm = New FrmResLibAux
   Frm.Show vbModal
   Set Frm = Nothing

   Me.MousePointer = vbDefault

End Sub

Private Sub Cb_Mes_Click()
   Call LoadValOImp

End Sub

Private Sub Form_Load()
   Dim MesActual As Integer

   MesActual = GetMesActual()
   
   Cb_Mes.AddItem " "
   Cb_Mes.ItemData(Cb_Mes.NewIndex) = 0
            
   Call FillMes(Cb_Mes)
               
   If lMes > 0 Then
      Cb_Mes.ListIndex = lMes
   Else
      If MesActual > 0 Then
         Cb_Mes.ListIndex = MesActual
      Else
         Cb_Mes.ListIndex = GetUltimoMesConMovs()
      End If
   End If

   Lb_Ano = lAno
   
   If Not lVerBotonRes Then
      Bt_ResLibAux.Visible = False
   End If
   
   Call SetupGrid
   
   Call LoadValOImp


End Sub
Private Sub SetupGrid()
   
   Call FGrSetup(Grid, True)
   
   Grid.ColWidth(C_TIPOLIB) = 0
   Grid.ColWidth(C_CODVALLIB) = 0
   Grid.ColWidth(C_LIBRO) = 1000
   Grid.ColWidth(C_DESC) = 4000
   Grid.ColWidth(C_VALOR) = 1300
   
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_DESC) = "Impuesto"
   Grid.TextMatrix(0, C_VALOR) = "Valor"
   
   Call FGrVRows(Grid, 2)

End Sub


Private Sub LoadValOImp()
   Dim Where As String
   Dim ResOImp() As ResOImp_t
   Dim i As Integer, j As Integer
   
   Where = " Year(FEmision) = " & lAno
   
'   If ItemData(Cb_TipoLib) > 0 Then
'
'      If ItemData(Cb_TipoLib) = T_COMPRASVENTAS Then
'         Where = Where & " AND Documento.TipoLib IN (" & LIB_COMPRAS & ", " & LIB_VENTAS & ")"
'      Else
'         Where = Where & " AND Documento.TipoLib = " & ItemData(Cb_TipoLib)
'      End If
'
'   End If
   
   If ItemData(Cb_Mes) > 0 Then
      Where = Where & " AND Month(FEmision) = " & ItemData(Cb_Mes)
   End If
   
   Call GenResOImp(Where, ResOImp)
   
   Grid.Redraw = False
   
   Grid.Rows = Grid.FixedRows
   
   For i = 0 To UBound(ResOImp)
      j = i + Grid.FixedRows
      If ResOImp(i).CodValLib <> 0 Then
         If Grid.TextMatrix(j, C_LIBRO) <> Grid.TextMatrix(j - 1, C_LIBRO) Then
            Grid.TextMatrix(j, C_LIBRO) = gTipoLib(ResOImp(i).TipoLib)
         End If
         Grid.TextMatrix(j, C_DESC) = ResOImp(i).DescValLib
         Grid.TextMatrix(j, C_VALOR) = Format(ResOImp(i).Valor, NEGNUMFMT)
      Else
         Exit For
      End If
   Next i
   
   Call FGrVRows(Grid, 2)
   Grid.Redraw = True
   
End Sub

Public Function FView(ByVal Mes As Integer, ByVal Ano As Integer, Optional ByVal VerBotonRes As Boolean = True)

   lMes = Mes
   lAno = Ano
   lVerBotonRes = VerBotonRes
   Me.Show vbModal

End Function

