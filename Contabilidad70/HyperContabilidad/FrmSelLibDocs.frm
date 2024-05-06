VERSION 5.00
Begin VB.Form FrmSelLibDocs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libros y Documentos"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "FrmSelLibDocs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Periodo"
      Height          =   975
      Left            =   1320
      TabIndex        =   12
      Top             =   4080
      Width           =   4095
      Begin VB.ComboBox Cb_Ano 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   420
         Width           =   1335
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   14
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   480
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3315
      Index           =   0
      Left            =   1320
      TabIndex        =   11
      Top             =   420
      Width           =   4095
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Otros Documentos Full"
         Height          =   195
         Index           =   6
         Left            =   660
         TabIndex        =   15
         Tag             =   "8"
         Top             =   2550
         Width           =   2295
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "TODOS"
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   5
         Tag             =   "0"
         Top             =   2580
         Width           =   1635
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Otros Documentos"
         Height          =   195
         Index           =   5
         Left            =   660
         TabIndex        =   4
         Tag             =   "5"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Compras"
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   0
         Tag             =   "1"
         Top             =   480
         Width           =   2355
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Ventas"
         Height          =   195
         Index           =   2
         Left            =   660
         TabIndex        =   1
         Tag             =   "2"
         Top             =   900
         Width           =   2235
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Docs. Remuneraciones"
         Height          =   195
         Index           =   4
         Left            =   660
         TabIndex        =   3
         Tag             =   "4"
         Top             =   1740
         Width           =   2295
      End
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Retenciones"
         Height          =   195
         Index           =   3
         Left            =   660
         TabIndex        =   2
         Tag             =   "3"
         Top             =   1320
         Width           =   2355
      End
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "Seleccionar..."
      Height          =   315
      Left            =   5700
      TabIndex        =   8
      Top             =   540
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5700
      TabIndex        =   9
      Top             =   900
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   360
      Picture         =   "FrmSelLibDocs.frx":000C
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   10
      Top             =   480
      Width           =   585
   End
End
Attribute VB_Name = "FrmSelLibDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lTipoLib As Integer
Dim lRc As Integer
Dim lMes As Integer
Dim lAño As Integer
Dim lPeriodo As Boolean
Dim lIncluyeTodos As Boolean

Private Sub Bt_Cancelar_Click()

   lRc = vbCancel
   Unload Me

End Sub

Private Sub Bt_Sel_Click()
   Dim Lib As Integer
   Dim i As Integer
   Dim IdDoc As Long
   Dim Frm As Form

   'For i = 0 To LIB_OTROS
   For i = 1 To UBound(gTipoLibNew)
      If Op_Libros(i).Value <> 0 Then
         lTipoLib = Op_Libros(i).Tag ' i
      End If
   Next i
     
   If lTipoLib = 0 Then
      lTipoLib = LIB_OTROS
   End If
   
   lMes = ItemData(Cb_Mes)
   lAño = Val(Cb_Ano)
   
   lRc = vbOK
   
   'Unload Me

   If lPeriodo Then
      'ahora lo hacemos acá
      If lTipoLib = LIB_COMPRAS Or lTipoLib = LIB_VENTAS Then
   
         If gCtasBas.IdCtaIVACred <= 0 Or gCtasBas.IdCtaIVADeb <= 0 Then
            MsgBox1 "No es posible ingresar documentos a los Libros de Compras y Ventas sin antes definir la configuración de las cuentas de IVA y Otros Impuestos." & vbNewLine & vbNewLine & "Utilice el botón ""Definir Cuentas Básicas"" provisto en el menú ""Configuración Inicial"".", vbExclamation + vbOKOnly
            Exit Sub
         End If
   
         Me.MousePointer = vbHourglass
   
         Set Frm = New FrmCompraVenta
         Call Frm.FEdit(lTipoLib, lMes, lAño, IdDoc)
   
         Me.MousePointer = vbDefault
   
      ElseIf lTipoLib = LIB_RETEN Then
   
         If gCtasBas.IdCtaImpRet <= 0 Or gCtasBas.IdCtaNetoHon <= 0 Then
            MsgBox1 "No es posible ingresar documentos al Libro de Retenciones sin antes definir la configuración de las cuentas de Impuesto Retenido y Neto Retención." & vbNewLine & vbNewLine & "Utilice el botón ""Definir Cuentas Básicas"" provisto en el menú ""Configuración Inicial"".", vbExclamation + vbOKOnly
            Exit Sub
         End If
   
         Me.MousePointer = vbHourglass
   
         Set Frm = New FrmLibRetenciones
         Call Frm.FEdit(lMes, lAño, IdDoc)
   
         Me.MousePointer = vbDefault
   
      Else
         Me.MousePointer = vbHourglass
   
         Set Frm = New FrmLstDoc
         Call Frm.FEdit(lTipoLib, lMes, lAño, True)
   
'         Set Frm = New FrmOtrosDocs
'         Call Frm.FEdit(lTipoLib, lMes, lAño, IdDoc)
   
         Me.MousePointer = vbDefault
   
      End If
      
   Else
      
   Me.MousePointer = vbHourglass
      
     If lTipoLib = LIB_COMPRAS Or lTipoLib = LIB_VENTAS Then
         Set Frm = New FrmCompraVenta
         Call Frm.FView(lTipoLib)
         
      ElseIf lTipoLib = LIB_RETEN Then
         Set Frm = New FrmLibRetenciones
         Call Frm.FView
         
      Else
         Set Frm = New FrmLstDoc
         Call Frm.FView(lTipoLib)
      End If
      
      Me.MousePointer = vbDefault

   End If
   
   Set Frm = Nothing
   
End Sub


Private Sub Form_Load()
   Dim i As Integer
   Dim MesActual As Integer
   Dim DbName As String
   
   If Not lIncluyeTodos Then
      Op_Libros(0).visible = False
   End If
   
   Op_Libros(1).Value = True
   
   If lPeriodo = True Then
   
      'Cb_Mes.AddItem gNomMes(Month(Now) - 1)
      'Cb_Mes.ItemData(Cb_Mes.NewIndex) = Month(Now) - 1
      
      MesActual = GetMesActual()
      'Cb_Mes.AddItem gNomMes(MesActual)
      'Cb_Mes.ItemData(Cb_Mes.NewIndex) = MesActual
      'Cb_Mes.ListIndex = 0
      
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
      Cb_Ano.AddItem gEmpresa.Ano - 1
      Cb_Ano.AddItem gEmpresa.Ano - 2
      Cb_Ano.AddItem gEmpresa.Ano - 3
      Cb_Ano.AddItem gEmpresa.Ano - 4
      Cb_Ano.AddItem gEmpresa.Ano - 5
      ' si se agregan mas años agregar tambien en el FrmLibRetenciones Metodo FillCb
  
   Else
      Fr_Periodo.visible = False
      Me.Height = Me.Height - Fr_Periodo.Height - 300
   End If
   
   #If DATACON = 1 Then       'Access

   If gEmpresa.TieneAnoAnt Then

         DbName = gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb"
         If ExistFile(DbName) Then
         
            Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano - 1)
            Call CorrigeBase
            Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)

         End If
   End If

#End If
   

   
End Sub
Public Function FSelect(TipoLib As Integer, Optional ByVal IncluyeTodos As Boolean = True) As Integer

   lIncluyeTodos = IncluyeTodos
   Me.Show vbModal
   
   TipoLib = lTipoLib
  
   FSelect = lRc
End Function
Public Function FSelectMes(TipoLib As Integer, Mes As Integer, Año As Integer, Optional ByVal IncluyeTodos As Boolean = True) As Integer

   lPeriodo = True
   lIncluyeTodos = IncluyeTodos
   Me.Show vbModal
   
   TipoLib = lTipoLib
   Mes = lMes
   Año = lAño
   
   FSelectMes = lRc
   
End Function

Private Sub Op_Libros_DblClick(Index As Integer)
   Call PostClick(Bt_Sel)
End Sub
