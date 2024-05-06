VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLstDocCuotas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Pago de Documentos a Plazo"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   Icon            =   "FrmLstDocCuotas.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_DocCuotas 
      Caption         =   "Detalle de Cuotas"
      Height          =   810
      Index           =   1
      Left            =   11820
      Picture         =   "FrmLstDocCuotas.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Ver detalle de Cuotas Documento"
      Top             =   2700
      Width           =   1335
   End
   Begin VB.PictureBox Pc_Check 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   11520
      Picture         =   "FrmLstDocCuotas.frx":04E9
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   46
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Fr_Botones 
      Height          =   555
      Left            =   60
      TabIndex        =   32
      Top             =   0
      Width           =   13095
      Begin VB.CommandButton Bt_DocCuotas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   540
         Picture         =   "FrmLstDocCuotas.frx":0560
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Ver detalle de Cuotas Documento"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Close 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   11760
         TabIndex        =   45
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton Bt_DetDoc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "FrmLstDocCuotas.frx":0A3D
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Detalle comprobante seleccionado"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Calc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         Picture         =   "FrmLstDocCuotas.frx":0EA2
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Calculadora"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_ConvMoneda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3540
         Picture         =   "FrmLstDocCuotas.frx":1203
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Convertir moneda"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Calendar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         Picture         =   "FrmLstDocCuotas.frx":15A1
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Calendario"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Preview 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Picture         =   "FrmLstDocCuotas.frx":19CA
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Print 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2580
         Picture         =   "FrmLstDocCuotas.frx":1E71
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Imprimir listado"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_CopyExcel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Picture         =   "FrmLstDocCuotas.frx":232B
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Copiar Excel"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Orden 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         Picture         =   "FrmLstDocCuotas.frx":2770
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ordenar listado por columna seleccionada"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Sum 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Picture         =   "FrmLstDocCuotas.frx":2B60
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   120
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4875
      Left            =   0
      TabIndex        =   21
      Top             =   3600
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   8599
      _Version        =   393216
      Rows            =   15
      Cols            =   15
      FixedCols       =   6
      SelectionMode   =   1
   End
   Begin VB.Frame Fr_Filtro 
      Caption         =   "Listar por"
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   60
      TabIndex        =   31
      Top             =   660
      Width           =   11655
      Begin VB.Frame Frame4 
         Caption         =   "Cuotas"
         Height          =   1275
         Left            =   9060
         TabIndex        =   50
         Top             =   1500
         Width           =   2415
         Begin VB.ComboBox Cb_DocCuotas 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   540
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Ver cuotas:"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   53
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Información adicional"
         Height          =   1275
         Left            =   180
         TabIndex        =   38
         Top             =   1500
         Width           =   8715
         Begin VB.CheckBox Ch_SaldosVig 
            Caption         =   "Saldos Vigentes"
            Height          =   195
            Left            =   7080
            TabIndex        =   18
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox Tx_Valor 
            Height          =   315
            Left            =   5280
            TabIndex        =   17
            Top             =   780
            Width           =   1335
         End
         Begin VB.TextBox Tx_Descrip 
            Height          =   315
            Left            =   5280
            MaxLength       =   100
            TabIndex        =   12
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox Tx_FEmision 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Tx_FVenc 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   13
            Top             =   780
            Width           =   1095
         End
         Begin VB.CommandButton Bt_FechaE 
            Caption         =   "?"
            Height          =   315
            Index           =   0
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   360
            Width           =   215
         End
         Begin VB.CommandButton Bt_FechaV 
            Caption         =   "?"
            Height          =   315
            Index           =   0
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   780
            Width           =   230
         End
         Begin VB.CommandButton Bt_FechaE 
            Caption         =   "?"
            Height          =   315
            Index           =   1
            Left            =   3900
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   360
            Width           =   215
         End
         Begin VB.TextBox Tx_FEmision 
            Height          =   315
            Index           =   1
            Left            =   2820
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton Bt_FechaV 
            Caption         =   "?"
            Height          =   315
            Index           =   1
            Left            =   3900
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   780
            Width           =   230
         End
         Begin VB.TextBox Tx_FVenc 
            Height          =   315
            Index           =   1
            Left            =   2820
            TabIndex        =   15
            Top             =   780
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor total:"
            Height          =   195
            Index           =   9
            Left            =   4380
            TabIndex        =   44
            Top             =   840
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   3
            Left            =   4320
            TabIndex        =   43
            Top             =   420
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Emisión:"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   42
            Top             =   420
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Venc.:"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   41
            Top             =   840
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "-->"
            Height          =   195
            Index           =   10
            Left            =   2640
            TabIndex        =   40
            Top             =   420
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "-->"
            Height          =   195
            Index           =   11
            Left            =   2640
            TabIndex        =   39
            Top             =   840
            Width           =   180
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Documento"
         Height          =   1215
         Left            =   4860
         TabIndex        =   35
         Top             =   240
         Width           =   5415
         Begin VB.ComboBox Cb_Estado 
            Height          =   315
            Left            =   780
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Tx_NumDoc 
            Height          =   315
            Left            =   3420
            MaxLength       =   15
            TabIndex        =   7
            Top             =   720
            Width           =   1815
         End
         Begin VB.ComboBox Cb_TipoDoc 
            Height          =   315
            Left            =   3420
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   300
            Width           =   1815
         End
         Begin VB.ComboBox Cb_TipoLib 
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   300
            Width           =   2415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Index           =   6
            Left            =   180
            TabIndex        =   47
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N° Doc.:"
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   37
            Top             =   780
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Doc.:"
            Height          =   195
            Index           =   0
            Left            =   2640
            TabIndex        =   36
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Entidad"
         Height          =   1215
         Left            =   180
         TabIndex        =   33
         Top             =   240
         Width           =   4515
         Begin VB.CheckBox Ch_Rut 
            Caption         =   "RUT:"
            CausesValidation=   0   'False
            Height          =   255
            Left            =   2220
            TabIndex        =   1
            Top             =   360
            Width           =   225
         End
         Begin VB.ComboBox Cb_Nombre 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   3135
         End
         Begin VB.ComboBox Cb_Entidad 
            Height          =   315
            Left            =   180
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   300
            Width           =   1575
         End
         Begin VB.TextBox Tx_Rut 
            Height          =   315
            Left            =   3000
            MaxLength       =   12
            TabIndex        =   2
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "RUT:"
            Height          =   195
            Left            =   2460
            TabIndex        =   49
            Top             =   360
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Razón Social:"
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   34
            Top             =   780
            Width           =   990
         End
      End
      Begin VB.CommandButton Bt_Search 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   10380
         Picture         =   "FrmLstDocCuotas.frx":2C04
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   300
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   60
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   8520
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   14
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmLstDocCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDDOC = 0
Const C_RUT = 1
Const C_ENTIDAD = 2
Const C_TIPOLIB = 3
Const C_IDTIPOLIB = 4
Const C_TIPODOC = 5
Const C_NUMDOC = 6
Const C_FEMISION = 7
Const C_VALOR = 8
Const C_NUMCUOTAS = 9
Const C_IDDOCCUOTA = 10
Const C_CUOTA = 11
Const C_MONTOCUOTA = 12
Const C_ESTADOCUOTA = 13
Const C_FVENC = 14
Const C_SALDO = 15
Const C_ESTADO = 16
Const C_IDESTADO = 17
Const C_DOCASOC = 18
Const C_DESC = 19

Const NCOLS = C_DESC

Const F_INICIO = 0
Const F_FIN = 1


Const M_IDENTIDAD = 1
Const M_RUT = 2
Const M_NOTVALIDRUT = 3

Dim lOrdenGr(C_DESC) As String
Dim lOrdenSel As Integer    'orden seleccionado o actual

Dim lOper As Integer
Dim lRc As Integer
Dim lcbNombre As ClsCombo
Dim lTipoLib As Integer
Dim lTipoDoc As Integer
Dim lEstadoDoc As Integer
Dim lValidarEstadoDoc As Boolean
Dim lNewDocLib As Boolean 'indca si se permite crear nuevos docs de libros auxiliares (compras, ventas y retenciones)
Dim lShowBtNew As Boolean

Dim lLstIdDoc() As LstDoc_t

Dim lOrientacion As Integer

Dim lMes As Integer

Dim lTogCheck As Boolean
Dim lAño As Integer
Public Function FView(Optional ByVal TipoLib As Integer = 0, Optional ByVal TipoDoc As Integer = 0, Optional ByVal EstadoDoc As Integer = 0, Optional ByVal Mes As Integer = 0, Optional ByVal Año As Integer = 0) As Integer
   
   lOper = O_VIEW
   lTipoLib = TipoLib
   lTipoDoc = TipoDoc
   lEstadoDoc = EstadoDoc
   lMes = Mes
   lAño = Año
   
   Me.Show vbModal
   
   FView = vbOK
   
End Function
Private Sub SelCbEntidad(Clasif As Integer)
   Dim Q1 As String
   
   lcbNombre.Clear
   If Clasif >= 0 Then
      Q1 = "SELECT Nombre, idEntidad, Rut, abs(NotValidRut) FROM Entidades"
      Q1 = Q1 & " WHERE Clasif" & Clasif & "=" & CON_CLASIF
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.Id
      Q1 = Q1 & " ORDER BY Nombre "
      Call lcbNombre.FillCombo(DbMain, Q1, -1)
   End If
End Sub

Private Sub Bt_Close_Click()
   lRc = vbCancel
   
   Unload Me

End Sub

Private Sub Bt_ConvMoneda_Click()
   Dim Frm As FrmConverMoneda
   Dim Col As Integer
   Dim Row As Integer
   Dim Valor As Double
   
   Col = Grid.Col
   Row = Grid.Row
   
   If Col = C_VALOR Then
      Valor = vFmt(Grid.TextMatrix(Row, Col))
   End If
   
   Set Frm = New FrmConverMoneda
   Frm.FSelect (Valor)
   Set Frm = Nothing

End Sub

Private Sub Bt_CopyExcel_Click()
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de copiar.", vbExclamation
      Exit Sub
   End If
   
   Call FGr2Clip(Grid, Me.Caption)
End Sub
Private Sub Bt_DetDoc_Click()
   Dim Frm As FrmDoc
   Dim IdDoc As Long
   Dim Rc As Integer
   
   IdDoc = Val(Grid.TextMatrix(Grid.Row, C_IDDOC))
   If IdDoc <= 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmDoc
   Call Frm.FView(IdDoc)
   Set Frm = Nothing
      
End Sub


Private Sub Bt_DocCuotas_Click(Index As Integer)
   Dim Frm As FrmDocCuotas
   Dim IdDoc As Long
   Dim Rc As Integer
   Dim Row As Integer
   Dim Msg As String
   Dim FVenc As Long
   Dim NumCuotas As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   IdDoc = Val(Grid.TextMatrix(Row, C_IDDOC))
   If IdDoc = 0 Then    'registro en blanco
      Exit Sub
   End If
   
   Set Frm = New FrmDocCuotas
   Call Frm.FView(IdDoc)
   Set Frm = Nothing
         
End Sub


Private Sub Bt_Orden_Click()
   Call OrdenaPorCol(Grid.Col)
End Sub

Private Sub Bt_Print_Click()
   Dim OldOrientation As Integer
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   OldOrientation = Printer.Orientation
   
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
End Sub

Private Sub Bt_Search_Click()
   
   Me.MousePointer = vbHourglass
   
   If Valida() Then
      Call LoadGrid
   End If
   
   Me.MousePointer = vbDefault

End Sub

Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid)
   
   Set Frm = Nothing

End Sub
Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de seleccionar la vista previa.", vbExclamation
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   
End Sub


Private Sub Cb_DocCuotas_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_Entidad_Click()
      
   Cb_Nombre.Clear
   If ItemData(Cb_Entidad) >= 0 Then
      Call SelCbEntidad(ItemData(Cb_Entidad))
   Else
      Tx_Rut = ""
   End If
   
   Call EnableFrm(True)
   
End Sub
Private Sub Cb_Estado_Click()
   Call EnableFrm(True)
   
   If ItemData(Cb_Estado) = ED_CENTPAG Or ItemData(Cb_Estado) = ED_PAGADO Then
      Ch_SaldosVig = 1
   End If

End Sub

Private Sub cb_Nombre_Click()
   
   If lcbNombre.ListIndex >= 0 Then
      Tx_Rut = FmtCID(lcbNombre.Matrix(M_RUT), Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0)
      Ch_Rut = IIf(Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0, 1, 0)
   End If
   
   Call EnableFrm(True)
   
End Sub
Private Sub Cb_TipoDoc_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_TipoLib_Click()
   Dim Q1 As String
   Dim i As Integer
   Dim TipoLib As Integer
   
   Cb_TipoDoc.Clear
   
   TipoLib = ItemData(Cb_TipoLib)
   
   If TipoLib > 0 Then
   
      Call FillTipoDoc(Cb_TipoDoc, TipoLib, True, True)
      Cb_TipoDoc.ListIndex = -1
      
      If TipoLib = LIB_OTROS And Cb_Estado.ListCount > 0 Then   'dejamos sin selección de estado
         Cb_Estado.ListIndex = 0
      End If
   End If
      
   Call EnableFrm(True)
   
End Sub


Private Sub Ch_Rut_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_SaldosVig_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_VerCuotasPagadas_Click()
   Call EnableFrm(True)

End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim Q1 As String
   Dim F1 As Long
   Dim F2 As Long
   Dim StrSort As String
   Dim MesActual As Integer
      
  lOrientacion = ORIENT_HOR
   
   Call BtFechaImg(Bt_FechaE(F_INICIO))
   Call BtFechaImg(Bt_FechaE(F_FIN))
   Call BtFechaImg(Bt_FechaV(F_INICIO))
   Call BtFechaImg(Bt_FechaV(F_FIN))
   
   'SE LLENA COMBOS
   Set lcbNombre = New ClsCombo
   Call lcbNombre.SetControl(Cb_Nombre)
   
   Ch_Rut = 1
   
   Cb_Entidad.AddItem ""
   Cb_Entidad.ItemData(Cb_Entidad.NewIndex) = -1
   For i = ENT_CLIENTE To ENT_OTRO
      Cb_Entidad.AddItem gClasifEnt(i)
      Cb_Entidad.ItemData(Cb_Entidad.NewIndex) = i
      
   Next i
   Cb_Entidad.ListIndex = 0     'para no seleccionar ninguno al partir
   
   Cb_TipoLib.AddItem ""
   Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = 0
   For i = 1 To UBound(gTipoLibNew)
      Cb_TipoLib.AddItem ReplaceStr(gTipoLibNew(i).Nombre, "Libro de ", "")
      Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = gTipoLibNew(i).Id 'i
   Next i
   
   Cb_TipoLib.ListIndex = 0
   If lTipoLib > 0 Then
      Call SelItem(Cb_TipoLib, lTipoLib)
   End If
   
   If lTipoDoc > 0 Then
      Call SelItem(Cb_TipoDoc, lTipoDoc)
   End If
      
   Cb_Estado.AddItem ""
   Cb_Estado.ItemData(Cb_Estado.NewIndex) = 0
   
   For i = 1 To MAX_ESTADODOC
      Cb_Estado.AddItem gEstadoDoc(i)
      Cb_Estado.ItemData(Cb_Estado.NewIndex) = i
   Next i
   
   Cb_Estado.AddItem "Centraliz. y Pagados"
   Cb_Estado.ItemData(Cb_Estado.NewIndex) = ED_CENTPAG
   
   Cb_Estado.ListIndex = 0
   
   If lEstadoDoc <> 0 Then
      Call SelItem(Cb_Estado, lEstadoDoc)
   End If
   
   Call CbAddItem(Cb_DocCuotas, "(todas)", 0)
   Call CbAddItem(Cb_DocCuotas, "Pendientes", ED_PENDIENTE)
   Call CbAddItem(Cb_DocCuotas, "Pagadas", ED_PAGADO)
   Cb_DocCuotas.ListIndex = 1
   
   If lMes > 0 And lAño > 0 Then
      Call FirstLastMonthDay(DateSerial(lAño, lMes, 1), F1, F2)
      Call SetTxDate(Tx_FEmision(F_INICIO), F1)
      Call SetTxDate(Tx_FEmision(F_FIN), F2)
   Else
      MesActual = GetMesActual()
      If MesActual > 0 Then
         Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, MesActual, 1), F1, F2)
      Else
         Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, GetUltimoMesConMovs(), 1), F1, F2)
      End If
      Call SetTxDate(Tx_FEmision(F_INICIO), F1)
      Call SetTxDate(Tx_FEmision(F_FIN), F2)
   End If
   
   
   StrSort = "Documento.TipoLib, Documento.TipoDoc, Documento.NumDoc, NumCuota "
                              
   lOrdenGr(C_RUT) = "Entidades.RUT, " & StrSort
   lOrdenGr(C_ENTIDAD) = "Entidades.Nombre, " & StrSort
   lOrdenGr(C_TIPOLIB) = StrSort & ", Documento.FEmision, Entidades.Nombre"
   lOrdenGr(C_TIPODOC) = "Documento.TipoDoc, Documento.TipoLib, Documento.NumDoc, NumCuota "
   lOrdenGr(C_NUMDOC) = "Documento.NumDoc, Entidades.Nombre, NumCuota "
   lOrdenGr(C_FEMISION) = "Documento.FEmision, Entidades.Nombre, " & StrSort
   lOrdenGr(C_FVENC) = "Documento.FVenc, Entidades.Nombre, Documento.FEmision, NumCuota "
   lOrdenGr(C_VALOR) = "Documento.Total, Entidades.Nombre, Documento.FEmision, NumCuota "
   lOrdenGr(C_SALDO) = "Documento.SaldoDoc, Entidades.Nombre, Documento.FEmision, NumCuota "
   lOrdenGr(C_ESTADO) = "Documento.Estado, Entidades.Nombre, Documento.FEmision,NumCuota "
   lOrdenGr(C_DOCASOC) = StrSort
   lOrdenGr(C_DESC) = "Documento.Descrip, Entidades.Nombre, Documento.FEmision, NumCuota "
   lOrdenGr(C_CUOTA) = "NumCuota, Documento.TipoLib, Documento.TipoDoc, Documento.NumDoc "
   lOrdenGr(C_MONTOCUOTA) = "MontoCuota," & StrSort
   lOrdenGr(C_ESTADOCUOTA) = "DocCuotas.Estado," & StrSort
   
   lOrdenSel = C_TIPOLIB
   
   Call SetUpGrid
   
   Me.MousePointer = vbHourglass
   
   Call RecalcSaldos(gEmpresa.Id, gEmpresa.Ano)
   Call RecalcSaldosFulle(gEmpresa.Id, gEmpresa.Ano)
   
   Me.MousePointer = vbDefault

   Call LoadGrid

End Sub

Private Sub Bt_FechaE_Click(Index As Integer)
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FEmision(Index))
   
   Set Frm = Nothing
End Sub
Private Sub Bt_FechaV_Click(Index As Integer)
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FVenc(Index))
   
   Set Frm = Nothing
End Sub

Private Sub Form_Resize()
   
   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - GridTot.Height - 500
   Grid.Width = Me.Width - 230
   GridTot.Top = Grid.Top + Grid.Height + 30
   GridTot.Width = Grid.Width - 230
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows

End Sub

Private Sub Grid_DblClick()
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   Call PostClick(Bt_DocCuotas(1))

End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol

End Sub

Private Sub Tx_Descrip_Change()
   Call EnableFrm(True)

End Sub

Private Sub Tx_FEmision_Change(Index As Integer)
   Call EnableFrm(True)

End Sub

Private Sub Tx_FEmision_GotFocus(Index As Integer)
   Call DtGotFocus(Tx_FEmision(Index))
End Sub

Private Sub Tx_FEmision_KeyPress(Index As Integer, KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub

Private Sub Tx_FEmision_LostFocus(Index As Integer)
   
   If Trim$(Tx_FEmision(Index)) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FEmision(Index))
   
End Sub

Private Sub Tx_FVenc_Change(Index As Integer)
   Call EnableFrm(True)

End Sub

Private Sub Tx_FVenc_GotFocus(Index As Integer)
   Call DtGotFocus(Tx_FVenc(Index))
End Sub

Private Sub Tx_FVenc_LostFocus(Index As Integer)
   
   If Trim$(Tx_FVenc(Index)) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FVenc(Index))
   
End Sub

Private Sub Tx_FVenc_KeyPress(Index As Integer, KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub


Private Sub Tx_NumDoc_Change()
   Call EnableFrm(True)

End Sub

Private Sub Tx_NumDoc_KeyPress(KeyAscii As Integer)
   Call KeyName(KeyAscii)
End Sub

Private Sub Tx_Rut_Change()
   Call EnableFrm(True)

End Sub

Private Sub Tx_Rut_LostFocus()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdEnt As Long
   Dim i As Integer
   Dim AuxRut As String

   If Tx_Rut = "" Then
      Cb_Entidad.ListIndex = 0  'en blanco
      Exit Sub
   End If
   
'   If Not MsgValidCID(Tx_Rut) Then
'      Tx_Rut.SetFocus
'      Exit Sub
'
'   End If
      
   Q1 = "SELECT IdEntidad, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5 FROM Entidades WHERE Rut = '" & vFmtCID(Tx_Rut, Ch_Rut <> 0) & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.Id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   IdEnt = 0
   
   If Rs.EOF = False Then   'existe
      IdEnt = vFld(Rs("IdEntidad"))
            
      'seleccionamos el tipo de entidad y esto llena la lista de nombres de entidades
      For i = 0 To MAX_ENTCLASIF
         If Cb_Entidad.ItemData(i) >= 0 Then
            If vFld(Rs("Clasif" & Cb_Entidad.ItemData(i))) <> 0 Then
               Cb_Entidad.ListIndex = i
               Exit For
            End If
         End If
      Next i
   
      'ahora seleccionamos la entidad
      For i = 0 To Cb_Nombre.ListCount - 1
         If lcbNombre.Matrix(M_IDENTIDAD, i) = IdEnt Then
            lcbNombre.ListIndex = i
            Exit For
         End If
      Next i
      
      Call EnableFrm(True)
      
   Else
      MsgBox1 "Este RUT no ha sido ingresado al sistema.", vbExclamation + vbOKOnly
      Cb_Entidad.ListIndex = -1
      
   End If
      
   Call CloseRs(Rs)
         
   If Ch_Rut <> 0 Then
      AuxRut = FmtCID(vFmtCID(Tx_Rut))
      If AuxRut <> "0-0" Then
         Tx_Rut = AuxRut
      End If
   End If
      

End Sub

Private Sub Tx_Rut_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      Call Tx_Rut_LostFocus
      KeyAscii = 0
   ElseIf Ch_Rut <> 0 Then
      Call KeyCID(KeyAscii)
   Else
      Call KeyName(KeyAscii)
      Call KeyUpper(KeyAscii)
   End If
   
End Sub

Private Sub Bt_Calc_Click()
   Call Calculadora
End Sub

Private Sub Bt_Calendar_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
   
   Call Frm.SelDate(Fecha)
   
   Set Frm = Nothing
End Sub

Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(0) As String
   Dim Encabezados(2) As String
   
   Printer.Orientation = lOrientacion
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = "Listado de Documentos"
   gPrtReportes.Titulos = Titulos
      
   Encabezados(0) = "Docs. :" & vbTab & Cb_TipoLib
   Encabezados(1) = "Estado:" & vbTab & Cb_Estado
   If Trim(Tx_FEmision(0)) <> "" Then
      Encabezados(2) = "Fecha :" & vbTab & Tx_FEmision(0) & " - " & Tx_FEmision(1)
   End If
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
   Total(C_NUMDOC) = "Total"
   Total(C_VALOR) = GridTot.TextMatrix(0, C_VALOR)
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_IDDOC
   gPrtReportes.NTotLines = 1
   

End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
    
   Grid.ColWidth(C_IDDOC) = 0
   Grid.ColWidth(C_TIPOLIB) = 830
   Grid.ColWidth(C_IDTIPOLIB) = 0
   Grid.ColWidth(C_TIPODOC) = 450
   Grid.ColWidth(C_NUMDOC) = 950
   Grid.ColWidth(C_RUT) = 1100
   Grid.ColWidth(C_ENTIDAD) = 1800
   Grid.ColWidth(C_VALOR) = 1200
   Grid.ColWidth(C_SALDO) = 1200
   Grid.ColWidth(C_FEMISION) = 800
   Grid.ColWidth(C_NUMCUOTAS) = 0
   Grid.ColWidth(C_IDDOCCUOTA) = 0
   Grid.ColWidth(C_CUOTA) = 700
   Grid.ColWidth(C_MONTOCUOTA) = 1200
   Grid.ColWidth(C_ESTADOCUOTA) = 900
   Grid.ColWidth(C_FVENC) = 800
   Grid.ColWidth(C_ESTADO) = 830
   Grid.ColWidth(C_IDESTADO) = 0
   Grid.ColWidth(C_DOCASOC) = 1400
   Grid.ColWidth(C_DESC) = 2800
   
         
   Grid.ColAlignment(C_TIPOLIB) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPODOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_DESC) = flexAlignLeftCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDO) = flexAlignRightCenter
   Grid.ColAlignment(C_CUOTA) = flexAlignRightCenter
   Grid.ColAlignment(C_MONTOCUOTA) = flexAlignRightCenter
   Grid.ColAlignment(C_FEMISION) = flexAlignRightCenter
   Grid.ColAlignment(C_FVENC) = flexAlignRightCenter
   Grid.ColAlignment(C_ESTADO) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_TIPOLIB) = "Libro"
   Grid.TextMatrix(0, C_TIPODOC) = "TD"
   Grid.TextMatrix(0, C_NUMDOC) = "N° Doc."
   Grid.TextMatrix(0, C_ESTADOCUOTA) = "Est. Cuota"
   Grid.TextMatrix(0, C_ESTADO) = "Est. Doc."
   Grid.TextMatrix(0, C_RUT) = "RUT"
   Grid.TextMatrix(0, C_ENTIDAD) = "Razón Social"
   Grid.TextMatrix(0, C_VALOR) = "Total"
   Grid.TextMatrix(0, C_SALDO) = "Saldo Doc."
   Grid.TextMatrix(0, C_FEMISION) = "Emisión"
   Grid.TextMatrix(0, C_CUOTA) = "Cuota"
   Grid.TextMatrix(0, C_MONTOCUOTA) = "Monto Cuota"
   Grid.TextMatrix(0, C_FVENC) = "Vencim."
   Grid.TextMatrix(0, C_DOCASOC) = "Doc. Asoc."
   Grid.TextMatrix(0, C_DESC) = "Descripción"
        
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
End Sub
Private Sub LoadGrid(Optional ByVal IdDoc As Long = 0)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim Row As Integer
   Dim Total As Double
   Dim TotSaldo As Double
   Dim NotValidRut As Boolean
   Dim TmpTbl As String
         
   Grid.Redraw = False
   
   If Trim(Tx_Rut) <> "" Then
      IdEnt = GetIdEntidad(Trim(Tx_Rut), NombEnt, NotValidRut)
      Where = Where & " AND Documento.IdEntidad = " & IdEnt
   End If
      
   If ItemData(Cb_TipoLib) > 0 Then
      Where = Where & " AND Documento.TipoLIB = " & ItemData(Cb_TipoLib)
   End If
   
   If ItemData(Cb_TipoDoc) > 0 Then
      Where = Where & " AND Documento.TipoDoc = " & ItemData(Cb_TipoDoc)
   End If
   
   If ItemData(Cb_Estado) = ED_CENTPAG Then
      Where = Where & " AND Documento.Estado IN ( " & ED_CENTRALIZADO & "," & ED_PAGADO & ")"
   ElseIf ItemData(Cb_Estado) > 0 Then
      Where = Where & " AND Documento.Estado = " & ItemData(Cb_Estado)
   End If
   
   If Trim(Tx_NumDoc) <> "" Then
      Where = Where & " AND Documento.NumDoc = '" & Trim(Tx_NumDoc) & "'"
   End If
   
   If Tx_FEmision(F_INICIO) <> "" And Tx_FEmision(F_FIN) <> "" Then
      Where = Where & " AND (Documento.FEmisionOri BETWEEN " & GetTxDate(Tx_FEmision(F_INICIO)) & " AND " & GetTxDate(Tx_FEmision(F_FIN)) & ")"
   End If

   If Trim(Tx_FVenc(F_INICIO)) <> "" And Trim(Tx_FVenc(F_FIN)) <> "" Then
      Where = Where & " AND (Documento.FVenc BETWEEN " & GetTxDate(Tx_FVenc(F_INICIO)) & " AND " & GetTxDate(Tx_FVenc(F_FIN)) & ")"
   End If

   If Trim(Tx_Descrip) <> "" Then
      Where = Where & " AND " & GenLike(DbMain, Tx_Descrip, "Documento.Descrip", 3)
   End If
   
   If vFmt(Tx_Valor) <> 0 Then
      Where = Where & " AND Documento.Total = " & vFmt(Tx_Valor)
   End If
   
   If Ch_SaldosVig <> 0 Then
      Where = Where & " AND Documento.SaldoDoc <> 0 "
   End If
   
   If CbItemData(Cb_DocCuotas) > 0 Then
      Where = Where & " AND DocCuotas.Estado = " & CbItemData(Cb_DocCuotas)
   End If
   
   If Where <> "" Then
      Where = " WHERE " & Mid(Where, 6)
   End If
   
   

'   TmpTbl = DbGenTmpName2(gdbtype,"tdoccuota_")
'   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)
'

'   If Ch_VerCuotasPagadas = 0 Then
'      'seleccionamos las N primeras cuotas impagas
'      Q1 = "SELECT IdDocCuota, DocCuotas.IdDoc, DocCuotas.NumCuota, MontoCuota, FechaExigPago, Estado INTO " & TmpTbl
'      Q1 = Q1 & " FROM DocCuotas INNER JOIN vPrimeraDocCuota ON DocCuotas.IdDoc = vPrimeraDocCuota.IdDoc AND (DocCuotas.NumCuota BETWEEN vPrimeraDocCuota.NumCuota1 AND vPrimeraDocCuota.NumCuota1 + " & Val(Tx_NumCuotas) - 1 & ") "
'      Q1 = Q1 & " ORDER BY DocCuotas.IdDoc, DocCuotas.NumCuota"
'   Else
'      'seleccionamos las N últimas cuotas (pagadas o no pagadas)
'      Q1 = "SELECT IdDocCuota, DocCuotas.IdDoc, DocCuotas.NumCuota, MontoCuota, FechaExigPago, Estado INTO " & TmpTbl
'      Q1 = Q1 & " FROM DocCuotas INNER JOIN vUltimaDocCuota ON DocCuotas.IdDoc = vUltimaDocCuota.IdDoc AND (DocCuotas.NumCuota BETWEEN vUltimaDocCuota.NumCuota1 AND vUltimaDocCuota.NumCuota1 - " & Val(Tx_NumCuotas) - 1 & ") "
'      Q1 = Q1 & " ORDER BY DocCuotas.IdDoc, DocCuotas.NumCuota"
'   End If
'
'   Call ExecSQL(DbMain, Q1)
'




   Q1 = "SELECT Documento.IdDoc, TipoLib, TipoDoc, NumDoc, NumDocHasta, Documento.IdEntidad, Entidades.Rut, Entidades.Nombre "
   Q1 = Q1 & ", Entidades.NotValidRut, FEmision, FEmisionOri, FVenc, Total, Descrip, Documento.Estado as EstadoDoc, SaldoDoc, IdDocAsoc "
   Q1 = Q1 & ", Documento.NumCuotas, IdDocCuota, NumCuota, MontoCuota, FechaExigPago, DocCuotas.Estado as EstadoCuota"
   Q1 = Q1 & " FROM (Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " AND Entidades.IdEmpresa = Documento.IdEmpresa )"
   Q1 = Q1 & " LEFT JOIN DocCuotas ON Documento.IdDoc = DocCuotas.IdDoc  "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "DocCuotas")
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.Id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel) & ", IdDocCuota "

   Set Rs = OpenRs(DbMain, Q1)

   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   Do While Rs.EOF = False
      
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
      
      If IdDoc > 0 And vFld(Rs("IdDoc")) = IdDoc Then
         Row = i
      End If
      
      Grid.TextMatrix(i, C_TIPOLIB) = Left(ReplaceStr(gTipoLib(vFld(Rs("TipoLib"))), "Libro de ", ""), 9)
      Grid.TextMatrix(i, C_IDTIPOLIB) = vFld(Rs("TipoLib"))
      Grid.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
      Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
      
      If vFld(Rs("IdEntidad")) <> 0 Then
         Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = False)
         Grid.TextMatrix(i, C_ENTIDAD) = vFld(Rs("Nombre"), True)
      End If
      
      
      Grid.TextMatrix(i, C_FEMISION) = Format(vFld(Rs("FEmisionOri")), SDATEFMT)
      
      Grid.TextMatrix(i, C_NUMCUOTAS) = IIf(vFld(Rs("NumCuotas")) > 0, vFld(Rs("NumCuotas")), "")
      Grid.TextMatrix(i, C_IDDOCCUOTA) = vFld(Rs("IdDocCuota"))
      Grid.TextMatrix(i, C_CUOTA) = IIf(vFld(Rs("NumCuota")) > 0, vFld(Rs("NumCuota")) & " de " & vFld(Rs("NumCuotas")), "")
      Grid.TextMatrix(i, C_MONTOCUOTA) = IIf(vFld(Rs("MontoCuota")) > 0, Format(vFld(Rs("MontoCuota")), NUMFMT), "")
      Grid.TextMatrix(i, C_ESTADOCUOTA) = Left(gEstadoDoc(vFld(Rs("EstadoCuota"))), 9)

      
      If vFld(Rs("IdDocCuota")) > 0 Then
         If vFld(Rs("FechaExigPago")) > 0 Then
            Grid.TextMatrix(i, C_FVENC) = Format(vFld(Rs("FechaExigPago")), SDATEFMT)
         End If
      
      ElseIf vFld(Rs("FVenc")) > 0 Then
         Grid.TextMatrix(i, C_FVENC) = Format(vFld(Rs("FVenc")), SDATEFMT)
      End If
      
      Grid.TextMatrix(i, C_VALOR) = Format(vFld(Rs("Total")), NUMFMT)
      Total = Total + vFld(Rs("Total"))
      
      Grid.TextMatrix(i, C_SALDO) = Format(vFld(Rs("SaldoDoc")), NEGNUMFMT)
      TotSaldo = TotSaldo + vFld(Rs("SaldoDoc"))
                  
      Grid.TextMatrix(i, C_DESC) = vFld(Rs("Descrip"), True)
      Grid.TextMatrix(i, C_ESTADO) = Left(gEstadoDoc(vFld(Rs("EstadoDoc"))), 9)
      Grid.TextMatrix(i, C_IDESTADO) = vFld(Rs("EstadoDoc"))
      
      If vFld(Rs("IdDocAsoc")) <> 0 Then
         Grid.TextMatrix(i, C_DOCASOC) = GetInfoDoc(vFld(Rs("IdDocAsoc")))
      End If
             
      Rs.MoveNext
      i = i + 1
   Loop

   Call CloseRs(Rs)
'   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)
   
   GridTot.TextMatrix(0, C_FVENC) = "TOTAL"
   GridTot.TextMatrix(0, C_VALOR) = Format(Total, NUMFMT)
   GridTot.TextMatrix(0, C_SALDO) = Format(TotSaldo, NEGNUMFMT)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
   'Marco la columna Ordenada
   If lOrdenSel <> C_DOCASOC Then
      Grid.Row = 0
      Grid.Col = lOrdenSel
      Set Grid.CellPicture = FrmMain.Pc_Flecha
   End If
   
   If Row = 0 Then
      Row = Grid.FixedRows
   End If
   
   Call FGrSelRow(Grid, Row)
      
   Grid.Redraw = True
   Call EnableFrm(False)
   
End Sub

Private Function Valida() As Boolean

   Valida = False
   
   If Trim(Tx_Rut) <> "" And Cb_Entidad.ListIndex < 0 Then
      MsgBox1 "RUT inválido.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If GetTxDate(Tx_FEmision(F_INICIO)) > GetTxDate(Tx_FEmision(F_FIN)) Then
      MsgBox1 "Rango de fecha de emisión inválido.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If GetTxDate(Tx_FVenc(F_INICIO)) > GetTxDate(Tx_FVenc(F_FIN)) Then
      MsgBox1 "Rango de fecha de vencimiento inválido.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   Valida = True

End Function

Private Sub Tx_Valor_Change()
   Call EnableFrm(True)

End Sub

Private Sub Tx_Valor_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      Call Tx_Valor_LostFocus
      KeyAscii = 0
   Else
      Call KeyNum(KeyAscii)
   End If
   

End Sub

Private Sub Tx_Valor_LostFocus()
   Tx_Valor = Format(vFmt(Tx_Valor), NUMFMT)
   
End Sub
Private Sub Grid_Click()
   Dim Col As Integer
   Dim Row As Integer
         
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Row >= Grid.FixedRows Then
      Exit Sub
   End If

   Call OrdenaPorCol(Col)
   
   
End Sub
Private Sub OrdenaPorCol(ByVal Col As Integer)
   
   Me.MousePointer = vbHourglass
   
   'Desmarco  columna Ordenada
   If lOrdenSel <> C_DOCASOC Then
      Grid.Row = 0
      Grid.Col = lOrdenSel
      Set Grid.CellPicture = LoadPicture()
   End If
   
   lOrdenSel = Col
   
   Call LoadGrid
      
   Me.MousePointer = vbDefault
      
End Sub
Private Sub EnableFrm(bool As Boolean)

   Bt_Search.Enabled = bool
   
End Sub

Private Function EstadoValido(ByVal Row As Integer, Optional ByVal ShowMsg As Boolean = True) As Boolean
   Dim TipoLib As Integer

   If lValidarEstadoDoc = False Then
      EstadoValido = True
      Exit Function
   End If
   
   'si no es libro de compras, ventas o retenciones, puede tener cualquier estado
   TipoLib = Val(Grid.TextMatrix(Row, C_IDTIPOLIB))

   If TipoLib <> LIB_COMPRAS And TipoLib <> LIB_VENTAS And TipoLib <> LIB_RETEN Then
      EstadoValido = True
      Exit Function
   End If
      
   'es libro compras, ventas o retenciones => estado IN(Centralizado o Pagado)
   EstadoValido = False
   
   'cambiamos la validación por warning para permitir pagar un doc antes de centralizarlo
   
   If Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_CENTRALIZADO And Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_PAGADO Then
      If ShowMsg Then
         If MsgBox1("Atención:" & vbNewLine & vbNewLine & "El documento debiera tener estado " & gEstadoDoc(ED_CENTRALIZADO) & " o " & gEstadoDoc(ED_PAGADO) & ". Este documento está en estado " & gEstadoDoc(Val(Grid.TextMatrix(Row, C_IDESTADO))) & "." & vbNewLine & vbNewLine & "¿Desea continuar?", vbExclamation + vbYesNo) = vbNo Then
            Exit Function
         End If
      Else
         Exit Function
      End If
   End If

   EstadoValido = True
   
End Function

Private Sub Tx_RUT_Validate(Cancel As Boolean)
   
   If Tx_Rut = "" Then
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_Rut, Ch_Rut <> 0) Then
      Cancel = True
      Exit Sub
   End If
   
End Sub

