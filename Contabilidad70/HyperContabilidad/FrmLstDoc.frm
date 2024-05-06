VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLstDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listar Documentos"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   Icon            =   "FrmLstDoc.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13215
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Pc_Check 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   11520
      Picture         =   "FrmLstDoc.frx":000C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   52
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Pc_HdCheck 
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   11280
      Picture         =   "FrmLstDoc.frx":0083
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   38
      Top             =   480
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.CommandButton Bt_DelDoc 
      Caption         =   "&Eliminar Doc"
      Height          =   675
      Left            =   11880
      Picture         =   "FrmLstDoc.frx":03E8
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Bt_ModDoc 
      Caption         =   "&Modificar Doc"
      Height          =   675
      Left            =   11880
      Picture         =   "FrmLstDoc.frx":086C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Bt_NewDoc 
      Caption         =   "&Nuevo Doc"
      Height          =   675
      Left            =   11880
      Picture         =   "FrmLstDoc.frx":0CE0
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Fr_Botones 
      Height          =   555
      Left            =   60
      TabIndex        =   37
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
         Left            =   540
         Picture         =   "FrmLstDoc.frx":117D
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Ver detalle de Cuotas Documento"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Bt_Close 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   11760
         TabIndex        =   51
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
         Picture         =   "FrmLstDoc.frx":165A
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Picture         =   "FrmLstDoc.frx":1ABF
         Style           =   1  'Graphical
         TabIndex        =   34
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
         Picture         =   "FrmLstDoc.frx":1E20
         Style           =   1  'Graphical
         TabIndex        =   33
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
         Picture         =   "FrmLstDoc.frx":21BE
         Style           =   1  'Graphical
         TabIndex        =   35
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
         Picture         =   "FrmLstDoc.frx":25E7
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Picture         =   "FrmLstDoc.frx":2A8E
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Picture         =   "FrmLstDoc.frx":2F48
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Picture         =   "FrmLstDoc.frx":338D
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Picture         =   "FrmLstDoc.frx":377D
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "&Seleccionar"
      Height          =   675
      Left            =   11880
      Picture         =   "FrmLstDoc.frx":3821
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2880
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4875
      Left            =   0
      TabIndex        =   22
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
      TabIndex        =   36
      Top             =   660
      Width           =   11655
      Begin VB.Frame Frame4 
         Caption         =   "Cuotas"
         Height          =   1275
         Left            =   9060
         TabIndex        =   56
         Top             =   1500
         Width           =   2415
         Begin VB.TextBox Tx_NumCuotas 
            Height          =   315
            Left            =   1860
            MaxLength       =   2
            TabIndex        =   20
            Top             =   660
            Width           =   315
         End
         Begin VB.CheckBox Ch_VerCuotas 
            Caption         =   "Ver Detalle de  Cuotas"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Num. Cuotas por Doc.:"
            Height          =   195
            Left            =   180
            TabIndex        =   57
            Top             =   720
            Width           =   1620
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Información adicional Documento"
         Height          =   1275
         Left            =   180
         TabIndex        =   44
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
            TabIndex        =   50
            Top             =   840
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Index           =   3
            Left            =   4320
            TabIndex        =   49
            Top             =   420
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Emisión:"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   48
            Top             =   420
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Venc.:"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   47
            Top             =   840
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "-->"
            Height          =   195
            Index           =   10
            Left            =   2640
            TabIndex        =   46
            Top             =   420
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "-->"
            Height          =   195
            Index           =   11
            Left            =   2640
            TabIndex        =   45
            Top             =   840
            Width           =   180
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Documento"
         Height          =   1215
         Left            =   4860
         TabIndex        =   41
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
            TabIndex        =   53
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N° Doc.:"
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   43
            Top             =   780
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Doc.:"
            Height          =   195
            Index           =   0
            Left            =   2640
            TabIndex        =   42
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Entidad"
         Height          =   1215
         Left            =   180
         TabIndex        =   39
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
            TabIndex        =   55
            Top             =   360
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Razón Social:"
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   40
            Top             =   780
            Width           =   990
         End
      End
      Begin VB.CommandButton Bt_Search 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   10380
         Picture         =   "FrmLstDoc.frx":3CEF
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   300
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   60
      TabIndex        =   54
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
Attribute VB_Name = "FrmLstDoc"
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
Const C_CHECK = 7
Const C_FEMISION = 8
Const C_VALOR = 9
Const C_NUMCUOTAS = 10
Const C_IDDOCCUOTA = 11
Const C_CUOTA = 12
Const C_NUMCUOTA = 13
Const C_MONTOCUOTA = 14
Const C_FVENC = 15
Const C_SALDO = 16
Const C_ESTADO = 17
Const C_IDESTADO = 18
Const C_DOCASOC = 19
Const C_DESC = 20
'Const NCOLS = C_DESC
'feña
Const C_TRATAMIENTO = 21
Const NCOLS = C_TRATAMIENTO
'fin feña


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

Friend Function FSelect(LstIdDoc() As LstDoc_t, Optional ByVal TipoLib As Integer = 0, Optional ByVal EstadoDoc As Integer = 0, Optional ByVal ValidarEstadoDoc As Boolean = False, Optional ByVal ShowBtNew As Boolean = True) As Integer
   
   lOper = O_SELECT
   lTipoLib = TipoLib
   
   lEstadoDoc = EstadoDoc   'estado Doc preseleccionado
   lValidarEstadoDoc = ValidarEstadoDoc
   lShowBtNew = ShowBtNew
   
   Me.Show vbModal
   LstIdDoc = lLstIdDoc
   CodTipoLib = lTipoLib
   
   FSelect = lRc
   
End Function
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
Public Function FEdit(Optional ByVal TipoLib As Integer = 0, Optional ByVal Mes As Integer = 0, Optional ByVal Año As Integer = 0, Optional ByVal NewDocLib As Boolean = False) As Integer
   
   lOper = O_EDIT
   lTipoLib = TipoLib
   lMes = Mes
   lAño = Año
   lNewDocLib = NewDocLib
   
   Me.Show vbModal
   
   FEdit = vbOK
   
End Function
Private Sub SelCbEntidad(Clasif As Integer)
   Dim Q1 As String
   
   lcbNombre.Clear
   If Clasif >= 0 Then
      Q1 = "SELECT Nombre, idEntidad, Rut, abs(NotValidRut) FROM Entidades"
      Q1 = Q1 & " WHERE Clasif" & Clasif & "=" & CON_CLASIF
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
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

Private Sub Bt_DelDoc_Click()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim LstComp As String
   Dim IdDoc As Long
   
   IdDoc = Val(Grid.TextMatrix(Grid.Row, C_IDDOC))
   If IdDoc <= 0 Then
      Exit Sub
   End If
   
   'vemos si hay comprobantes que hacen referencia al documento
   Q1 = "SELECT DISTINCT Correlativo, Fecha "
   Q1 = Q1 & " FROM Comprobante INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE IdDoc = " & IdDoc
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   
'   If Val(Grid.TextMatrix(Grid.Row, C_IDTIPOLIB)) = LIB_OTROFULL Then
'    Q1 = Replace(Replace(Replace(Q1, "Comprobante", "ComprobanteFull"), "MovComprobante.", "MovComprobanteFull."), "MovComprobante ", "MovComprobanteFull ")
'   End If
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
      LstComp = LstComp & ", Comprobante N° " & vFld(Rs("Correlativo")) & " del " & Format(vFld(Rs("Fecha")), DATEFMT) & vbNewLine
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   If LstComp <> "" Then
      LstComp = Mid(LstComp, 3)
      MsgBox1 "Los siguientes comprobantes hacen referencia a este documento:" & vbNewLine & vbNewLine & LstComp, vbExclamation + vbOKOnly
      Exit Sub
   End If
      
   If MsgBox1("¿Está seguro que desea borrar este documento?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
     
'   Call ExecSQL(DbMain, "DELETE * FROM Documento WHERE IdDoc = " & IdDoc)
   Q1 = " WHERE IdDoc = " & IdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
                                                                  
   
   'If Val(Grid.TextMatrix(Grid.Row, C_IDTIPOLIB)) <> LIB_OTROFULL Then
       Call DeleteSQL(DbMain, "Documento", Q1)
       
    '   Call ExecSQL(DbMain, "DELETE * FROM MovDocumento WHERE IdDoc = " & IdDoc)
       Q1 = " WHERE IdDoc = " & IdDoc
       Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'ffv delete
       Call DeleteSQL(DbMain, "MovDocumento", Q1)
       
       '3133008
       
        Dim PathDbAnoAnt As String
        Dim ConnStr As String

        #If DATACON = 1 Then
        Dim DbAnoAnt As Database
        #Else
        Dim DbAnoAnt As ADODB.Connection
        Set DbAnoAnt = DbMain
        #End If

   If gDbType = SQL_ACCESS Then
        PathDbAnoAnt = Replace(Replace(Replace(gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb", "HyperContabilidad\", ""), "LPContabSQL", "LPContab"), "..\", "")

        If ExistFile(PathDbAnoAnt) Then
          ConnStr = ";PWD=" & PASSW_PREFIX & gEmpresa.Rut & ";"
          Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)

        Else
          'Exit Sub
        End If
    End If

    Q1 = ""
    Q1 = "Update Documento Set FExported = null WHERE NumDoc = '" & Val(Grid.TextMatrix(Grid.Row, C_NUMDOC)) & "'"
    Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1
    Q1 = Q1 & " And TipoLib = " & Val(Grid.TextMatrix(Grid.Row, C_IDTIPOLIB))
    Q1 = Q1 & " And TipoDoc = " & FindTipoDoc(Val(Grid.TextMatrix(Grid.Row, C_IDTIPOLIB)), Grid.TextMatrix(Grid.Row, C_TIPODOC))
    Q1 = Q1 & " And FEmisionOri = " & GetDate(Grid.TextMatrix(Grid.Row, C_FEMISION))
    Q1 = Q1 & " And Total = " & Abs(vFmt(Grid.TextMatrix(Grid.Row, C_VALOR)))

    Call ExecSQL(DbAnoAnt, Q1)
    If gDbType = SQL_ACCESS Then

    Call CloseDb(DbAnoAnt)
    End If

    '3133008
   
   'Else
     'Call DeleteSQL(DbMain, "DocumentoFull", Q1)
   'End If
   
   Call LoadGrid
   
End Sub
Private Sub Bt_DetDoc_Click()
   Dim Frm As FrmDoc
   Dim IdDoc As Long
   Dim Rc As Integer
   Dim TipoLib As Integer
   
   IdDoc = Val(Grid.TextMatrix(Grid.Row, C_IDDOC))
   If IdDoc <= 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmDoc
   TipoLib = 0
   If Val(Grid.TextMatrix(Grid.Row, C_IDTIPOLIB)) = LIB_OTROFULL Then
    TipoLib = 8
   End If
   Call Frm.FView(IdDoc, TipoLib)
   Set Frm = Nothing
      
   Me.MousePointer = vbHourglass
   Call LoadGrid
   Me.MousePointer = vbDefault
   
End Sub


Private Sub Bt_DocCuotas_Click()
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
   If Grid.TextMatrix(Row, C_IDDOC) = 0 Then    'registro en blanco
      Exit Sub
   End If
   
   Set Frm = New FrmDocCuotas
   Call Frm.FView(IdDoc)
   Set Frm = Nothing
         
End Sub

Private Sub Bt_ModDoc_Click()
   Dim Frm As FrmDoc
   Dim IdDoc As Long
   Dim Rc As Integer
   
   IdDoc = Val(Grid.TextMatrix(Grid.Row, C_IDDOC))
   If IdDoc <= 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmDoc
   'Rc = Frm.FEdit(IdDoc, lTipoLib)
   Rc = Frm.FEdit(IdDoc, ItemData(Cb_TipoLib))
   Set Frm = Nothing
   
   If Rc = vbOK And IdDoc > 0 Then
      Call LoadGrid(IdDoc)
   End If
End Sub
Private Sub Bt_NewDoc_Click()
   Dim Frm As Form
   Dim IdDoc As Long
   Dim Rc As Integer
   Dim TipoLib As Integer
   Dim Mes As Integer
   Dim Año As Integer
   
   If gEmpresa.FCierre <> 0 Then
      MsgBox1 "Este periodo está cerrado.", vbExclamation + vbOKOnly
      Exit Sub
   End If
      
   If ValidaIngresoDoc() = False Then
      Exit Sub
   End If
   
   If lOper = O_SELECT And lShowBtNew Then
      ReDim lLstIdDoc(0)
      lLstIdDoc(0).IdDoc = 0
      lRc = vbOK
      Unload Me
      Exit Sub
   End If

   TipoLib = ItemData(Cb_TipoLib)

   If lNewDocLib Then
      If TipoLib = 0 Then
   
         Set Frm = New FrmSelLibDocs
         Rc = Frm.FSelectMes(TipoLib, Mes, Año, True)
         Set Frm = Nothing
   
         If Rc <> vbOK Then
            Exit Sub
         End If
   
      End If
   End If
      
   If Mes = 0 Or Año = 0 Then
      Mes = GetMesActual()
      Año = gEmpresa.Ano
   End If
   
   If lNewDocLib Then
   
      If TipoLib = LIB_COMPRAS Or TipoLib = LIB_VENTAS Then
   
         If gCtasBas.IdCtaIVACred <= 0 Or gCtasBas.IdCtaIVADeb <= 0 Or gCtasBas.IdCtaOtrosImpCred <= 0 Or gCtasBas.IdCtaOtrosImpDeb <= 0 Then
            MsgBox1 "No es posible ingresar documentos a los Libros de Compras y Ventas sin antes definir la configuración de las cuentas de IVA y Otros Impuestos.", vbExclamation + vbOKOnly
            Exit Sub
         End If
   
         Set Frm = New FrmCompraVenta
         Call Frm.FEdit(TipoLib, Mes, Año, IdDoc)
   
      ElseIf TipoLib = LIB_RETEN Then
   
         If gCtasBas.IdCtaImpRet <= 0 Or gCtasBas.IdCtaNetoHon <= 0 Then
            MsgBox1 "No es posible ingresar documentos al Libro de Retenciones sin antes definir la configuración de las cuentas de Impuesto Retenido y Neto Retención.", vbExclamation + vbOKOnly
            Exit Sub
         End If
   
         Set Frm = New FrmLibRetenciones
         Call Frm.FEdit(Mes, Año, IdDoc)
   
      Else
         Set Frm = New FrmDoc
         Rc = Frm.FNew(TipoLib, IdDoc, False, Mes, Año)
      End If
      
   Else
      Set Frm = New FrmDoc
      Rc = Frm.FNew(TipoLib, IdDoc, False, Mes, Año)
   End If
   
   Set Frm = Nothing
      
   If Rc = vbOK And IdDoc > 0 Then
      Call LoadGrid(IdDoc)
   End If

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
   
   If valida() Then
      Call LoadGrid
   End If
   
   Me.MousePointer = vbDefault

End Sub

Private Sub Bt_Sel_Click()
   Dim IdDoc As Long, IdDocSel As Long
   Dim i As Integer
   Dim j As Integer
   Dim Row As Integer
   Dim TipoLib As String
   Dim PrimeraNumCuota As Integer, NumCuotaSel As Integer
   
   ReDim lLstIdDoc(100)
      
   Row = Grid.Row
   IdDoc = 0
   IdDocSel = 0
   PrimeraNumCuota = 0
   NumCuotaSel = 0
      
   j = 0
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_IDDOC) = "" Then
         Exit For
      End If
           
      If IdDoc <> Val(Grid.TextMatrix(i, C_IDDOC)) Then
         IdDoc = Val(Grid.TextMatrix(i, C_IDDOC))
         PrimeraNumCuota = Val(Grid.TextMatrix(i, C_NUMCUOTA))
         'lTipoLib = Val(Grid.TextMatrix(Grid.Row, C_IDTIPOLIB))
      End If
                 
      Grid.Row = i
      Grid.Col = C_CHECK
      
      If Grid.CellPicture <> 0 Then
      
         'verificamos que seleccione sólo docs de un mismo TipoLib
         If TipoLib = "" Then
            TipoLib = Grid.TextMatrix(i, C_TIPOLIB)
            lTipoLib = Val(Grid.TextMatrix(Grid.Row, C_IDTIPOLIB))
            
         ElseIf TipoLib <> Grid.TextMatrix(i, C_TIPOLIB) Then
            MsgBox1 "Los documentos seleccionados no son todos del mismo libro.", vbExclamation + vbOKOnly
            Exit Sub
         End If
            
         If Val(Grid.TextMatrix(Grid.Row, C_IDESTADO)) = ED_ANULADO Then
            MsgBox1 "Alguno de los documentos seleccionados está Anulado. Revise la lista antes de continuar.", vbExclamation + vbOKOnly
            Exit Sub
         End If
         
         If IdDocSel <> Val(Grid.TextMatrix(i, C_IDDOC)) Then
            IdDocSel = Val(Grid.TextMatrix(i, C_IDDOC))
            NumCuotaSel = Val(Grid.TextMatrix(i, C_NUMCUOTA))
            If NumCuotaSel > PrimeraNumCuota Then
               MsgBox1 "Entre documentos seleccionados, se ha saltado la primera cuota pendiente de pago. " & vbCrLf & vbCrLf & "Debe seleccionar las cuotas en forma ordenada y correlativa.", vbExclamation
               Exit Sub
            End If
         ElseIf Val(Grid.TextMatrix(i, C_NUMCUOTA)) > NumCuotaSel + 1 Then
            MsgBox1 "Entre las cuotas seleccionadas, hay algunas que están saltadas. " & vbCrLf & vbCrLf & "Debe seleccionar las cuotas en forma ordenada y correlativa.", vbExclamation
            Exit Sub
         Else
            NumCuotaSel = Val(Grid.TextMatrix(i, C_NUMCUOTA))
         End If
               
        If j > UBound(lLstIdDoc) Then
            ReDim Preserve lLstIdDoc(j + 10)
         End If
         
         lLstIdDoc(j).IdDoc = Val(Grid.TextMatrix(i, C_IDDOC))
         lLstIdDoc(j).IdDocCuota = Val(Grid.TextMatrix(i, C_IDDOCCUOTA))
         lLstIdDoc(j).MontoCuota = vFmt(Grid.TextMatrix(i, C_MONTOCUOTA))
         lLstIdDoc(j).NumCuotas = vFmt(Grid.TextMatrix(i, C_NUMCUOTAS))
   
         j = j + 1
      End If
      
   Next i
   
   Grid.Row = Row
   
   If lLstIdDoc(0).IdDoc = 0 Then    'no marcó ninguno
   
      IdDoc = Val(Grid.TextMatrix(Grid.Row, C_IDDOC))
      lTipoLib = Val(Grid.TextMatrix(Grid.Row, C_IDTIPOLIB))
   
      If IdDoc > 0 And Val(Grid.TextMatrix(Grid.Row, C_IDESTADO)) <> ED_ANULADO Then
      
         If Not EstadoValido(Grid.Row) Then
            Exit Sub
         End If
      
         lLstIdDoc(0).IdDoc = IdDoc
         lLstIdDoc(0).IdDocCuota = Val(Grid.TextMatrix(Grid.Row, C_IDDOCCUOTA))
         lLstIdDoc(0).MontoCuota = vFmt(Grid.TextMatrix(Grid.Row, C_MONTOCUOTA))
         lLstIdDoc(0).NumCuotas = vFmt(Grid.TextMatrix(Grid.Row, C_NUMCUOTAS))
      End If
   
   End If
   
   If lLstIdDoc(0).IdDoc > 0 Then
      lRc = vbOK
      Unload Me
   Else
      MsgBeep vbExclamation
   End If
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
      
      If (TipoLib = LIB_OTROS Or TipoLib = LIB_REMU) And Cb_Estado.ListCount > 0 Then    'dejamos sin selección de estado
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

Private Sub Ch_VerCuotas_Click()
   Call EnableFrm(True)

   If Ch_VerCuotas = 0 Then
      Grid.ColWidth(C_CUOTA) = 0
      Grid.ColWidth(C_MONTOCUOTA) = 0
      Grid.TextMatrix(0, C_CUOTA) = ""
      Grid.TextMatrix(0, C_MONTOCUOTA) = ""
   Else
      Grid.ColWidth(C_CUOTA) = 900
      Grid.ColWidth(C_MONTOCUOTA) = 1100
      Grid.TextMatrix(0, C_CUOTA) = "Cuota"
      Grid.TextMatrix(0, C_MONTOCUOTA) = "Monto Cuota"
   End If
  
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
      Cb_TipoLib.ItemData(Cb_TipoLib.NewIndex) = gTipoLibNew(i).id 'i
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
   
   Ch_VerCuotas.Value = 1
   Tx_NumCuotas = 1
   
   StrSort = "Documento.TipoLib, Documento.TipoDoc, Documento.NumDoc, NumCuota "
                           
   lOrdenGr(C_RUT) = "Entidades.RUT, " & StrSort
   lOrdenGr(C_ENTIDAD) = "Entidades.Nombre, " & StrSort
   lOrdenGr(C_TIPOLIB) = StrSort & ", Documento.FEmision, Entidades.Nombre"
   lOrdenGr(C_TIPODOC) = "Documento.TipoDoc, Documento.TipoLib, Documento.NumDoc "
   lOrdenGr(C_NUMDOC) = "Documento.NumDoc, Entidades.Nombre "
   lOrdenGr(C_CHECK) = StrSort
   lOrdenGr(C_FEMISION) = "Documento.FEmision, Entidades.Nombre, " & StrSort
   lOrdenGr(C_FVENC) = "Documento.FVenc, Entidades.Nombre, Documento.FEmision "
   lOrdenGr(C_VALOR) = "Documento.Total, Entidades.Nombre, Documento.FEmision "
   lOrdenGr(C_NUMCUOTAS) = "NumCuotas," & StrSort
   lOrdenGr(C_SALDO) = "Documento.SaldoDoc, Entidades.Nombre, Documento.FEmision "
   lOrdenGr(C_ESTADO) = "Documento.Estado, Entidades.Nombre, Documento.FEmision "
   lOrdenGr(C_DOCASOC) = StrSort
   lOrdenGr(C_DESC) = "Documento.Descrip, Entidades.Nombre, Documento.FEmision "
   lOrdenGr(C_CUOTA) = "NumCuota, Documento.TipoLib, Documento.TipoDoc, Documento.NumDoc "
   lOrdenGr(C_NUMCUOTA) = "NumCuota, Documento.TipoLib, Documento.TipoDoc, Documento.NumDoc "
   lOrdenGr(C_MONTOCUOTA) = "MontoCuota," & StrSort
   
   lOrdenSel = C_TIPOLIB
   
   Select Case lOper
      Case O_VIEW
         Me.Caption = "Listar Documentos"
         Bt_Sel.visible = False
         Bt_NewDoc.visible = False
         Bt_ModDoc.visible = False
         Bt_DelDoc.visible = False
      
      Case O_SELECT
         Me.Caption = "Seleccionar Documento"
         If lShowBtNew = True Then
            Bt_NewDoc.visible = True
         Else
            Bt_NewDoc.visible = False
         End If
         Bt_ModDoc.visible = False
         Bt_DelDoc.visible = False
            
      Case O_EDIT
         Me.Caption = "Listar/Editar Documentos"
         Bt_Sel.visible = False
         
   End Select
   
   Bt_DocCuotas.visible = gFunciones.DocCuotas
   Bt_DocCuotas.Enabled = gFunciones.DocCuotas
            
   Call SetUpGrid
   
   Me.MousePointer = vbHourglass
   
   Call RecalcSaldos(gEmpresa.id, gEmpresa.Ano)
   Call RecalcSaldosFulle(gEmpresa.id, gEmpresa.Ano)
   
   Me.MousePointer = vbDefault

   Call LoadGrid
   
   Call SetupPriv

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
   Dim Col As Integer
   Dim Row As Integer
   Dim i As Integer
   Dim MsgCero As Boolean
   Dim MsgEstado As Boolean
   
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Row < Grid.FixedRows And Col <> C_CHECK Then
      Exit Sub
   End If
   
   If Col <> C_CHECK Then
      
      If lOper = O_SELECT Then
         Call PostClick(Bt_DetDoc)
         
      ElseIf lOper = O_EDIT Then
         Call PostClick(Bt_ModDoc)
         
      ElseIf lOper = O_VIEW Then
         Call PostClick(Bt_DetDoc)
      End If
      
      Exit Sub
      
   End If
   
   'Es C_CHECK
   
   If Row < Grid.FixedRows Then
   
      For i = Grid.FixedRows To Grid.rows - 1
         
         If Grid.TextMatrix(i, C_IDDOC) = "" Then
            Exit For
         End If
         
         If Val(Grid.TextMatrix(i, C_IDESTADO)) <> ED_ANULADO Then
         
            Grid.Row = i
            Grid.Col = C_CHECK
               
            If lTogCheck Then    'desmarcamos todo
               If Grid.CellPicture <> 0 Then
                  Set Grid.CellPicture = LoadPicture()
               End If
            Else                 'marcamos todo
               If Grid.CellPicture = 0 Then
                  
                  If EstadoValido(i, False) Then
                     
                     If vFmt(Grid.TextMatrix(i, C_SALDO)) = 0 And (Val(Grid.TextMatrix(i, C_IDTIPOLIB)) = LIB_COMPRAS Or Val(Grid.TextMatrix(i, C_IDTIPOLIB)) = LIB_VENTAS Or Val(Grid.TextMatrix(i, C_TIPOLIB)) = LIB_RETEN) Then
                        If Not MsgCero Then
                           MsgBox1 "Aención: Los documentos que tienen saldo cero no serán marcados.", vbInformation
                           MsgCero = True
                        End If
                     
                     Else
                        Call FGrSetPicture(Grid, i, C_CHECK, Pc_Check, 0)
                        DoEvents
                     
                     End If
                     
                  ElseIf Not MsgEstado Then
                     MsgBox1 "Atención: Hay documentos que no están en estado " & gEstadoDoc(ED_CENTRALIZADO) & " o " & gEstadoDoc(ED_PAGADO) & " que no serán marcados.", vbInformation
                     MsgEstado = True
                  End If
               End If
            End If
         End If
         
      Next i
      
      lTogCheck = Not lTogCheck
            
   ElseIf Val(Grid.TextMatrix(Row, C_IDDOC)) > 0 And Val(Grid.TextMatrix(Row, C_IDESTADO)) <> ED_ANULADO Then
                  
      If Not EstadoValido(Row) Then
         Exit Sub
      End If
      
      Grid.Row = Row
      Grid.Col = Col
      
      If Grid.CellPicture = 0 And vFmt(Grid.TextMatrix(Row, C_SALDO)) = 0 And (Val(Grid.TextMatrix(Row, C_IDTIPOLIB)) = LIB_COMPRAS Or Val(Grid.TextMatrix(Row, C_IDTIPOLIB)) = LIB_VENTAS Or Val(Grid.TextMatrix(Row, C_TIPOLIB)) = LIB_RETEN) Then
         If MsgBox1("El saldo de este documento es cero." & vbNewLine & vbNewLine & "¿Desea marcarlo de todas maneras?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
         End If
      End If
      
      
      If Grid.CellPicture = 0 Then
         Call FGrSetPicture(Grid, Row, C_CHECK, Pc_Check, 0)
      Else
         Set Grid.CellPicture = LoadPicture()
      End If
      
   End If

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

Private Sub Tx_NumCuotas_Change()
   Call EnableFrm(True)

End Sub

Private Sub Tx_NumCuotas_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_NumCuotas_LostFocus()
   If Val(Tx_NumCuotas) = 0 Then
      Tx_NumCuotas = 1
   End If
   
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
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
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
   ColWi(C_CHECK) = 0
   
   If ColWi(C_CUOTA) > 0 Then
      ColWi(C_DESC) = 0
   End If
   
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
   Grid.ColWidth(C_CHECK) = 0
   Grid.ColWidth(C_VALOR) = 1200
   Grid.ColWidth(C_SALDO) = 1200
   Grid.ColWidth(C_FEMISION) = 800
   Grid.ColWidth(C_NUMCUOTAS) = 0
   Grid.ColWidth(C_IDDOCCUOTA) = 0
   Grid.ColWidth(C_CUOTA) = 700
   Grid.ColWidth(C_NUMCUOTA) = 0
   Grid.ColWidth(C_MONTOCUOTA) = 1200
   Grid.ColWidth(C_FVENC) = 800
   Grid.ColWidth(C_ESTADO) = 900
   Grid.ColWidth(C_IDESTADO) = 0
   Grid.ColWidth(C_DOCASOC) = 1400
   Grid.ColWidth(C_DESC) = 2800
   
         
   Grid.ColAlignment(C_TIPOLIB) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPODOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_CHECK) = flexAlignCenterCenter
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
   Grid.TextMatrix(0, C_TRATAMIENTO) = "Tratamiento"
    
   If lOper = O_SELECT Then
      Grid.ColWidth(C_CHECK) = 300
      Grid.Row = 0
      Grid.Col = C_CHECK
      Set Grid.CellPicture = Pc_HdCheck
      Grid.CellPictureAlignment = flexAlignCenterCenter
   End If
    
   Call FGrSetup(Grid)
   Call FGrTotales(Grid, GridTot)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
End Sub
Private Sub LoadGrid(Optional ByVal IdDoc As Long = 0)
   Dim Q1 As String
   Dim Q2 As String
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
   
   If Where <> "" Then
      Where = " WHERE " & Mid(Where, 6)
   End If

'   TmpTbl = DbGenTmpName2(gDbType, "tdoccuota_")
   TmpTbl = DbGenTmpName2(SQL_ACCESS, "tdoccuota_")   'forzamos para que no le ponga el # al nombre
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)

   Q1 = "SELECT IdDocCuota, DocCuotas.IdDoc, DocCuotas.NumCuota, MontoCuota, FechaExigPago, " & gEmpresa.id & " As IdEmpresa, " & gEmpresa.Ano & " As Ano INTO " & TmpTbl
   Q1 = Q1 & " FROM DocCuotas INNER JOIN vPrimeraDocCuota ON DocCuotas.IdDoc = vPrimeraDocCuota.IdDoc "
   Q1 = Q1 & " AND (DocCuotas.NumCuota BETWEEN vPrimeraDocCuota.NumCuota1 AND vPrimeraDocCuota.NumCuota1 + " & Val(Tx_NumCuotas) - 1 & ") "
   Q1 = Q1 & JoinEmpAno(gDbType, "DocCuotas", "vPrimeraDocCuota")
   Q1 = Q1 & " WHERE DocCuotas.Estado = " & ED_PENDIENTE
   Q1 = Q1 & " AND DocCuotas.IdEmpresa = " & gEmpresa.id     '   & " AND DocCuotas.Ano = " & gEmpresa.Ano   'FCA podría ser de cualñquier año (10 dic 2019)
   Q1 = Q1 & " ORDER BY DocCuotas.IdDoc"
   Call ExecSQL(DbMain, Q1)
   


   Q1 = "SELECT Documento.IdDoc, TipoLib, TipoDoc, NumDoc, NumDocHasta, Documento.IdEntidad, Entidades.Rut, Entidades.Nombre "
   Q1 = Q1 & ", Entidades.NotValidRut, FEmision, FEmisionOri, FVenc, Total, Descrip, Documento.Estado, SaldoDoc, IdDocAsoc "
   'Q1 = Q1 & ", Documento.NumCuotas, IdDocCuota, NumCuota, MontoCuota, FechaExigPago ,0 as tratamiento  "
   Q1 = Q1 & ", Documento.NumCuotas, IdDocCuota, NumCuota, MontoCuota, FechaExigPago , tratamiento  "
   Q1 = Q1 & " FROM (Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "Entidades", True, True) & ")"
   Q1 = Q1 & " LEFT JOIN " & TmpTbl & " ON Documento.IdDoc = " & TmpTbl & ".IdDoc  "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", TmpTbl)
   Q1 = Q1 & Where
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   'Q2 = Replace(Replace(Q1, "Documento", "DocumentoFull"), ",0", ", DocumentoFull.tratamiento")
   
   'Q1 = Q1 & " UNION ALL " & Q2
   If lOrdenGr(lOrdenSel) <> "" Then
      Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel) & ", IdDocCuota "

   Else
      Q1 = Q1 & " ORDER BY Documento.TipoLib, Documento.TipoDoc, Documento.NumDoc, NumCuota "
      
   End If

   Set Rs = OpenRs(DbMain, Q1)

   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   Do While Rs.EOF = False
      
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
      
      If IdDoc > 0 And vFld(Rs("IdDoc")) = IdDoc Then
         Row = i
      End If
      
      Grid.TextMatrix(i, C_TIPOLIB) = Left(ReplaceStr(gTipoLibNew(IIf(vFld(Rs("TipoLib")) = 8, 6, vFld(Rs("TipoLib")))).Nombre, "Libro de ", ""), 9)
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
      Grid.TextMatrix(i, C_NUMCUOTA) = vFld(Rs("NumCuota"))
      Grid.TextMatrix(i, C_MONTOCUOTA) = IIf(vFld(Rs("MontoCuota")) > 0, Format(vFld(Rs("MontoCuota")), NUMFMT), "")

      
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
      Grid.TextMatrix(i, C_ESTADO) = Left(gEstadoDoc(vFld(Rs("Estado"))), 9)
      Grid.TextMatrix(i, C_IDESTADO) = vFld(Rs("Estado"))
      
      If vFld(Rs("IdDocAsoc")) <> 0 Then
         Grid.TextMatrix(i, C_DOCASOC) = GetInfoDoc(vFld(Rs("IdDocAsoc")))
      End If
      
      'FEÑA
      Grid.TextMatrix(i, C_TRATAMIENTO) = ""
      If vFld(Rs("Tratamiento")) > 0 Then
        Grid.TextMatrix(i, C_TRATAMIENTO) = IIf(vFld(Rs("Tratamiento")) = 1, "A", "P")
      End If
      'FIN FEÑA
             
      Rs.MoveNext
      i = i + 1
   Loop

   Call CloseRs(Rs)
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTbl)
   
   GridTot.TextMatrix(0, C_FVENC) = "TOTAL"
   GridTot.TextMatrix(0, C_VALOR) = Format(Total, NUMFMT)
   GridTot.TextMatrix(0, C_SALDO) = Format(TotSaldo, NEGNUMFMT)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
   'Marco la columna Ordenada
   If lOrdenSel <> C_CHECK And lOrdenSel <> C_DOCASOC Then
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

Private Function valida() As Boolean

   valida = False
   
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
   
   valida = True

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
   If lOrdenSel <> C_CHECK And lOrdenSel <> C_DOCASOC Then
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
   Bt_Sel.Enabled = Not bool
'   bt_Print.Enabled = Not bool
'   Bt_Preview.Enabled = Not bool
'   Bt_CopyExcel.Enabled = Not bool
   
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

Private Function SetupPriv()
   
   If lOper = O_EDIT Then
      If Not ChkPriv(PRV_ING_DOCS) Then
         Bt_NewDoc.Enabled = False
         Bt_ModDoc.Caption = "Ver"
         Bt_DelDoc.Enabled = False
      End If
   End If
   
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

