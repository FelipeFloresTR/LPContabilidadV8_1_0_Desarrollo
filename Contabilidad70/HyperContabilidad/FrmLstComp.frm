VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLstComp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Comprobantes"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   Icon            =   "FrmLstComp.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   13320
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5235
      Left            =   0
      TabIndex        =   22
      Top             =   2700
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   9234
      _Version        =   393216
      Rows            =   25
      Cols            =   11
      FixedCols       =   2
      BackColorBkg    =   16777215
   End
   Begin VB.PictureBox Pc_Prt 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   8460
      Picture         =   "FrmLstComp.frx":000C
      ScaleHeight     =   240
      ScaleWidth      =   270
      TabIndex        =   47
      Top             =   60
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox Pc_Check 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   8160
      Picture         =   "FrmLstComp.frx":039C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   46
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame 
      Height          =   555
      Left            =   60
      TabIndex        =   45
      Top             =   0
      Width           =   13215
      Begin VB.CommandButton Bt_traspasar 
         Caption         =   "Copiar a Mes Sig."
         Height          =   315
         Left            =   9480
         TabIndex        =   58
         Top             =   180
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.PictureBox Pc_HdCheck 
         AutoSize        =   -1  'True
         Height          =   210
         Left            =   8820
         Picture         =   "FrmLstComp.frx":0413
         ScaleHeight     =   150
         ScaleWidth      =   150
         TabIndex        =   56
         Top             =   60
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.CommandButton Bt_CambiarEstadoComps 
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
         Picture         =   "FrmLstComp.frx":0778
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Cambiar estado a comprobantes seleccionados"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Auditoria 
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
         Left            =   3660
         Picture         =   "FrmLstComp.frx":0BE1
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Ver Auditoría de Movimientos de Comprobantes"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_ViewCompRes 
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
         Left            =   540
         Picture         =   "FrmLstComp.frx":1038
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ver comprobante seleccionado en forma resumida"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_DelComp 
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
         Left            =   1080
         Picture         =   "FrmLstComp.frx":143D
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Eliminar comprobante seleccionado"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_PrtSelComps 
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
         Left            =   3120
         Picture         =   "FrmLstComp.frx":1839
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Imprimir comprobantes seleccionados"
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
         Left            =   2040
         Picture         =   "FrmLstComp.frx":1CD4
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Left            =   1620
         Picture         =   "FrmLstComp.frx":20C4
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton bt_CopyExcel 
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
         Left            =   5040
         Picture         =   "FrmLstComp.frx":2168
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Copiar Excel"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_DetComp 
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
         Picture         =   "FrmLstComp.frx":25AD
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Detalle comprobante seleccionado"
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
         Left            =   4620
         Picture         =   "FrmLstComp.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Imprimir listado en pantalla"
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
         Left            =   4200
         Picture         =   "FrmLstComp.frx":2ECC
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   11760
         TabIndex        =   37
         Top             =   180
         Width           =   1275
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
         Left            =   6420
         Picture         =   "FrmLstComp.frx":3373
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Calendario"
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
         Left            =   5580
         Picture         =   "FrmLstComp.frx":379C
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Convertir moneda"
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
         Left            =   6000
         Picture         =   "FrmLstComp.frx":3B3A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Calculadora"
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listar por"
      ForeColor       =   &H00FF0000&
      Height          =   1995
      Left            =   60
      TabIndex        =   38
      Top             =   600
      Width           =   13215
      Begin VB.ComboBox Cb_Sucursal 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1440
         Width           =   1635
      End
      Begin VB.CheckBox Ch_DTE 
         Caption         =   "DTE"
         Height          =   315
         Left            =   2520
         TabIndex        =   15
         Top             =   1440
         Width           =   675
      End
      Begin VB.ComboBox Cb_TipoAjuste 
         Height          =   315
         Left            =   9480
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox Tx_NumDoc 
         Height          =   315
         Left            =   1140
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox Ch_Rut 
         Caption         =   "RUT:"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   5880
         TabIndex        =   17
         Top             =   1440
         Width           =   225
      End
      Begin VB.ComboBox Cb_Cuentas 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1020
         Width           =   4215
      End
      Begin VB.TextBox Tx_Rut 
         Height          =   315
         Left            =   6600
         MaxLength       =   12
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox Cb_Nombre 
         Height          =   315
         Left            =   9300
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1440
         Width           =   3735
      End
      Begin VB.ComboBox Cb_Entidad 
         Height          =   315
         Left            =   7920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1440
         Width           =   1395
      End
      Begin VB.ComboBox Cb_TipoLib 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1020
         Width           =   1695
      End
      Begin VB.ComboBox Cb_TipoDoc 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1020
         Width           =   1635
      End
      Begin VB.TextBox Tx_Valor 
         Height          =   315
         Left            =   11760
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1275
      End
      Begin VB.CommandButton Bt_Search 
         Caption         =   "&Listar"
         Height          =   675
         Left            =   11760
         Picture         =   "FrmLstComp.frx":3E9B
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Index           =   1
         Left            =   3840
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Bt_Calendario 
         Height          =   315
         Index           =   1
         Left            =   5220
         Picture         =   "FrmLstComp.frx":42D9
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Tx_IdComp 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1635
      End
      Begin VB.ComboBox Cb_Tipo 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Index           =   0
         Left            =   1140
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox Tx_Glosa 
         Height          =   315
         Left            =   6600
         TabIndex        =   8
         ToolTipText     =   "Para buscar ingrese una palabra de la glosa o parte de ella"
         Top             =   600
         Width           =   4215
      End
      Begin VB.CommandButton Bt_Calendario 
         Height          =   315
         Index           =   0
         Left            =   2520
         Picture         =   "FrmLstComp.frx":434E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Bt_Glosas 
         Height          =   315
         Left            =   10860
         Picture         =   "FrmLstComp.frx":43C3
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Suc.:"
         Height          =   195
         Index           =   12
         Left            =   3300
         TabIndex        =   57
         Top             =   1500
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Ajuste:"
         Height          =   195
         Index           =   11
         Left            =   8520
         TabIndex        =   55
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Doc.:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   54
         Top             =   1500
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "RUT:"
         Height          =   195
         Left            =   6120
         TabIndex        =   53
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   10
         Left            =   5880
         TabIndex        =   52
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Libro:"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   51
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Doc.:"
         Height          =   195
         Index           =   8
         Left            =   3300
         TabIndex        =   50
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor mov.:"
         Height          =   195
         Index           =   6
         Left            =   10920
         TabIndex        =   49
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   5
         Left            =   3300
         TabIndex        =   44
         Top             =   660
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° comp.:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   2
         Left            =   3300
         TabIndex        =   42
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desde:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   3
         Left            =   5880
         TabIndex        =   40
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   195
         Index           =   4
         Left            =   5880
         TabIndex        =   39
         Top             =   660
         Width           =   450
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   0
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   7980
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   8
      FixedCols       =   2
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmLstComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDCOMP = 0
Const C_CORRCOMP = 1
Const C_CHECK = 2
Const C_TIPO = 3
Const C_ESTADO = 4
Const C_FECHA = 5
Const C_DEBE = 6
Const C_GLOSA = 7
Const C_TAJUSTE = 8
Const C_IDTAJUSTE = 9
Const C_USUARIO = 10
Const C_FIMPORT = 11
Const C_DETALLE = 12
Const C_LNGFECHA = 13
Const C_IDTIPO = 14
Const C_IDESTADO = 15
Const C_FMT = 16
'Const NCOLS = C_FMT
'FEÑA
Const C_TIPOLIB = 17
Const NCOLS = C_TIPOLIB

'FIN FEÑA

'2861591
Const TX_ACTFIJO = "AF >>"
'2861591

Dim lOrdenGr(C_FIMPORT) As String
Dim lOrdenSel As Integer    'orden seleccionado o actual

Const M_IDENTIDAD = 1
Const M_RUT = 2
Const M_NOTVALIDRUT = 3

Const F_INICIO = 0
Const F_FIN = 1

Dim lcbNombre As ClsCombo

Dim Oper As Integer

Dim lOrientacion As Integer

Private Sub Bt_Auditoria_Click()
   Dim Frm As FrmAuditoria
   
   Set Frm = New FrmAuditoria
   Frm.Show vbModal
   Set Frm = Nothing
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

Private Sub Bt_Calendario_Click(Index As Integer)
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_Fecha(Index))
   Set Frm = Nothing

End Sub

Private Sub Bt_CambiarEstadoComps_Click()
   Dim i As Integer
   Dim Frm As FrmCambioEstadoComp
   Dim n As Integer
   Dim LstComp As String
   Dim Rc As Integer
   Dim NewEstado As Integer
   Dim nIgual As Integer
   Dim Q1 As String
   
   If MsgBox1("Esta opción permite cambiar el estado de los comprobantes seleccionados de Pendientes a Aprobados y viceversa." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Set Frm = New FrmCambioEstadoComp
   Rc = Frm.FEdit(NewEstado)
   If Rc = vbCancel Then
      Exit Sub
   End If
   
   n = 0
   LstComp = ""
   For i = Grid.FixedRows To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_IDCOMP)) <= 0 Then
         Exit For
      End If
      
      Grid.Row = i
      Grid.Col = C_CHECK
      
      If Grid.CellPicture <> 0 Then
         If Grid.TextMatrix(i, C_IDESTADO) <> EC_APROBADO And Grid.TextMatrix(i, C_IDESTADO) <> EC_PENDIENTE Then
            MsgBox1 "Para hacer cambio de estado sólo puede seleccionar comprobantes en estado Pendiente o Aprobado.", vbExclamation
            Exit Sub
         End If
         If Grid.TextMatrix(i, C_IDESTADO) = NewEstado Then
            nIgual = nIgual + 1
         End If
         LstComp = LstComp & ", " & Grid.TextMatrix(i, C_IDCOMP)
         n = n + 1
      End If
   Next i
   
   LstComp = Mid(LstComp, 2)

   Set Frm = Nothing
   
   If n = 0 Then
      MsgBox1 "No hay comprobantes seleccionados.", vbExclamation
      Exit Sub
   ElseIf n = nIgual Then
      MsgBox1 "Todos los comprobantes seleccionados ya tienen el estado " & UCase(gEstadoComp(NewEstado)), vbExclamation
      Exit Sub
   ElseIf nIgual > 0 Then
      If MsgBox1("Algunos de los comprobantes seleccionados ya tienen el estado " & UCase(gEstadoComp(NewEstado)) & vbCrLf & vbCrLf & "Se realizará la operación para el resto de los comprobantes" & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
   ElseIf MsgBox1("Se cambiará el estado de todos los comprobantes seleccionados a " & UCase(gEstadoComp(NewEstado)) & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
      Exit Sub
   End If
      
   Me.MousePointer = vbHourglass
   
   Q1 = "UPDATE Comprobante SET Estado = " & NewEstado & " WHERE IdComp IN (" & LstComp & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   '3376884
   Call SeguimientoMovComprobante(0, gEmpresa.id, gEmpresa.Ano, "FrmLstComp.Bt_CambiarEstadoComps_Click", Q1, 1, "WHERE IdComp IN (" & LstComp & ") AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano, 1, 2)
   'Fin 3376884
   
   'actualizamos el estado en la lista
   For i = Grid.FixedRows To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_IDCOMP)) <= 0 Then
         Exit For
      End If
      
      Grid.Row = i
      Grid.Col = C_CHECK
      
      If Grid.CellPicture <> 0 Then
         Grid.TextMatrix(i, C_IDESTADO) = NewEstado
         Grid.TextMatrix(i, C_ESTADO) = gEstadoComp(NewEstado)
      End If
   Next i
   
   Me.MousePointer = vbDefault

End Sub

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_ConvMoneda_Click()
   Dim Frm As FrmConverMoneda
   Dim Col As Integer
   Dim Row As Integer
   Dim Valor As Double
   
   Col = Grid.Col
   Row = Grid.Row
   
   'If Col <> C_DEBE Then
   '   MsgBox1 "Esta opción se utiliza sólo en la columna Valor.", vbExclamation
   '   Exit Sub
   'End If
   
   'If Trim(Grid.TextMatrix(Row, C_IDCOMP)) = "" Then
   '   Exit Sub
   'End If
   
   If Col = C_DEBE Then
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
   
      '2861591
     If MsgBox1("¿Desea Exportar todos los comprobantes con su detalle?." & vbNewLine & "Si opcion es NO, se copiaran comprobantes sin detalle", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
         Export_DetComprobante
     Else
     Call FGr2Clip(Grid, "Número comprobante: " & Tx_IdComp & " Tipo: " & Cb_Tipo & " Estado: " & Cb_Estado & " Fecha Inicio: " & Tx_Fecha(0) & " Fecha Término: " & Tx_Fecha(1))
    End If
    '2861591
   
   'Call FGr2Clip(Grid, "Número comprobante: " & Tx_IdComp & " Tipo: " & Cb_Tipo & " Estado: " & Cb_Estado & " Fecha Inicio: " & Tx_Fecha(0) & " Fecha Término: " & Tx_Fecha(1))
End Sub

Private Sub Bt_DelComp_Click()
   Dim IdComp As Long, PcName As String
   Dim CorrComp As Long
   Dim FechaComp As Long
   Dim EstadoComp As Integer
   Dim TipoComp As Integer
   Dim TipoAjuste As Integer
   Dim TipoLibro As Integer
   Dim DocFull As Boolean
   
   DocFull = False
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   IdComp = Val(Grid.TextMatrix(Grid.Row, C_IDCOMP))
   
   If IdComp = 0 Then
      Exit Sub
   End If

   CorrComp = vFmt(Grid.TextMatrix(Grid.Row, C_CORRCOMP))
   FechaComp = vFmt(Grid.TextMatrix(Grid.Row, C_LNGFECHA))
   EstadoComp = vFmt(Grid.TextMatrix(Grid.Row, C_IDESTADO))
   TipoComp = vFmt(Grid.TextMatrix(Grid.Row, C_IDTIPO))
   TipoAjuste = vFmt(Grid.TextMatrix(Grid.Row, C_IDTAJUSTE))
   'FEÑA
   TipoLibro = vFmt(Grid.TextMatrix(Grid.Row, C_TIPOLIB))
   If TipoLibro <> 0 Then
    DocFull = True
   End If
   'FIN FEÑA
   
   PcName = IsLockedAction(DbMain, LK_COMPROBANTE, IdComp)
   If PcName <> "" Then
      MsgBox1 "Este comprobante se está editando en el equipo '" & PcName & "'. No puede ser eliminado.", vbInformation
      Exit Sub
   End If

   If DeleteComprobante(IdComp, , DocFull) = True Then

      MousePointer = vbHourglass
      Call LoadAll
      MousePointer = vbDefault
      
      Call AddLogComprobantes(IdComp, gUsuario.IdUsuario, O_DELETE, Now, EC_ELIMINADO, CorrComp, FechaComp, TipoComp, EC_ELIMINADO, TipoAjuste)

      MsgBox1 "El comprobante ha sido eliminado.", vbInformation + vbOKOnly
   
   End If

End Sub

Private Sub Bt_DetComp_Click()

   Call ViewDetComp(Grid.Row, Grid.Col)
   
End Sub

Private Sub Bt_Orden_Click()
   
   Call OrdenaPorCol(Grid.Col)
         
End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   
   If Bt_Search.Enabled = True Then
      MsgBox1 "Presione el botón Listar antes de seleccionar al vista previa.", vbExclamation
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

   Call ResetPrtBas(gPrtReportes)

End Sub

Private Sub Bt_Glosas_Click()
   Dim Frm As FrmGlosas
      
   Set Frm = New FrmGlosas

   Tx_Glosa = FrmGlosas.FSelect(Tx_Glosa)
   
   Set Frm = Nothing
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
   
   Call ResetPrtBas(gPrtReportes)

   
End Sub

Private Sub Bt_PrtSelComps_Click()
   Dim i As Integer
   Dim Frm As FrmComprobante
   Dim n As Integer
   Dim Msg As Boolean

   
   Set Frm = New FrmComprobante
   n = 0
   For i = Grid.FixedRows To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_IDCOMP)) <= 0 Then
         Exit For
      End If
      
      Grid.Row = i
      Grid.Col = C_CHECK
      
      If Grid.CellPicture <> 0 Then
         If Not Msg Then
            If MsgBox1("Se imprimirán todos los comprobantes seleccionados." & vbNewLine & vbNewLine & "¿Desea continuar?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
               Exit Sub
            End If
            Msg = True
         End If
         Call Frm.FPrtComp(Val(Grid.TextMatrix(i, C_IDCOMP)))
         n = n + 1
      End If
   Next i

   Set Frm = Nothing
   
   If n = 0 Then
      MsgBox1 "No hay comprobantes seleccionados.", vbExclamation
   End If
   
End Sub

Private Sub Bt_Search_Click()

   If valida() = False Then
      Exit Sub
   End If
   
   MousePointer = vbHourglass
   
   If ExitDemo() Then
      Unload Me
   End If
   
   Call LoadAll
   MousePointer = vbDefault
   
End Sub

Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid)
   
   Set Frm = Nothing
   
End Sub


Private Sub Bt_ViewCompRes_Click()
   Dim IdComp As Long
   Dim Frm As FrmComprobante
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
      
   IdComp = Val(Grid.TextMatrix(Row, C_IDCOMP))

   If IdComp <> 0 Then
      Set Frm = New FrmComprobante
      Call Frm.FViewResumido(IdComp)
      Set Frm = Nothing
      
      Grid.Row = Row
      Grid.Col = C_TIPO
      Grid.RowSel = Grid.Row
      Grid.ColSel = C_GLOSA
   End If

End Sub

Private Sub Cb_Cuentas_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_Estado_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_Sucursal_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_Tipo_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_TipoAjuste_Click()
   Call EnableFrm(True)

End Sub

Private Sub Cb_TipoDoc_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_DTE_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_Rut_Click()
   Call EnableFrm(True)

End Sub

Private Sub Form_Load()
   Dim MesActual As Integer
   Dim F1 As Long
   Dim F2 As Long
   
   lOrientacion = ORIENT_VER
      
   'Lleno el arreglo de orden de columnas
   lOrdenGr(C_CORRCOMP) = "Comprobante.Correlativo"
   lOrdenGr(C_TIPO) = "Comprobante.Tipo, Comprobante.Fecha, Comprobante.Correlativo"
   lOrdenGr(C_ESTADO) = "Comprobante.Estado, Comprobante.Fecha, Comprobante.Correlativo"
   lOrdenGr(C_FECHA) = "Comprobante.Fecha, Comprobante.Correlativo"
   lOrdenGr(C_DEBE) = "Comprobante.TotalDebe, Comprobante.Fecha, Comprobante.Correlativo"
   lOrdenGr(C_GLOSA) = "Comprobante.Glosa, Comprobante.Fecha, Comprobante.Correlativo"
   lOrdenGr(C_USUARIO) = "Usuarios.Usuario, Comprobante.Fecha, Comprobante.Correlativo"
   lOrdenGr(C_FIMPORT) = "Comprobante.FechaImport, Comprobante.Fecha, Comprobante.Correlativo"
   lOrdenGr(C_TAJUSTE) = "Comprobante.TipoAjuste, Comprobante.Fecha, Comprobante.Correlativo"
   
   lOrdenSel = C_FECHA
   
   MesActual = GetMesActual()
   If MesActual = 0 Then
      MesActual = GetUltimoMesConComps()
   End If
   Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, MesActual, 1), F1, F2)
   Call SetTxDate(Tx_Fecha(F_INICIO), F1)
   Call SetTxDate(Tx_Fecha(F_FIN), F2)
   
   Ch_Rut = 1

   Call FillCb
   Call SetUpGrid
   
   Bt_ViewCompRes.visible = gFunciones.ComprobanteResumido
   
   Call LoadAll
   
   Call SetupPriv
   
End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
    
   Grid.ColWidth(C_IDCOMP) = 0
   Grid.ColWidth(C_CORRCOMP) = 800
   If Oper = O_PRINT Then
      Grid.ColWidth(C_CHECK) = 300
   Else
      Grid.ColWidth(C_CHECK) = 0
   End If
   Grid.ColWidth(C_TIPO) = 800
   Grid.ColWidth(C_ESTADO) = 900
   Grid.ColWidth(C_FECHA) = FW_FECHA
   Grid.ColWidth(C_DEBE) = 1300
 '  Grid.ColWidth(C_HABER) = 1300
   Grid.ColWidth(C_FIMPORT) = FW_FECHA
   Grid.ColWidth(C_DETALLE) = 250
   Grid.ColWidth(C_TAJUSTE) = 400
   Grid.ColWidth(C_IDTAJUSTE) = 0
   Grid.ColWidth(C_GLOSA) = 5030 - Grid.ColWidth(C_CHECK)
   Grid.ColWidth(C_USUARIO) = 1300
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_LNGFECHA) = 0
   Grid.ColWidth(C_IDTIPO) = 0
   Grid.ColWidth(C_IDESTADO) = 0
      
   Grid.ColAlignment(C_IDCOMP) = flexAlignRightCenter
   Grid.ColAlignment(C_CORRCOMP) = flexAlignRightCenter
   Grid.ColAlignment(C_TIPO) = flexAlignLeftCenter
   Grid.ColAlignment(C_ESTADO) = flexAlignLeftCenter
   Grid.ColAlignment(C_FECHA) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_FIMPORT) = flexAlignRightCenter
 '  Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   Grid.ColAlignment(C_GLOSA) = flexAlignLeftCenter
   Grid.ColAlignment(C_TAJUSTE) = flexAlignCenterCenter
   
   Grid.TextMatrix(0, C_CORRCOMP) = "N° Comp."
   Grid.TextMatrix(0, C_TIPO) = "Tipo"
   Grid.TextMatrix(0, C_ESTADO) = "Estado"
   Grid.TextMatrix(0, C_FECHA) = "Fecha"
   Grid.TextMatrix(0, C_DEBE) = "Valor"
 '  Grid.TextMatrix(0, C_HABER) = "Haber"
   Grid.TextMatrix(0, C_GLOSA) = "Glosa"
   Grid.TextMatrix(0, C_TAJUSTE) = "Ajus"
   Grid.TextMatrix(0, C_USUARIO) = "Usuario"
   Grid.TextMatrix(0, C_FIMPORT) = "Importado"
   Grid.TextMatrix(0, C_FMT) = "       .FMT"
   
   Grid.Row = 0
   Grid.Col = C_CHECK
   'Set Grid.CellPicture = Pc_Prt
   Set Grid.CellPicture = Pc_HdCheck
   Grid.CellPictureAlignment = flexAlignCenterCenter
   
   
   
   Grid.Row = 0
   Grid.Col = C_DETALLE
   Set Grid.CellPicture = FrmMain.Pc_Lupa

   GridTot.Cols = Grid.Cols

   Call FGrSetup(Grid)

   Call FGrVRows(Grid)
   
   GridTot.TextMatrix(0, C_FECHA) = "TOTAL:"

   Call FGrTotales(Grid, GridTot)

End Sub
Private Sub LoadAll(Optional ByVal IdComp As Long = 0)
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Wh As String
   Dim JoinStr As String
   Dim UsrJoin As String

   Grid.Redraw = False
   
   If IdComp > 0 Then
      Wh = " WHERE IdComp = " & IdComp
   Else
      Wh = CreateWhere(JoinStr)
   End If
      
   Q1 = "SELECT DISTINCT Comprobante.IdComp, Comprobante.Correlativo, Comprobante.Fecha, "
   Q1 = Q1 & " Comprobante.Tipo, Comprobante.Estado, Comprobante.Glosa, "
   Q1 = Q1 & " Comprobante.TotalDebe, Comprobante.TotalHaber, Usuarios.Usuario, Comprobante.FechaImport, Comprobante.TipoAjuste ,0 as TipoLib"
   
   UsrJoin = " Comprobante LEFT JOIN Usuarios ON Comprobante.IdUsuario = Usuarios.IdUsuario "
   
   If JoinStr = "" Then
      Q1 = Q1 & " FROM " & UsrJoin
   Else
      Q1 = Q1 & " FROM (((( " & UsrJoin & ")" & JoinStr
   End If
   
   Q1 = Q1 & Wh
   
   If Wh <> "" Then
      Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Else
      Q1 = Q1 & " WHERE Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   End If
   

    'Q2 = Replace(Replace(Q1, "Comprobante", "ComprobanteFull"), ",0", ",8")
    'Q1 = Q1 & " UNION ALL " & Q2

   
   Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel)
   Set Rs = OpenRs(DbMain, Q1)
   
   If IdComp <= 0 Then
      Grid.rows = Grid.FixedRows
      i = Grid.FixedRows
   Else
      i = Grid.Row
   End If
   
   Do While Rs.EOF = False
   
      If IdComp <= 0 Then
         Grid.rows = i + 1
      End If
      
      Grid.TextMatrix(i, C_IDCOMP) = vFld(Rs("IdComp"))
      Grid.TextMatrix(i, C_CORRCOMP) = vFld(Rs("Correlativo"))
      Grid.TextMatrix(i, C_TIPO) = gTipoComp(vFld(Rs("Tipo")))
      Grid.TextMatrix(i, C_IDTIPO) = vFld(Rs("Tipo"))
      
      If vFld(Rs("Estado")) <= UBound(gEstadoComp) Then
         Grid.TextMatrix(i, C_ESTADO) = gEstadoComp(vFld(Rs("Estado")))
         Grid.TextMatrix(i, C_IDESTADO) = vFld(Rs("Estado"))
      Else
         Grid.TextMatrix(i, C_ESTADO) = gEstadoComp(EC_PENDIENTE)
         Grid.TextMatrix(i, C_IDESTADO) = EC_PENDIENTE
      End If
      
      Grid.TextMatrix(i, C_FECHA) = FmtDate(vFld(Rs("Fecha")))
      Grid.TextMatrix(i, C_LNGFECHA) = vFld(Rs("Fecha"))
      Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("TotalDebe")), NUMFMT)
    ' Grid.TextMatrix(i, C_HABER) = Format(vFld(Rs("TotalHaber")), NUMFMT)
      Grid.TextMatrix(i, C_GLOSA) = vFld(Rs("Glosa"), True)
      Grid.TextMatrix(i, C_TAJUSTE) = Left(gTipoAjuste(vFld(Rs("TipoAjuste"))), 1)
      Grid.TextMatrix(i, C_IDTAJUSTE) = vFld(Rs("TipoAjuste"))
      Grid.TextMatrix(i, C_USUARIO) = vFld(Rs("Usuario"), True)
      Grid.TextMatrix(i, C_FIMPORT) = FmtDate(vFld(Rs("FechaImport")))
      Grid.TextMatrix(i, C_TIPOLIB) = vFld(Rs("TipoLib"))
      
      'si cambia mes, insertamos una línea en la impresión
      If vFld(Rs("IdComp")) > 0 And i > Grid.FixedRows And month(vFld(Rs("Fecha"))) <> month(VFmtDate(Grid.TextMatrix(i - 1, C_FECHA))) Then
         Grid.TextMatrix(i, C_FMT) = "L"
      End If
            
      Grid.TextMatrix(i, C_DETALLE) = ">>"
      
      'Call FGrSetPicture(Grid, i, C_DETALLE, FrmMain.Pc_Flecha, vbButtonFace)
      
      Rs.MoveNext

      i = i + 1
      
   Loop
   
   Call CloseRs(Rs)
   
   If IdComp <= 0 Then
      Call FGrVRows(Grid)
      Grid.rows = Grid.rows + 1
      
      Grid.TopRow = Grid.FixedRows
      
      'Marco la columna Ordenada
      Grid.TopRow = Grid.FixedRows
      
      Grid.Row = 0
      Grid.Col = lOrdenSel
      Set Grid.CellPicture = FrmMain.Pc_Flecha
   End If

   Grid.Col = C_CORRCOMP
   Grid.RowSel = Grid.Row
   Grid.ColSel = Grid.Col

   Grid.Redraw = True

   Call CalcTot
   
   Call EnableFrm(False)
End Sub

Private Sub OrdenaPorCol(ByVal Col As Integer)
   
   If Col = C_CHECK Or Col = C_DETALLE Then
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   'Desmarco  columna Ordenada
   Grid.Row = 0
   Grid.Col = lOrdenSel
   Set Grid.CellPicture = LoadPicture()
   
   lOrdenSel = Col
   
   Call LoadAll
      
   Me.MousePointer = vbDefault
      
End Sub
Private Sub UpdateComp()

   If Grid.Row <= 0 Then
      Exit Sub
   End If
   
   'Call LoadAll(Val(Grid.TextMatrix(Grid.Row, C_IDCOMP)))
   Call LoadAll
   
End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - GridTot.Height - 500
   GridTot.Top = Grid.Top + Grid.Height + 30
   'Grid.Width = Me.Width - 230
   GridTot.Width = Grid.Width - 230
   
   Call FGrVRows(Grid)

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

Private Sub Grid_DblClick()
   Dim Col As Integer
   Dim Row As Integer
   Dim i As Integer
         
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Row < Grid.FixedRows And Col = C_CHECK Then    'marcamos todos los comprobantes para imprimir
   
      Grid.Redraw = False
      
      For i = Grid.FixedRows To Grid.rows - 1
         If Grid.TextMatrix(i, C_IDCOMP) = "" Then
            Exit For
         End If
         
         Grid.Row = i
         Grid.Col = C_CHECK
         
         If Grid.CellPicture = 0 Then
            Call FGrSetPicture(Grid, i, C_CHECK, Pc_Check, 0)
         End If
      Next i
      
      Grid.Redraw = True
      
   ElseIf Col = C_CHECK Then
   
      If Val(Grid.TextMatrix(Row, C_IDCOMP)) <> 0 Then
         If Grid.CellPicture = 0 Then
            Call FGrSetPicture(Grid, Row, Col, Pc_Check, 0)
             '2861591
                Bt_traspasar.visible = True
            
            '2861591
         Else
            Set Grid.CellPicture = LoadPicture()
         End If
      End If

   Else
      Call ViewDetComp(Row, Col)
      
   End If
   
End Sub
Private Sub ViewDetComp(ByVal Row As Integer, ByVal Col As Integer)
   Dim IdComp As Long
   Dim Frm As FrmComprobante

   If Row < Grid.FixedRows Then
      Exit Sub
   End If
      
   IdComp = Val(Grid.TextMatrix(Row, C_IDCOMP))
   CodTipoLib = Val(Grid.TextMatrix(Row, C_TIPOLIB))

   If IdComp <> 0 Then
      Set Frm = New FrmComprobante
      Call Frm.FEdit(IdComp, False)
      Set Frm = Nothing
      Call UpdateComp
      Grid.Row = Row
      Grid.Col = C_TIPO
      Grid.RowSel = Grid.Row
      Grid.ColSel = C_GLOSA
   End If
            
End Sub
Public Sub FView()
   Oper = O_VIEW
   
   Me.Show vbModal
End Sub
Public Sub FEdit()
   Oper = O_EDIT      'permite cambia estado comprobante (anular o aprobar)
   
   Me.Show vbModal
End Sub

Public Sub FPrint()
   Oper = O_PRINT     'impresión masiva de comprobantes, mediante selección en grilla con un check
   
   Me.Show vbModal
End Sub

Private Function CreateWhere(JoinStr As String) As String
   Dim Wh As String
   Dim F1 As Long, F2 As Long
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim Idx As Integer
   Dim CodCuenta As String
   Dim NotValidRut As Boolean

   Wh = ""
   
   If Val(Tx_IdComp) <> 0 Then
      Wh = Wh & " AND Comprobante.Correlativo=" & Val(Tx_IdComp)
   End If
   
   If Cb_Estado.ListIndex > 0 Then
      Wh = Wh & " AND Comprobante.Estado=" & ItemData(Cb_Estado)
   End If
   
   If ItemData(Cb_TipoAjuste) > 0 Then
      If ItemData(Cb_TipoAjuste) = TAJUSTE_FINANCIERO Then
         Wh = Wh & " AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
      Else
         Wh = Wh & " AND Comprobante.TipoAjuste=" & ItemData(Cb_TipoAjuste)
      End If
   End If
   
   If Cb_Tipo.ListIndex > 0 Then
      Wh = Wh & " AND Comprobante.Tipo=" & ItemData(Cb_Tipo)
   End If
   
   F1 = GetTxDate(Tx_Fecha(0))
   F2 = GetTxDate(Tx_Fecha(1))
   
   If F1 <> 0 And F2 <> 0 Then
      Wh = Wh & " AND (Comprobante.Fecha BETWEEN " & F1 & " AND " & F2 & ")"
   End If
      
   If Trim(Tx_Glosa) <> "" Then
      Wh = Wh & " AND " & GenLike(DbMain, Tx_Glosa, "Comprobante.Glosa", 3)
   End If
   
   If vFmt(Tx_Valor) <> 0 Or Trim(Tx_NumDoc) <> "" Or Cb_TipoLib.ListIndex > 0 Or Cb_TipoDoc.ListIndex > 0 Or Trim(Tx_Rut) <> "" Or ItemData(Cb_Cuentas) <> -1 Or ItemData(Cb_Sucursal) > 0 Or Ch_DTE <> 0 Then
      JoinStr = " INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp )"
      JoinStr = JoinStr & " LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc) "
      
      If vFmt(Tx_Valor) <> 0 Then
         Wh = Wh & " AND (MovComprobante.Debe = " & vFmt(Tx_Valor) & " OR MovComprobante.Haber = " & vFmt(Tx_Valor) & ")"
      End If
      
      If Trim(Tx_Rut) <> "" Then
         IdEnt = GetIdEntidad(Trim(Tx_Rut), NombEnt, NotValidRut)
         If IdEnt > 0 Then
            Wh = Wh & " AND Documento.IdEntidad = " & IdEnt
         Else
            Tx_Rut = ""
            Cb_Entidad.ListIndex = 0
            Cb_Nombre.ListIndex = 0
         End If
      
      End If
      
      If ItemData(Cb_TipoLib) > 0 Then
         Wh = Wh & " AND Documento.TipoLib = " & ItemData(Cb_TipoLib)
      End If
      
      If ItemData(Cb_TipoDoc) > 0 Then
         Wh = Wh & " AND Documento.TipoDoc = " & ItemData(Cb_TipoDoc)
      End If
      
      If Trim(Tx_NumDoc) <> "" Then
         Wh = Wh & " AND Documento.NumDoc = '" & ParaSQL(Tx_NumDoc) & "'"
      End If
   
      If Trim(Ch_DTE) <> 0 Then
         Wh = Wh & " AND Documento.DTE <> 0 "
      End If
   
      If ItemData(Cb_Sucursal) > 0 Then
         Wh = Wh & " AND Documento.idSucursal = " & CbItemData(Cb_Sucursal)
      End If
   
      If ItemData(Cb_Cuentas) <> -1 Then
      
         Idx = InStr(Cb_Cuentas, " ")
         
         If Idx > 0 Then
            
            CodCuenta = Left(Cb_Cuentas, Idx - 1)
            Wh = Wh & " AND " & GenWhereCuentas(CodCuenta)
                     
            JoinStr = JoinStr & " LEFT JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
         End If
      
      End If
      
   End If
   

   If JoinStr <> "" Then
      JoinStr = JoinStr & ")"
   End If

   If Wh <> "" Then
      Wh = " WHERE " & Mid(Wh, 5)
   End If
   
   CreateWhere = Wh
   
End Function

Private Sub Tx_Fecha_Change(Index As Integer)
   Call EnableFrm(True)

End Sub

Private Sub Tx_Fecha_GotFocus(Index As Integer)
   Call DtGotFocus(Tx_Fecha(Index))
End Sub

Private Sub Tx_Fecha_KeyPress(Index As Integer, KeyAscii As Integer)
   Call KeyDate(KeyAscii)
End Sub

Private Sub Tx_Fecha_LostFocus(Index As Integer)

   If Trim$(Tx_Fecha(Index)) = "" Then
      Exit Sub
   End If
   Call DtLostFocus(Tx_Fecha(Index))
      
End Sub
Private Sub FillCb()
   Dim i As Integer
   Dim Q1 As String
      
   Call AddItem(Cb_Tipo, "(todos)", -1)
   For i = 1 To N_TIPOCOMP
      Call AddItem(Cb_Tipo, gTipoComp(i), i)
   Next i
   Cb_Tipo.ListIndex = 0
   
   Call AddItem(Cb_Estado, "(todos)", -1)
   For i = 1 To UBound(gEstadoComp)        '- 1    'para no incluir estado erróneo
      Call AddItem(Cb_Estado, gEstadoComp(i), i)
   Next i
   Cb_Estado.ListIndex = 0
   
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_FINANCIERO), TAJUSTE_FINANCIERO)
   Call AddItem(Cb_TipoAjuste, gTipoAjuste(TAJUSTE_TRIBUTARIO), TAJUSTE_TRIBUTARIO)
   Call CbSelItem(Cb_TipoAjuste, TAJUSTE_FINANCIERO)
   
   Call AddItem(Cb_TipoLib, "", 0)
   For i = 1 To UBound(gTipoLibNew)
      If gTipoLibNew(i).Nombre = "" Then
         Exit For
      End If
      'Call AddItem(Cb_TipoLib, gTipoLib(i), i)
      Call AddItem(Cb_TipoLib, gTipoLibNew(i).Nombre, gTipoLibNew(i).id)
   Next i
   
   Set lcbNombre = New ClsCombo
   Call lcbNombre.SetControl(Cb_Nombre)
   
   Call AddItem(Cb_Entidad, "", -1)
   For i = ENT_CLIENTE To ENT_OTRO
      Call AddItem(Cb_Entidad, gClasifEnt(i), i)
      
   Next i
   Cb_Entidad.ListIndex = 0     'para no seleccionar ninguno al partir
   
   Call FillCbCuentas(Cb_Cuentas)
   
   Call AddItem(Cb_Sucursal, " ", 0)
   Q1 = "SELECT Descripcion, IdSucursal FROM Sucursales WHERE IdEmpresa = " & gEmpresa.id & " ORDER BY Descripcion"
   Call FillCombo(Cb_Sucursal, DbMain, Q1, -1)


End Sub
Private Function valida() As Boolean
   Dim F1 As Long, F2 As Long
   
   valida = False
   
   F1 = GetDate(Tx_Fecha(0))
   F2 = GetDate(Tx_Fecha(1))
   
   If F1 = 0 And F2 <> 0 Then
      MsgBox1 "Debe ingresar el primer rango de fecha.", vbExclamation + vbOKOnly
      Tx_Fecha(0).SetFocus
      Exit Function
   End If
   
   If F1 <> 0 And F2 = 0 Then
      MsgBox1 "Debe ingresar el segundo rango de fecha.", vbExclamation + vbOKOnly
      Tx_Fecha(1).SetFocus
      Exit Function
   End If
   
   If F1 > F2 Then
      MsgBox1 "Fecha de inicio mayor que la de término.", vbExclamation + vbOKOnly
      Tx_Fecha(1).SetFocus
      Exit Function
   End If
   
   If Trim(Tx_Rut) <> "" And Val(lcbNombre.Matrix(M_IDENTIDAD)) = 0 Then
      MsgBox1 "El RUT ingresado no es válido o no ha sido ingresado al sistema.", vbExclamation + vbOKOnly
      Tx_Rut.SetFocus
      Exit Function
   End If
   
   valida = True
End Function
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(0) As String
   Dim Encabezados(2) As String
   
   Printer.Orientation = lOrientacion
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Caption
   gPrtReportes.Titulos = Titulos
      
   Encabezados(0) = "Tipo:" & vbTab & Cb_Tipo
   Encabezados(1) = "Estado:" & vbTab & Cb_Estado
   If Trim(Tx_Fecha(0)) <> "" Then
      Encabezados(2) = "Fecha: " & vbTab & Tx_Fecha(0) & " - " & Tx_Fecha(1)
   End If
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
               
   ColWi(C_DETALLE) = 0
   
   Total(C_CORRCOMP) = "Total"
   Total(C_DEBE) = GridTot.TextMatrix(0, C_DEBE)
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_IDCOMP
   gPrtReportes.NTotLines = 1
   

End Sub
Private Sub CalcTot()
   Dim TotDebe As Double
   Dim i As Integer
   
   TotDebe = 0
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.RowHeight(i) > 0 Then     ' no está borrado
         TotDebe = TotDebe + vFmt(Grid.TextMatrix(i, C_DEBE))
      End If
   Next i
         
   GridTot.TextMatrix(0, C_DEBE) = Format(TotDebe, BL_NUMFMT)
   
End Sub
  
Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol
End Sub
Private Sub Cb_TipoLib_Click()
   Dim Q1 As String
   Dim i As Integer
   Dim TipoLib As Integer
   
   Cb_TipoDoc.Clear
   
   TipoLib = ItemData(Cb_TipoLib)
   
   If TipoLib > 0 Then
   
      Call FillTipoDoc(Cb_TipoDoc, TipoLib, True, True)
      Cb_TipoDoc.ListIndex = 0
         
   End If
   
   Call EnableFrm(True)
   
End Sub


Private Sub Tx_Glosa_Change()
   Call EnableFrm(True)

End Sub

Private Sub Tx_IdComp_Change()
   Call EnableFrm(True)

End Sub

Private Sub Tx_NumDoc_Change()
   Call EnableFrm(True)

End Sub

Private Sub Tx_Rut_Change()
   Call EnableFrm(True)
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


Private Sub Tx_Valor_Change()
   Call EnableFrm(True)

End Sub

Private Sub Tx_Valor_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub

Private Sub Tx_Valor_LostFocus()
   Tx_Valor = Format(vFmt(Tx_Valor), BL_NUMFMT)
End Sub
      
Private Sub cb_Nombre_Click()
   
   If lcbNombre.ListIndex >= 0 Then
      Tx_Rut = FmtCID(lcbNombre.Matrix(M_RUT), Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0)
      Ch_Rut = IIf(Val(lcbNombre.Matrix(M_NOTVALIDRUT)) = 0, 1, 0)
   End If
   
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

Private Sub SelCbEntidad(Clasif As Integer)
   Dim Q1 As String
   
   lcbNombre.Clear
   If Clasif >= 0 Then
      Q1 = "SELECT Nombre, idEntidad, Rut, NotValidRut FROM Entidades"
      Q1 = Q1 & " WHERE Clasif" & Clasif & "=" & CON_CLASIF
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      Q1 = Q1 & " ORDER BY Nombre "
      Call lcbNombre.FillCombo(DbMain, Q1, -1)
   End If
   
End Sub
Private Sub EnableFrm(bool As Boolean)

   Bt_Search.Enabled = bool
'   bt_Print.Enabled = Not bool
'   Bt_Preview.Enabled = Not bool
'   Bt_CopyExcel.Enabled = Not bool
   
End Sub

Private Sub SetupPriv()
   
   If Not ChkPriv(PRV_ADM_COMP) Then
      Bt_DelComp.Enabled = False
   End If
      

End Sub

Private Sub Tx_RUT_Validate(Cancel As Boolean)
   
   If Tx_Rut = "" Then
      Exit Sub
   End If
   
   If Not MsgValidCID(Tx_Rut, Ch_Rut <> 0) Then
      Cancel = True
      Exit Sub
   End If
   
End Sub

'2861591
Private Sub Bt_traspasar_Click()
   Dim i As Integer
   Dim n As Integer
   Dim ERR As Integer
   Dim LstComp As String
   Dim lidComp As String
   Dim idMov As String
   Dim Rc As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Q2 As String
   Dim Rs2 As Recordset
   Dim Q3 As String
   Dim Rs3 As Recordset
   Dim FldArray(16) As AdvTbAddNew_t
   Dim FldArrayMovCom(17) As AdvTbAddNew_t
   Dim lTblComprobante As String
   Dim lTblMovComprobante As String
   Dim FNameLogImp As String
   Dim sWhere As String, WhConWhere As String, WhConAnd As String, WhTAjuste As String
   Dim MesActual As Integer
   Dim lCorrelativo As Long

   FNameLogImp = gImportPath & "\Log\TraspasoComprobante-" & Format(Now, "yyyymmdd") & ".log"

   lTblComprobante = "Comprobante"
   lTblMovComprobante = "MovComprobante"
   ERR = 0
If MsgBox1("¿Esta seguro que desea traspasar los comprobantes seleccionados al siguiente mes?." & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub

Else
  For i = Grid.FixedRows To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_IDCOMP)) <= 0 Then
         Exit For
      End If

      Grid.Row = i
      Grid.Col = C_CHECK

      If Grid.CellPicture <> 0 Then

         LstComp = LstComp & ", " & Grid.TextMatrix(i, C_IDCOMP)

      End If
   Next i

   LstComp = Mid(LstComp, 2)

   Q1 = ""
   Q1 = "Select IdComp,Correlativo,Fecha,Tipo,Estado,Glosa,TotalDebe,TotalHaber,IdUsuario,FechaCreacion, "
   Q1 = Q1 & " ImpResumido,EsCCMM,FechaImport,TipoAjuste,OtrosIngEg14TER,IdEmpresa,Ano,IdComp "
   Q1 = Q1 & " From Comprobante "
   Q1 = Q1 & " Where IdComp in (" & LstComp & ")"

    Set Rs = OpenRs(DbMain, Q1)

    Do While Rs.EOF = False

    Dim Fecha As Date
         Fecha = DateAdd("m", 1, FmtDate(vFld(Rs("Fecha"))))

    MesActual = month(Fecha)

            If gTipoCorrComp = TCC_UNICO Then

            If gPerCorrComp = TCC_MENSUAL Then   'si es anual o continuo sWhere = ""
               sWhere = SqlMonthLng("Fecha") & " = " & MesActual
            End If

         ElseIf gTipoCorrComp = TCC_TIPOCOMP Then
            sWhere = " Tipo = " & vFld(Rs("Tipo"))

            If gPerCorrComp = TCC_MENSUAL Then
               sWhere = sWhere & " AND " & SqlMonthLng("Fecha") & " = " & MesActual    'SQL Server tiene los días desplazados en dos
            End If

         End If

         'agregamos el tipo de ajuste
         If ItemData(Cb_TipoAjuste) = TAJUSTE_TRIBUTARIO Then
            WhTAjuste = " TipoAjuste = " & TAJUSTE_TRIBUTARIO
         Else
            WhTAjuste = " TipoAjuste IN ( " & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
         End If

         If sWhere <> "" Then
            sWhere = sWhere & " AND " & WhTAjuste
         Else
            sWhere = WhTAjuste
         End If

         If sWhere <> "" Then
            WhConWhere = " WHERE " & sWhere & " AND Correlativo > 0"
            WhConAnd = " AND " & sWhere  ' sin > 0
         Else
            WhConWhere = " WHERE Correlativo > 0"

         End If

            Q3 = "SELECT Max(Correlativo) as N FROM " & lTblComprobante & WhConWhere
            Q3 = Q3 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Set Rs3 = OpenRs(DbMain, Q3)

            Do While Rs3.EOF = False
            If Rs3.EOF = False Then
               lCorrelativo = vFld(Rs3("N")) + 1
            Else
               lCorrelativo = 1
            End If

            FldArray(0).FldName = "Correlativo"
            FldArray(0).FldValue = lCorrelativo
            FldArray(0).FldIsNum = True
            Rs3.MoveNext

            Loop

            Call CloseRs(Rs3)




         FldArray(1).FldName = "Fecha"
         FldArray(1).FldValue = CLng(Fecha)
         FldArray(1).FldIsNum = True

         FldArray(2).FldName = "Tipo"
         FldArray(2).FldValue = vFld(Rs("Tipo"))
         FldArray(2).FldIsNum = True

         FldArray(3).FldName = "Estado"
         FldArray(3).FldValue = EC_PENDIENTE
         FldArray(3).FldIsNum = True

         FldArray(4).FldName = "Glosa"
         FldArray(4).FldValue = vFld(Rs("Glosa"))
         FldArray(4).FldIsNum = False

         FldArray(5).FldName = "TotalDebe"
         FldArray(5).FldValue = vFld(Rs("TotalDebe"))
         FldArray(5).FldIsNum = True

         FldArray(6).FldName = "TotalHaber"
         FldArray(6).FldValue = vFld(Rs("TotalHaber"))
         FldArray(6).FldIsNum = True

         FldArray(7).FldName = "IdUsuario"
         FldArray(7).FldValue = vFld(Rs("IdUsuario"))
         FldArray(7).FldIsNum = True

         FldArray(8).FldName = "FechaCreacion"
         FldArray(8).FldValue = CLng(Int(Now))
         FldArray(8).FldIsNum = True

         FldArray(9).FldName = "ImpResumido"
         FldArray(9).FldValue = vFld(Rs("ImpResumido"))
         FldArray(9).FldIsNum = True

         FldArray(10).FldName = "EsCCMM"
         FldArray(10).FldValue = vFld(Rs("EsCCMM"))
         FldArray(10).FldIsNum = True

         FldArray(11).FldName = "FechaImport"
         FldArray(11).FldValue = vFld(Rs("FechaImport"))
         FldArray(11).FldIsNum = True

         FldArray(12).FldName = "TipoAjuste"
         FldArray(12).FldValue = vFld(Rs("TipoAjuste"))
         FldArray(12).FldIsNum = True

         FldArray(13).FldName = "OtrosIngEg14TeR"
         FldArray(13).FldValue = vFld(Rs("OtrosIngEg14TeR"))
         FldArray(13).FldIsNum = True

         FldArray(14).FldName = "IdEmpresa"
         FldArray(14).FldValue = vFld(Rs("IdEmpresa"))
         FldArray(14).FldIsNum = True

         FldArray(15).FldName = "Ano"
         FldArray(15).FldValue = vFld(Rs("Ano"))
         FldArray(15).FldIsNum = True


         lidComp = AdvTbAddNewMult(DbMain, lTblComprobante, "IdComp", FldArray)
         
         '3376884
         Call SeguimientoComprobantes(CLng(lidComp), gEmpresa.id, gEmpresa.Ano, "FrmLstComp.Bt_traspasar_Click", Q1, 1, "", gUsuario.IdUsuario, 1, 2)
         'fin 3376884

            If lidComp = -1 Then
              Call AddLogImp(FNameLogImp, "Traspaso Comprobante ", 0, "No fue posible Traspasar Comprobante " & vFld(Rs("IdComp")) & ".")
              ERR = ERR + 1
            End If

            Q2 = ""
            Q2 = "Select IdComp,IdDoc,Orden,IdCuenta,Debe,Haber,Glosa,idCCosto,idAreaNeg,IdCartola, "
            Q2 = Q2 & " DeCentraliz,DePago,DeRemu,Nota,IdDocCuota,IdEmpresa,Ano "
            Q2 = Q2 & " From MovComprobante "
            Q2 = Q2 & " Where IdComp =" & vFld(Rs("IdComp"))
            Q2 = Q2 & " Order By Orden asc"

             Set Rs2 = OpenRs(DbMain, Q2)

             Do While Rs2.EOF = False

         FldArrayMovCom(0).FldName = "IdComp"
         FldArrayMovCom(0).FldValue = lidComp
         FldArrayMovCom(0).FldIsNum = True

'         Dim Fecha As Date
'        Fecha = DateAdd("m", 1, FmtDate(vFld(Rs("Fecha"))))

         FldArrayMovCom(1).FldName = "IdDoc"
         FldArrayMovCom(1).FldValue = 0
         FldArrayMovCom(1).FldIsNum = True

         FldArrayMovCom(2).FldName = "Orden"
         FldArrayMovCom(2).FldValue = vFld(Rs2("Orden"))
         FldArrayMovCom(2).FldIsNum = True

         FldArrayMovCom(3).FldName = "IdCuenta"
         FldArrayMovCom(3).FldValue = vFld(Rs2("IdCuenta"))
         FldArrayMovCom(3).FldIsNum = True

         FldArrayMovCom(4).FldName = "Debe"
         FldArrayMovCom(4).FldValue = vFld(Rs2("Debe"))
         FldArrayMovCom(4).FldIsNum = False

         FldArrayMovCom(5).FldName = "Haber"
         FldArrayMovCom(5).FldValue = vFld(Rs2("Haber"))
         FldArrayMovCom(5).FldIsNum = True


         FldArrayMovCom(6).FldName = "Glosa"
         FldArrayMovCom(6).FldValue = vFld(Rs2("Glosa"))
         FldArrayMovCom(6).FldIsNum = True

         FldArrayMovCom(7).FldName = "IdCCosto"
         FldArrayMovCom(7).FldValue = vFld(Rs2("IdCCosto"))
         FldArrayMovCom(7).FldIsNum = True

         FldArrayMovCom(8).FldName = "IdAreaNeg"
         FldArrayMovCom(8).FldValue = vFld(Rs2("IdAreaNeg"))
         FldArrayMovCom(8).FldIsNum = True

         FldArrayMovCom(9).FldName = "idCartola"
         FldArrayMovCom(9).FldValue = vFld(Rs2("idCartola"))
         FldArrayMovCom(9).FldIsNum = True

         FldArrayMovCom(10).FldName = "DeCentraliz"
         FldArrayMovCom(10).FldValue = 0
         FldArrayMovCom(10).FldIsNum = True

         FldArrayMovCom(11).FldName = "DePago"
         FldArrayMovCom(11).FldValue = 0
         FldArrayMovCom(11).FldIsNum = True

         FldArrayMovCom(12).FldName = "DeRemu"
         FldArrayMovCom(12).FldValue = 0
         FldArrayMovCom(12).FldIsNum = True

         FldArrayMovCom(13).FldName = "Nota"
         FldArrayMovCom(13).FldValue = vFld(Rs2("Nota"))
         FldArrayMovCom(13).FldIsNum = True

         FldArrayMovCom(14).FldName = "IdDocCuota"
         FldArrayMovCom(14).FldValue = 0
         FldArrayMovCom(14).FldIsNum = True

         FldArrayMovCom(15).FldName = "IdEmpresa"
         FldArrayMovCom(15).FldValue = vFld(Rs2("IdEmpresa"))
         FldArrayMovCom(15).FldIsNum = True

         FldArrayMovCom(16).FldName = "Ano"
         FldArrayMovCom(16).FldValue = vFld(Rs2("Ano"))
         FldArrayMovCom(16).FldIsNum = True


          idMov = AdvTbAddNewMult(DbMain, lTblMovComprobante, "IdMov", FldArrayMovCom)
          
          '3376884
          Call SeguimientoMovComprobante(CLng(lidComp), gEmpresa.id, gEmpresa.Ano, "FrmLstComp.Bt_traspasar_Click", Q1, 1, "", 1, 1)
          'fin 3376884


            If idMov = -1 Then
            Call AddLogImp(FNameLogImp, "Traspaso Mov Comprobante ", 0, "No fue posible Traspasar Mov. Comprobante " & vFld(Rs("IdComp")) & ".")
            End If

    Rs2.MoveNext


    Loop

        Call CloseRs(Rs2)

    Rs.MoveNext

    n = n + 1
    Loop

  Call CloseRs(Rs)

  If ERR > 0 Then
   MsgBox1 "Se traspasaron " & n & " Comprobantes OK, no se traspaso " & ERR & ", Visualizar archivo " & FNameLogImp, vbExclamation + vbOKOnly
  Else

  MsgBox1 "Se traspasaron " & n & " Comprobantes .", vbInformation + vbOKOnly

  End If
 End If
End Sub
'2861591

'2861591
Public Function Export_DetComprobante() As Long
   Dim FPath As String
   Dim LogPath As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Rs2 As Recordset
   Dim Q2 As String
   Dim Buf As String
   Dim i As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Valor As Double
   Dim r As Integer
   Dim ExpDir As String
   Dim TipoArchivo As String
   Dim Fecha As Long
   Dim Descrip As String
   Dim fname As String
    Dim Wh As String
   Dim JoinStr As String
   Dim UsrJoin As String
   Dim lTblMovComprobante As String

   On Error Resume Next

   Sep = vbTab

   lTblMovComprobante = "MovComprobante"

   'Exportación HR-RAD BAseImponible 14D
   TipoArchivo = "Comprobantes_Det"

   ExpDir = gHRPath & "\RUTS"
   MkDir ExpDir

   ExpDir = ExpDir & "\" & Right("00000000" & gEmpresa.Rut, 8)
   MkDir ExpDir

   ExpDir = ExpDir & "\Comprobantes"
   MkDir ExpDir

   fname = TipoArchivo & "_" & Right(gEmpresa.Ano, 2) & ".xlsx"

   FPath = ExpDir & "\" & fname

   Fd = FreeFile
   ERR.Clear

'   Open FPath For Output As #Fd
   If ERR Then
      MsgErr FPath
      Export_DetComprobante = -ERR
      Exit Function
   End If

   On Error GoTo 0

  Wh = CreateWhere(JoinStr)

   Q1 = "SELECT DISTINCT Comprobante.IdComp, Comprobante.Correlativo, Comprobante.Fecha, "
   Q1 = Q1 & " Comprobante.Tipo, Comprobante.Estado, Comprobante.Glosa, "
   Q1 = Q1 & " Comprobante.TotalDebe, Comprobante.TotalHaber, Usuarios.Usuario, Comprobante.FechaImport, Comprobante.TipoAjuste"

   UsrJoin = " Comprobante LEFT JOIN Usuarios ON Comprobante.IdUsuario = Usuarios.IdUsuario "

   If JoinStr = "" Then
      Q1 = Q1 & " FROM " & UsrJoin
   Else
      Q1 = Q1 & " FROM (((( " & UsrJoin & ")" & JoinStr
   End If

   Q1 = Q1 & Wh

   If Wh <> "" Then
      Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Else
      Q1 = Q1 & " WHERE Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   End If
   Q1 = Q1 & " ORDER BY  Comprobante.IdComp "

   Set Rs = OpenRs(DbMain, Q1)

   r = 0
    'imprimimos el archivo
   Do While Rs.EOF = False

   Buf = Buf & "Número comprobante: " & vFld(Rs("Correlativo")) & Sep & "Tipo: " & gTipoComp(vFld(Rs("Tipo"))) & Sep & " Estado: " & gEstadoComp(vFld(Rs("Estado"))) & Sep & "Fecha: " & FmtDate(vFld(Rs("Fecha"))) & vbCrLf
   'Buf = "Número comprobante: " & vFld(Rs("Correlativo")) & vbTab & "Tipo: " & gTipoComp(vFld(Rs("Tipo"))) & vbTab & " Estado: " & gEstadoComp(vFld(Rs("Estado"))) & vbTab & "Fecha: " & FmtDate(vFld(Rs("Fecha")))
    

'   Print #Fd, Buf

   'Buf = ""
   

   Buf = Buf & "N° " & Sep & "Código Cuenta" & Sep & "Cuenta" & Sep & "Debe" & Sep & "Haber" & Sep & "Descripción" & Sep & "TD" & Sep & "Nº Doc." & Sep & "Entidad" & Sep & "Área Negocio" & Sep & "Centro Gestión" & Sep & "Act. Fijo" & vbCrLf
   'Buf = "N° " & vbTab & "Código Cuenta" & vbTab & "Cuenta" & vbTab & "Debe" & vbTab & "Haber" & vbTab & "Descripción" & vbTab & "TD" & vbTab & "Nº Doc." & vbTab & "Entidad" & vbTab & "Área Negocio" & vbTab & "Centro Gestión" & vbTab & "Act. Fijo"
   
   
'   Print #Fd, Buf

        Q2 = "SELECT "
        Q2 = Q2 & " IdMov, Orden, "
        Q2 = Q2 & lTblMovComprobante & ".IdCuenta, Cuentas.Codigo As CodCta, Cuentas.Nombre, Cuentas.Atrib" & ATRIB_ACTIVOFIJO & ",Cuentas.Atrib" & ATRIB_CONCILIACION & ","
        Q2 = Q2 & "Cuentas.Descripcion As DescCta, "
        Q2 = Q2 & lTblMovComprobante & ".Debe, " & lTblMovComprobante & ".Haber,"
        Q2 = Q2 & " Glosa," & lTblMovComprobante & ".IdAreaNeg," & lTblMovComprobante & ".IdCCosto, "
        Q2 = Q2 & " AreaNegocio.Descripcion As DescAreaNeg, CentroCosto.Descripcion As DescCCosto "
        Q2 = Q2 & ", Entidades.Nombre as NombEnt "
        Q2 = Q2 & ", NumDoc, TipoLib, TipoDoc, NumDocHasta, " & lTblMovComprobante & ".IdDoc, "
        Q2 = Q2 & lTblMovComprobante & ".DeCentraliz, " & lTblMovComprobante & ".DePago, " & lTblMovComprobante & ".Nota "
        Q2 = Q2 & " FROM ((((" & lTblMovComprobante
        Q2 = Q2 & " INNER JOIN Cuentas ON " & lTblMovComprobante & ".IdCuenta = Cuentas.IdCuenta "

           If gDbType = SQL_ACCESS Then
             Q2 = Q2 & "  AND " & lTblMovComprobante & ".IdEmpresa = Cuentas.IdEmpresa AND " & lTblMovComprobante & ".Ano = Cuentas.Ano)"
           Else
             Q2 = Q2 & "  AND " & lTblMovComprobante & ".IdEmpresa = Cuentas.IdEmpresa )"
           End If

        Q2 = Q2 & " LEFT JOIN AreaNegocio ON " & lTblMovComprobante & ".IdAreaNeg = AreaNegocio.IdAreaNegocio "
        Q2 = Q2 & "  AND " & lTblMovComprobante & ".IdEmpresa = AreaNegocio.IdEmpresa )"
        Q2 = Q2 & " LEFT JOIN CentroCosto ON " & lTblMovComprobante & ".IdCCosto = CentroCosto.IdCCosto "
        Q2 = Q2 & "  AND " & lTblMovComprobante & ".IdEmpresa = CentroCosto.IdEmpresa )"
        Q2 = Q2 & " LEFT JOIN Documento ON " & lTblMovComprobante & ".IdDoc=Documento.IdDoc "
        Q2 = Q2 & " AND " & lTblMovComprobante & ".IdEmpresa = Documento.IdEmpresa AND " & lTblMovComprobante & ".Ano = Documento.Ano)"
        Q2 = Q2 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
        Q2 = Q2 & " AND Entidades.IdEmpresa = Documento.IdEmpresa "
        Q2 = Q2 & " WHERE " & lTblMovComprobante & ".IdComp = " & vFld(Rs("IdComp"))
        Q2 = Q2 & " AND " & lTblMovComprobante & ".IdEmpresa = " & gEmpresa.id & " AND " & lTblMovComprobante & ".Ano = " & gEmpresa.Ano
        Q2 = Q2 & " ORDER BY Orden, IdMov"

        Set Rs2 = OpenRs(DbMain, Q2)

        Do While Rs2.EOF = False
'          Buf = ""
          Buf = Buf & vFld(Rs2("orden")) & Sep & vFld(Rs2("CodCta")) & Sep & vFld(Rs2("DescCta")) & Sep & vFld(Rs2("Debe")) & Sep & vFld(Rs2("Haber")) & Sep & vFld(Rs2("Glosa")) & Sep & GetDiminutivoDoc(vFld(Rs2("TipoLib")), vFld(Rs2("TipoDoc"))) & Sep & vFld(Rs2("NumDoc")) & Sep & vFld(Rs2("NombEnt"), True) & Sep & vFld(Rs2("DescAreaNeg"), True) & Sep & vFld(Rs2("DescCCosto"), True) & Sep & IIf(vFld(Rs2("Atrib" & ATRIB_ACTIVOFIJO)) <> 0, TX_ACTFIJO, "") & vbCrLf
          'Buf = vFld(Rs2("orden")) & vbTab & vFld(Rs2("CodCta")) & vbTab & vFld(Rs2("DescCta")) & vbTab & vFld(Rs2("Debe")) & vbTab & vFld(Rs2("Haber")) & vbTab & vFld(Rs2("Glosa")) & vbTab & GetDiminutivoDoc(vFld(Rs2("TipoLib")), vFld(Rs2("TipoDoc"))) & vbTab & vFld(Rs2("NumDoc")) & vbTab & vFld(Rs2("NombEnt"), True) & vbTab & vFld(Rs2("DescAreaNeg"), True) & vbTab & vFld(Rs2("DescCCosto"), True) & vbTab & IIf(vFld(Rs2("Atrib" & ATRIB_ACTIVOFIJO)) <> 0, TX_ACTFIJO, "")
          
'          Print #Fd, Buf

            Rs2.MoveNext
        Loop
        Call CloseRs(Rs2)

   Rs.MoveNext
   r = r + 1
   Loop

   Call CloseRs(Rs)

    Clipboard.Clear
    Clipboard.SetText Buf
'   Close Fd

   If r = 0 Then
      MsgBox1 "No existen datos para generar archivo.", vbInformation
      fname = ""
   Else
      FPath = ReplaceStr(FPath, "C:\HR\LPContab\..\", "C:\HR\")
      
        Dim oExcel As Object
        Dim oBook As Object
        Set oExcel = CreateObject("Excel.Application")
        Set oBook = oExcel.Workbooks.Add
       
        oExcel.DisplayAlerts = False
        
        'Paste the data
        oBook.Worksheets(1).Range("A1").Select
        oBook.Worksheets(1).Paste
    
'        'Open the text file
'         Set oBook = oExcel.Workbooks.Open(FPath)

        'Save as Excel workbook and Quit Excel
        oBook.SaveAs Replace(FPath, ".csv", ".xls")
        oExcel.Quit
      oExcel.DisplayAlerts = True
    
      MsgBox1 "Proceso de exportación Comprobante finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & FPath, vbInformation + vbOKOnly
   End If

   Export_DetComprobante = 0

End Function
'2861591

