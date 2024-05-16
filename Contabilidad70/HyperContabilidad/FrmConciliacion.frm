VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmConciliacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliación Bancaria"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   Icon            =   "FrmConciliacion.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   12600
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Cart 
      Caption         =   "Cartola del Banco"
      Height          =   2115
      Left            =   120
      TabIndex        =   34
      Top             =   6660
      Width           =   12315
      Begin VB.CommandButton Bt_AutoConcil 
         Caption         =   "Auto Conciliar..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   18
         ToolTipText     =   "Conciliación Automática"
         Top             =   1560
         Width           =   1635
      End
      Begin VB.CommandButton Bt_ImpCart 
         Caption         =   "Ingresar cartola..."
         Height          =   315
         Left            =   180
         TabIndex        =   17
         ToolTipText     =   "Ingresar o importar cartolas"
         Top             =   1140
         Width           =   1635
      End
      Begin VB.ComboBox Cb_Cartola 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   720
         Width           =   1875
      End
      Begin VB.ComboBox Cb_CartBanco 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   1875
      End
      Begin MSFlexGridLib.MSFlexGrid Gr_Cart 
         Height          =   1635
         Left            =   3060
         TabIndex        =   19
         ToolTipText     =   "Movimientos de la cartola. En azul los que calzan con el movimiento a conciliar."
         Top             =   300
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   10
         Cols            =   8
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin VB.Label La_Concil 
         AutoSize        =   -1  'True
         Caption         =   "Conciliados"
         Height          =   195
         Left            =   1740
         TabIndex        =   40
         Top             =   2100
         Width           =   810
      End
      Begin VB.Label La_nCalzan 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   2580
         TabIndex        =   37
         ToolTipText     =   "Cantidad de movimientos que calzan con el movimiento a conciliar"
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cartola::"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   36
         Top             =   780
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   35
         Top             =   420
         Width           =   510
      End
   End
   Begin VB.Frame Fr_Conc 
      Height          =   5955
      Left            =   120
      TabIndex        =   26
      Top             =   660
      Width           =   12315
      Begin VB.CommandButton Bt_VerComp 
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
         Index           =   1
         Left            =   11760
         Picture         =   "FrmConciliacion.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Ver Comprobante del registro seleccionado"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Bt_VerComp 
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
         Index           =   0
         Left            =   5520
         Picture         =   "FrmConciliacion.frx":042F
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Ver Comprobante del registro seleccionado"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   9900
         Picture         =   "FrmConciliacion.frx":0852
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   300
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   7860
         Picture         =   "FrmConciliacion.frx":0B5C
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   300
         Width           =   230
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
         Index           =   1
         Left            =   11340
         Picture         =   "FrmConciliacion.frx":0E66
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Copiar Excel"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Bt_Preview 
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
         Index           =   1
         Left            =   10560
         Picture         =   "FrmConciliacion.frx":12AB
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   840
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
         Index           =   1
         Left            =   10920
         Picture         =   "FrmConciliacion.frx":1752
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir"
         Top             =   840
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
         Index           =   0
         Left            =   5100
         Picture         =   "FrmConciliacion.frx":1C0C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Copiar Excel"
         Top             =   840
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
         Index           =   0
         Left            =   4680
         Picture         =   "FrmConciliacion.frx":2051
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Imprimir"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Bt_Preview 
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
         Index           =   0
         Left            =   4260
         Picture         =   "FrmConciliacion.frx":250B
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Bt_Buscar 
         Height          =   435
         Left            =   10980
         Picture         =   "FrmConciliacion.frx":29B2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar una cuenta"
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   6840
         TabIndex        =   2
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   8880
         TabIndex        =   3
         Top             =   300
         Width           =   1035
      End
      Begin VB.ComboBox Cb_Cuentas 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   4275
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   180
         Picture         =   "FrmConciliacion.frx":2F02
         ScaleHeight     =   480
         ScaleWidth      =   525
         TabIndex        =   29
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox Tx_Tit 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Movimientos por conciliar"
         Top             =   840
         Width           =   4035
      End
      Begin VB.TextBox Tx_Tit 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   6435
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Movimientos conciliados"
         Top             =   840
         Width           =   4035
      End
      Begin VB.CommandButton Bt_ToRight 
         Height          =   375
         Left            =   6045
         Picture         =   "FrmConciliacion.frx":3329
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Conciliar"
         Top             =   2580
         Width           =   315
      End
      Begin VB.CommandButton Bt_ToLeft 
         Height          =   375
         Left            =   6045
         Picture         =   "FrmConciliacion.frx":3633
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Anular conciliación"
         Top             =   3180
         Width           =   315
      End
      Begin MSFlexGridLib.MSFlexGrid Gr_Conciliado 
         Height          =   4215
         Left            =   6420
         TabIndex        =   11
         ToolTipText     =   "Movimientos conciliados. En azul movimientos en cartola del año siguiente."
         Top             =   1140
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   7435
         _Version        =   393216
         Rows            =   21
         Cols            =   9
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid Gr_PorConcil 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   1140
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   21
         Cols            =   8
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid Gr_TotPorConc 
         Height          =   285
         Left            =   0
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   5580
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   503
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   16711680
         ScrollTrack     =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid Gr_TotConc 
         Height          =   525
         Left            =   6420
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Total Conciliados y Total Cartola"
         Top             =   5340
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   926
         _Version        =   393216
         Rows            =   3
         Cols            =   8
         FixedRows       =   2
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   16711680
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   1
         Left            =   8340
         TabIndex        =   31
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Index           =   0
         Left            =   6240
         TabIndex        =   30
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuentas:"
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   0
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   12300
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
         Left            =   960
         Picture         =   "FrmConciliacion.frx":393D
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Calendario"
         Top             =   180
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
         Left            =   120
         Picture         =   "FrmConciliacion.frx":3D66
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Convertir moneda"
         Top             =   180
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
         Left            =   540
         Picture         =   "FrmConciliacion.frx":4104
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Calculadora"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   10920
         TabIndex        =   24
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   9660
         TabIndex        =   23
         Top             =   180
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Columnas movimientos
Const C_FECHA = 0
Const C_TIPODOC = 1
Const C_NUMDOC = 2
Const C_CARGO = 3
Const C_ABONO = 4
Const C_SALDO = 5
Const C_GLOSA = 6
Const C_IDCOMP = 7
Const C_NUMCOMP = 8
Const C_FECHACOMP = 9
Const C_IDMOV = 10
Const C_IDCART = 11

Const NCOLS = C_IDCART

' Columnas cartola
Private Const CC_FECHA = 0
Private Const CC_DETALLE = 1
Private Const CC_NRODOC = 2
Private Const CC_CARGO = 3
Private Const CC_ABONO = 4
Private Const CC_HFECHA = 5
Private Const CC_IDMOV = 6
Private Const CC_IDDET = 7

Const nRows = 10
Const ROW_HEIGHT = 240

Const C_PORCONCIL = 0
Const C_CONCIL = 1

Dim lOrdenGr(C_IDMOV) As String
Dim lOrdenSel As Integer
Dim ModConcil As Boolean
Private lSaldoIni As Double   ' Saldo inicial de la cartola seleccionada

Dim nCartMov As Integer
Dim nCartSelMov As Integer
Dim RowCartSel As Integer
Dim bNoMeg As Boolean

Private Const MovConc = "Movimientos Conciliados"

Dim lRc As Integer

#If DATACON = 1 Then
Dim lDbAnoAnt As Database
#End If

Private Sub Bt_AutoConcil_Click()
   Dim r As Integer, n As Integer, m As Integer, r1 As Integer
   Dim NConcil As Integer

   If MsgBox1("¿Desea que el sistema concilie automáticamente todos aquellos movimientos que calzan con los datos de la cartola seleccionada?", vbQuestion Or vbYesNo Or vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If

   Me.MousePointer = vbHourglass
   DoEvents

   r = 1
   r1 = 0
   m = 0
   NConcil = 0
   Do While r < Gr_PorConcil.rows
      
      If Trim(Gr_PorConcil.TextMatrix(r, C_FECHA)) = "" Then
         Exit Do
      End If
      
      n = SelMovCartola(r)
      If n = 1 Then
         If Conciliar(r, False) Then
            NConcil = NConcil + 1
         End If
      Else
         If n > 1 Then
            If r1 = 0 Then
               r1 = r
            End If
         
            m = m + 1
         End If
         
         r = r + 1
      End If

   Loop
   
   
   If NConcil > 0 Then
      If NConcil = 1 Then
         MsgBox1 "Se concilió un movimiento.", vbInformation + vbOKOnly
      Else
         MsgBox1 "Se conciliaron " & NConcil & " movimientos.", vbInformation + vbOKOnly
      End If
   ElseIf m = 0 Then
      MsgBox1 "No se encontraron movimientos para conciliar.", vbInformation + vbOKOnly
   End If
   
   If m Then
      MsgBox1 "Hay " & m & " movimientos por conciliar que calzan con más de un movimiento en la cartola." & vbLf & "Seleccione el movimiento correspondiente en la cartola con un doble-clic.", vbInformation
      
      Call FGrSelRow(Gr_PorConcil, r1)
      n = SelMovCartola(r1)
   End If

   Me.MousePointer = vbDefault

End Sub

Private Sub Bt_Buscar_Click()
   Dim F1 As Long, F2 As Long
   Dim DbName As String
   Dim AuxDb As Database
   
   F1 = GetTxDate(Tx_Desde)
   F2 = GetTxDate(Tx_Hasta)
   
   If F1 = 0 Then
      MsgBox1 "Debe ingresar la fecha de inicio.", vbExclamation
      Tx_Desde.SetFocus
      Exit Sub
   End If
   
   If F2 = 0 Then
      MsgBox1 "Debe ingresar la fecha de término", vbExclamation
      Tx_Hasta.SetFocus
      Exit Sub
   End If
   
   If F1 > F2 Then
      MsgBox1 "Fecha de inicio es mayor que la de término", vbExclamation
      Tx_Hasta.SetFocus
      Exit Sub
      
   End If
   
   If Year(F1) <> Year(F2) Then
      MsgBox1 "El rango de fechas debe estar dentro de un mismo período o año", vbExclamation
      Tx_Hasta.SetFocus
      Exit Sub
      
   End If
   
   If Year(F1) < gEmpresa.Ano - 1 Then
      MsgBox1 "La fecha de inicio es anterior al 1 de enero " & gEmpresa.Ano - 1, vbExclamation
      Tx_Desde.SetFocus
      Exit Sub
      
   End If
   
   If Year(F2) > gEmpresa.Ano Then
      MsgBox1 "La fecha de término es posterior al 31 de diciembre " & gEmpresa.Ano, vbExclamation
      Tx_Hasta.SetFocus
      Exit Sub
      
   End If
      
#If DATACON = 1 Then

   If Year(F1) = gEmpresa.Ano - 1 Then    'periodo anterior
   
      If lDbAnoAnt Is Nothing Then        'no está abierta la DB del año anterior
      
         'abrimos base de datos año anterior
      
         If gEmpresa.TieneAnoAnt Then
         
            If gEmprSeparadas Then
         
               DbName = gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb"
               
               If ExistFile(DbName) Then
                  Call OpenDbEmp2(lDbAnoAnt, gEmpresa.Rut, gEmpresa.Ano - 1)
                  'hacemos CorrigeBase de año anterior por si las moscas
                  Set AuxDb = DbMain
                  Set DbMain = lDbAnoAnt
                  Call CorrigeBase
                  Set DbMain = AuxDb
                  
               End If
               
            End If
      
         Else
            MsgBox1 "Esta empresa no tiene datos ingresados en el sistema para el periodo anterior.", vbExclamation
      
         End If
      End If
      
   End If
#End If
   
   Call LoadAll
   
End Sub

Private Sub FillCartList()
   Dim Q1 As String, Rs As Recordset, r As Integer
   
   Gr_Cart.Redraw = False
   
   Call CleanCartList
   
   Q1 = "SELECT SaldoIni, TotCargo, TotAbono, FHasta"
   Q1 = Q1 & " FROM Cartola"
   Q1 = Q1 & " WHERE idCartola=" & ItemData(Cb_Cartola)
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
   
      If Year(GetTxDate(Tx_Desde)) < gEmpresa.Ano Then
         If MsgBox1("¿Desea que el sistema ajuste automáticamente el periodo de acuerdo al mes de la cartola?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
            Call SetTxDate(Tx_Hasta, vFld(Rs("FHasta")))
            Call SetTxDate(Tx_Desde, DateSerial(gEmpresa.Ano, 1, 1))
         End If
      Else
         Call SetTxDate(Tx_Hasta, vFld(Rs("FHasta")))
         Call SetTxDate(Tx_Desde, DateSerial(gEmpresa.Ano, 1, 1))
      
      End If
      
      'Gr_TotConc.Clear
      Gr_TotConc.TextMatrix(1, C_FECHA) = "Tot. Cart."
      Gr_TotConc.TextMatrix(1, C_NUMDOC) = Cb_Cartola
      Gr_TotConc.TextMatrix(1, C_CARGO) = Format(vFld(Rs("TotCargo")), NUMFMT)
      Gr_TotConc.TextMatrix(1, C_ABONO) = Format(vFld(Rs("TotAbono")), NUMFMT)
      
      lSaldoIni = vFld(Rs("SaldoIni"))
      
      Gr_TotConc.TextMatrix(1, C_SALDO) = Format(lSaldoIni + vFld(Rs("TotAbono")) - vFld(Rs("TotCargo")), NEGBL_NUMFMT)

   Else
   
      lSaldoIni = 0

   End If
   Call CloseRs(Rs)
   
   Q1 = "SELECT Fecha, Detalle, NumDoc, Cargo, Abono, idDetCartola, idMov"
   Q1 = Q1 & " FROM DetCartola"
   Q1 = Q1 & " WHERE idCartola=" & ItemData(Cb_Cartola)
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Fecha, NumDoc"
   
   Gr_Cart.rows = 1
   
   Set Rs = OpenRs(DbMain, Q1)
   
   r = 0
   Do Until Rs.EOF
   
      Gr_Cart.AddItem ""
      r = r + 1
   
      Gr_Cart.TextMatrix(r, CC_FECHA) = Format(vFld(Rs("Fecha")), F_SHORTDATE)
      Gr_Cart.TextMatrix(r, CC_HFECHA) = vFld(Rs("Fecha"))
      Gr_Cart.TextMatrix(r, CC_DETALLE) = FCase(vFld(Rs("Detalle"), True))
      Gr_Cart.TextMatrix(r, CC_NRODOC) = vFld(Rs("NumDoc"))
      Gr_Cart.TextMatrix(r, CC_CARGO) = Format(vFld(Rs("Cargo")), NEGBL_NUMFMT)
      Gr_Cart.TextMatrix(r, CC_ABONO) = Format(vFld(Rs("Abono")), NEGBL_NUMFMT)
      Gr_Cart.TextMatrix(r, CC_IDDET) = vFld(Rs("idDetCartola"))
      Gr_Cart.TextMatrix(r, CC_IDMOV) = vFld(Rs("idMov"))
      
      If Val(Gr_Cart.TextMatrix(r, CC_IDMOV)) <> 0 Then
         Call FGrSetRowStyle(Gr_Cart, r, "FC", COLOR_VERDEOSCURO)
      End If
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   nCartMov = Gr_Cart.rows - 1
   
   Call FGrVRows(Gr_Cart)
   
   Bt_Buscar.Enabled = True
   Gr_Cart.Redraw = True
   
   Call PostClick(Bt_Buscar)
   
End Sub


Private Sub Bt_CopyExcel_Click(Index As Integer)

   If Index = C_PORCONCIL Then
      Call FGr2Clip(Gr_PorConcil, "Fecha: " & Tx_Desde & " al " & Tx_Hasta)
   Else
      Call FGr2Clip(Gr_Conciliado, "Fecha: " & Tx_Desde & " al " & Tx_Hasta)
   End If
   
End Sub

Private Sub Bt_Fecha_Click(Index As Integer)
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   If Index = 0 Then
      Call Frm.TxSelDate(Tx_Desde)
   Else
      Call Frm.TxSelDate(Tx_Hasta)
   End If
   
   Set Frm = Nothing
   
End Sub

Private Sub Bt_ImpCart_Click()
   Dim Q1 As String, Frm As FrmImpCartola
   Dim idCartola As Long, IdCuenta As Long
   
   IdCuenta = CbItemData(Cb_CartBanco)
   idCartola = CbItemData(Cb_Cartola)
   
   Set Frm = New FrmImpCartola
   Frm.Show vbModal
   Set Frm = Nothing
   
   If gRc.Rc = vbOK Then
      Cb_CartBanco.Clear
   
      Q1 = "SELECT Descripcion, idCuenta FROM Cuentas WHERE Atrib" & ATRIB_CONCILIACION & "<>0"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call FillCombo(Cb_CartBanco, DbMain, Q1, IdCuenta)
      
      Tx_Tit(C_CONCIL) = MovConc

      Call CbSelItem(Cb_Cartola, idCartola)
      
   End If

End Sub

Private Sub Bt_Preview_Click(Index As Integer)
   Dim Frm As FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   
   If Index = C_PORCONCIL Then
      Call SetUpPrtGrid(Gr_PorConcil, Gr_TotPorConc, Tx_Tit(Index))
   Else
      Call SetUpPrtGrid(Gr_Conciliado, Gr_TotConc, Tx_Tit(Index))
   End If
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print(Index)
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Bt_Print_Click(Index As Integer)
   Dim OldOrientation As Integer
   
   OldOrientation = Printer.Orientation
   
   If Index = C_PORCONCIL Then
      Call SetUpPrtGrid(Gr_PorConcil, Gr_TotPorConc, Tx_Tit(Index))
   Else
      Call SetUpPrtGrid(Gr_Conciliado, Gr_TotConc, Tx_Tit(Index))
   End If
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation

End Sub

Private Sub Bt_ToLeft_Click()
   Dim Row As Integer
   Dim Row2 As Integer
   
   If Bt_ToLeft.Enabled = False Then
      Exit Sub
   End If

   Row = Gr_Conciliado.Row
   
   If Trim(Gr_Conciliado.TextMatrix(Row, C_FECHA)) = "" Or Row = 0 Or Trim(Gr_Conciliado.TextMatrix(Row, C_IDMOV)) = "" Then
      Exit Sub
   End If
   
   Row2 = FGrAddRow(Gr_PorConcil)
   
   Call ChangeGrid(Row, Gr_Conciliado, Gr_TotConc, Row2, Gr_PorConcil, Gr_TotPorConc)

End Sub




Private Sub Bt_ToRight_Click()
   Dim Row As Integer
   Dim Row2 As Integer
   
   If Bt_ToRight.Enabled = False Then
      Exit Sub
   End If
   
   Row = Gr_PorConcil.Row
   
   If Trim(Gr_PorConcil.TextMatrix(Row, C_FECHA)) = "" Or Row = 0 Or Trim(Gr_PorConcil.TextMatrix(Row, C_IDMOV)) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Cb_Cartola.ListIndex < 0 Then
      MsgBox1 "Seleccione una cartola.", vbExclamation
      Exit Sub
   End If
     
   Call Conciliar(Row, True, True)
        
End Sub

Private Sub Bt_VerComp_Click(Index As Integer)
   Dim Row As Integer
   Dim Col As Integer
   Dim Frm As FrmComprobante
   Dim Grid As MSFlexGrid
   
   If Index = C_PORCONCIL Then
      Set Grid = Gr_PorConcil
   Else
      Set Grid = Gr_Conciliado
   End If

   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_IDCOMP)) > 0 Then
      Set Frm = New FrmComprobante
      Call Frm.FView(Val(Grid.TextMatrix(Row, C_IDCOMP)), False)
      Set Frm = Nothing
   End If

End Sub

Private Sub Cb_CartBanco_Click()
   Dim Q1 As String
   
   Cb_Cartola.Clear
   
   Q1 = "SELECT " & SqlConcat(gDbType, "Ano", "' - '", "Cartola") & " as Cart, idCartola"
   Q1 = Q1 & " FROM Cartola"
   Q1 = Q1 & " WHERE idCuentaBco=" & ItemData(Cb_CartBanco)
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Ano Desc, cartola Desc"
   Call FillCombo(Cb_Cartola, DbMain, Q1, -2)
   
   Call CleanCartList
   
   
End Sub

Private Sub Cb_Cartola_Click()

   MousePointer = vbHourglass
   DoEvents

   Call FillCartList

   DoEvents

   Call SelMovCartola(Gr_PorConcil.Row)

   Bt_AutoConcil.Enabled = (Cb_Cartola.ListIndex >= 0 And nCartMov >= 1)

   Call FGrSetRowVisible(Gr_Cart, Gr_Cart.FixedRows)
   Call FGrSetRowVisible(Gr_PorConcil, Gr_PorConcil.FixedRows)
   Call FGrSetRowVisible(Gr_Conciliado, Gr_Conciliado.FixedRows)

   Call PostClick(Bt_Buscar)

   MousePointer = vbDefault

End Sub

Private Sub Cb_Cuentas_Click()
   Call FrmEnable(False)
End Sub

Private Sub Form_Load()
   Dim F1 As Long, F2 As Long
   Dim Q1 As String
   Dim ActDate As Long
   Dim MesActual As Integer
   
   lRc = vbCancel
   
   Call CorrigeCart
   
   'Lleno el arreglo de orden de columnas
   lOrdenGr(C_FECHA) = "Comprobante.Fecha, Documento.NumDoc, MovComprobante.Glosa, MovComprobante.idMov"
   lOrdenGr(C_TIPODOC) = "Documento.TipoDoc, Comprobante.Fecha"
   lOrdenGr(C_GLOSA) = "MovComprobante.Glosa"
   lOrdenGr(C_NUMDOC) = "Documento.NumDoc"
   lOrdenGr(C_CARGO) = "Documento.Haber"
   lOrdenGr(C_ABONO) = "Documento.Debe"
   lOrdenGr(C_SALDO) = "(Documento.Debe-Documento.Haber)"
   
   lOrdenSel = C_FECHA
   
   Call BtFechaImg(Bt_Fecha(0))
   Call BtFechaImg(Bt_Fecha(1))
   
   'Call FirstLastMonthDay(Int(Now), F1, F2)
   
   MesActual = GetMesActual()
   If MesActual = 0 Then
      MesActual = GetUltimoMesConMovs()
   End If
   
   ActDate = DateSerial(gEmpresa.Ano, MesActual, 1)
   
   Call FirstLastMonthDay(ActDate, F1, F2)
   
   F1 = DateSerial(gEmpresa.Ano, 1, 1)
   
   Call SetTxDate(Tx_Desde, F1)
   Call SetTxDate(Tx_Hasta, F2)
   
   Call SetUpGrid(Gr_Conciliado, Gr_TotConc)
   'Gr_Conciliado.ColWidth(C_IDCART) = 1000
   
   Call SetUpGrid(Gr_PorConcil, Gr_TotPorConc)
   
   Call FillCtasConcil(Cb_Cuentas)
   
   Call LoadAll
   
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   If gEmpresa.FCierre Then
      MsgBox1 "No se puede conciliar porque el período está cerrado.", vbInformation
   End If
   
   Call SetupPriv
   
   Call SetTxRO(Tx_Tit(C_PORCONCIL), True)
   Call SetTxRO(Tx_Tit(C_CONCIL), True)
   
   Q1 = "SELECT Descripcion, idCuenta FROM Cuentas WHERE Atrib" & ATRIB_CONCILIACION & "<>0"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call FillCombo(Cb_CartBanco, DbMain, Q1, -1)

   Call SetupForm
   
End Sub
Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_OK_Click()
   Dim Q1 As String, Rc As Long, r As Integer
   
   MousePointer = vbHourglass
   DoEvents
   
   Call SaveGrid(Gr_Conciliado, True)
   
   Call SaveGrid(Gr_PorConcil, False)
   
   For r = Gr_Cart.FixedRows To Gr_Cart.rows - 1
   
      If Gr_Cart.TextMatrix(r, CC_IDDET) = "" Then
         Exit For
      End If
      
      Q1 = "UPDATE DetCartola SET idMov=" & Val(Gr_Cart.TextMatrix(r, CC_IDMOV))
      Q1 = Q1 & " WHERE idDetCartola=" & Val(Gr_Cart.TextMatrix(r, CC_IDDET))
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Rc = ExecSQL(DbMain, Q1)
      
   Next r
   
   lRc = vbOK
   
   MousePointer = vbDefault
   
   Unload Me
   
End Sub
Private Sub SetUpGrid(Grid As MSFlexGrid, GridTot As Control)
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
   Grid.rows = nRows
   GridTot.Cols = Grid.Cols
   
   Grid.ColWidth(C_FECHA) = 800
   Grid.ColWidth(C_TIPODOC) = 300
   Grid.ColWidth(C_NUMDOC) = 780
   Grid.ColWidth(C_CARGO) = 1180
   Grid.ColWidth(C_ABONO) = 1180
   Grid.ColWidth(C_SALDO) = 1180
   Grid.ColWidth(C_GLOSA) = 4200
   Grid.ColWidth(C_IDCOMP) = 0
   Grid.ColWidth(C_NUMCOMP) = 1000
   Grid.ColWidth(C_FECHACOMP) = 1000

   Grid.ColWidth(C_IDMOV) = 0
   Grid.ColWidth(C_IDCART) = 0
   
   For i = 0 To Grid.Cols - 1
      GridTot.ColWidth(i) = Grid.ColWidth(i)
   Next i
      
   Grid.ColAlignment(C_FECHA) = flexAlignLeftCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignRightCenter
   Grid.ColAlignment(C_CARGO) = flexAlignRightCenter
   Grid.ColAlignment(C_ABONO) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDO) = flexAlignRightCenter
   Grid.ColAlignment(C_GLOSA) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPODOC) = flexAlignLeftCenter

   
   Grid.TextMatrix(0, C_FECHA) = "Fecha"
   Grid.TextMatrix(0, C_NUMDOC) = "N° Doc"
   Grid.TextMatrix(0, C_CARGO) = "Cargo"
   Grid.TextMatrix(0, C_ABONO) = "Abono"
   Grid.TextMatrix(0, C_SALDO) = "Saldo"
   Grid.TextMatrix(0, C_GLOSA) = "Detalle"
   Grid.TextMatrix(0, C_NUMCOMP) = "Comprob."
   Grid.TextMatrix(0, C_FECHACOMP) = "Fecha Comp."
   Grid.TextMatrix(0, C_GLOSA) = "Detalle"
   Grid.TextMatrix(0, C_TIPODOC) = "TD"
   
   Call FGrSetup(Grid)

     
   'Marco la columna Ordenada
   Grid.Row = 0
   Grid.Col = C_FECHA
   Set Grid.CellPicture = FrmMain.Pc_Flecha
   
End Sub

Public Function FEdit() As Integer
   Me.Show vbModal
   
   FEdit = lRc
End Function
Private Sub LoadAll(Optional ByVal LstIdPorCon As String = "", Optional ByVal LstIdCon As String = "")
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   Dim Row2 As Integer
   Dim WhereCta As String
   Dim i As Integer
   Dim WhereBas As String
   
   Gr_Conciliado.Redraw = False
   Gr_PorConcil.Redraw = False
   
   If ItemData(Cb_Cartola) > 0 Then
      Tx_Tit(C_CONCIL) = MovConc & " - cartola " & Cb_Cartola
   End If
   
   Row = 1
   Gr_Conciliado.rows = 1
   For i = 0 To Gr_TotConc.Cols - 1
      Gr_TotConc.TextMatrix(0, i) = ""
   Next i
   
   Gr_TotConc.TextMatrix(0, C_FECHA) = "Total"
   
   Row2 = 1
   Gr_PorConcil.rows = 1
   Gr_TotPorConc.Clear
   Gr_TotPorConc.TextMatrix(0, C_FECHA) = "Total"
   
   WhereCta = ""
   If ItemData(Cb_Cuentas) <> 0 Then
      WhereCta = " MovComprobante.idCuenta=" & ItemData(Cb_Cuentas)
   Else
      WhereCta = " Cuentas.Atrib" & ATRIB_CONCILIACION & "<> 0"
      'WhereCta = " Cuentas.Atrib" & ATRIB_CUENTABANCO & "<> 0"
   End If
   
   WhereBas = " Comprobante.Estado IN (" & EC_APROBADO & ", " & EC_PENDIENTE & ")"
   WhereBas = WhereBas & " AND Comprobante.TipoAjuste IN ( " & TAJUSTE_FINANCIERO & ", " & TAJUSTE_AMBOS & ")"
   
   WhereBas = WhereBas & " AND" & WhereCta
   
   
   'SALDO INICIAL
   Call SaldoInicial(True, Gr_Conciliado, Row, WhereBas)
   Call SaldoInicial(False, Gr_PorConcil, Row2, WhereBas)
   
   Q1 = "SELECT Fecha, MovComprobante.Glosa, TipoLib, TipoDoc, NumDoc, MovComprobante.Debe, MovComprobante.Haber, "
   Q1 = Q1 & " MovComprobante.IdCartola, MovComprobante.idMov, Comprobante.IdComp, Comprobante.Tipo, Comprobante.Correlativo, Comprobante.Fecha as FechaComp "
   Q1 = Q1 & " FROM ((MovComprobante "
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.idCuenta=Cuentas.idCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN Documento ON MovComprobante.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovComprobante")
   Q1 = Q1 & " WHERE Comprobante.Estado IN (" & EC_APROBADO & ", " & EC_PENDIENTE & ")"
   Q1 = Q1 & " AND Comprobante.TipoAjuste IN ( " & TAJUSTE_FINANCIERO & ", " & TAJUSTE_AMBOS & ")"
   
   If LstIdPorCon <> "" Then
      Q1 = Q1 & " AND IdMov IN (" & LstIdPorCon & "," & LstIdCon & ")"
   Else
      Q1 = Q1 & " AND" & WhereCta
      ' Traemos sólo los de la cartola seleccionada y los no conciliados
      
      If Year(GetTxDate(Tx_Desde)) = gEmpresa.Ano - 1 Then    'año anterior
         Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano - 1
         Q1 = Q1 & " AND (IdCartola=" & -1 * ItemData(Cb_Cartola)
      Else
         Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
         Q1 = Q1 & " AND (( IdCartola=" & ItemData(Cb_Cartola) & " OR IdCartola < 0 )"
      End If
      
      Q1 = Q1 & " OR ((IdCartola IS NULL OR IdCartola=0) AND Fecha BETWEEN " & GetTxDate(Tx_Desde) & " AND " & GetTxDate(Tx_Hasta) & "))"
   End If
   
   
   Q1 = Q1 & " ORDER BY " & lOrdenGr(lOrdenSel)
   
#If DATACON = 1 Then
   If Year(GetTxDate(Tx_Desde)) = gEmpresa.Ano - 1 Then
      If gEmprSeparadas Then
         Call OpenDbEmp2(lDbAnoAnt, gEmpresa.Rut, gEmpresa.Ano - 1)
         Set Rs = OpenRs(lDbAnoAnt, Q1)
      Else
         Set Rs = OpenRs(DbMain, Q1)
      End If
   Else
      Set Rs = OpenRs(DbMain, Q1)
   End If

#Else
   Set Rs = OpenRs(DbMain, Q1)

#End If
  
   If LstIdPorCon <> "" Then
      LstIdCon = "," & LstIdCon & ","
      LstIdPorCon = "," & LstIdPorCon & ","
   End If
  
   Do While Rs.EOF = False
   
      If LstIdPorCon = "" Then
   
         If vFld(Rs("IdCartola")) <> 0 Then
            Row = FillGrid(Gr_Conciliado, Gr_TotConc, Row, Rs)
         Else
            Row2 = FillGrid(Gr_PorConcil, Gr_TotPorConc, Row2, Rs)
         End If
      
      Else
      
         If InStr(LstIdPorCon, "," & vFld(Rs("IdMov")) & ",") > 0 Then
            Row2 = FillGrid(Gr_PorConcil, Gr_TotPorConc, Row2, Rs)
         Else
            Row = FillGrid(Gr_Conciliado, Gr_TotConc, Row, Rs)
        End If
      
      End If
      
      Rs.MoveNext
      
   Loop
   Call CloseRs(Rs)
   
   Call FGrVRows(Gr_Conciliado)
   Call FGrVRows(Gr_PorConcil)
   Call FrmEnable(True)
   
   'Marco la columna Ordenada
   Gr_Conciliado.Row = 0
   Gr_Conciliado.Col = lOrdenSel
   Set Gr_Conciliado.CellPicture = FrmMain.Pc_Flecha

   Gr_Conciliado.Col = C_FECHA
   Gr_Conciliado.RowSel = Gr_Conciliado.Row
   Gr_Conciliado.ColSel = Gr_Conciliado.Col

   Gr_PorConcil.Row = 0
   Gr_PorConcil.Col = lOrdenSel
   Set Gr_PorConcil.CellPicture = FrmMain.Pc_Flecha

   Gr_PorConcil.Col = C_FECHA
   Gr_PorConcil.RowSel = Gr_PorConcil.Row
   Gr_PorConcil.ColSel = Gr_PorConcil.Col

   Gr_Conciliado.Redraw = True
   Gr_PorConcil.Redraw = True

End Sub
Private Function FillGrid(Grid As Control, GridTot As Control, Row As Integer, Rs As Recordset) As Integer
   Dim SaldoAnt As Double, Cargo As Double, Abono As Double

   Grid.rows = Row + 1
   
   Grid.TextMatrix(Row, C_FECHA) = Format(vFld(Rs("Fecha")), F_SHORTDATE)
   Grid.TextMatrix(Row, C_GLOSA) = FCase(vFld(Rs("Glosa")))
   Grid.TextMatrix(Row, C_NUMDOC) = vFld(Rs("NumDoc"))
   Grid.TextMatrix(Row, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
   Grid.TextMatrix(Row, C_IDMOV) = vFld(Rs("idMov"))
   Grid.TextMatrix(Row, C_IDCOMP) = vFld(Rs("IdComp"))
   Grid.TextMatrix(Row, C_NUMCOMP) = Left(gTipoComp(vFld(Rs("Tipo"))), 1) & "-" & vFld(Rs("Correlativo"))
   Grid.TextMatrix(Row, C_FECHACOMP) = Format(vFld(Rs("FechaComp")), F_SHORTDATE)
  
   If StrComp(Grid.Name, Gr_Conciliado.Name, vbTextCompare) = 0 Then
      Grid.TextMatrix(Row, C_IDCART) = vFld(Rs("idCartola"))
      If vFld(Rs("idCartola")) < 0 And Year(vFld(Rs("Fecha"))) = gEmpresa.Ano Then
         Call FGrSetRowStyle(Grid, Row, "FC", COLOR_AZULOSCURO)
      End If
      
   End If
   
   Cargo = 0
   Abono = 0
   
   If vFld(Rs("Haber")) <> 0 Then
      Cargo = vFld(Rs("Haber"))
   
      Grid.TextMatrix(Row, C_CARGO) = Format(Cargo, NUMFMT)
      GridTot.TextMatrix(0, C_CARGO) = Format(vFmt(GridTot.TextMatrix(0, C_CARGO)) + Cargo, NUMFMT)
   Else
      Abono = vFld(Rs("Debe"))
            
      Grid.TextMatrix(Row, C_ABONO) = Format(Abono, NUMFMT)
      GridTot.TextMatrix(0, C_ABONO) = Format(vFmt(GridTot.TextMatrix(0, C_ABONO)) + Abono, NUMFMT)
   End If
   
   'If vFmt(Grid.TextMatrix(Row - 1, C_SALDO)) <> 0 Then
   If Grid.TextMatrix(Row - 1, C_SALDO) <> "" Then
   
      If Row = Grid.FixedRows And Grid.Name = Gr_Conciliado.Name Then
         SaldoAnt = lSaldoIni
      Else
         SaldoAnt = vFmt(Grid.TextMatrix(Row - 1, C_SALDO))
      End If
   
      Grid.TextMatrix(Row, C_SALDO) = Format(SaldoAnt - Cargo + Abono, NEGNUMFMT)
      GridTot.TextMatrix(0, C_SALDO) = Grid.TextMatrix(Row, C_SALDO)
   End If
      
   Row = Row + 1
   
   FillGrid = Row
   
End Function


Private Sub Form_Resize()
   Dim Hei As Long, wID As Long

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Hei = Me.Height - W.YCaption - W.yFrame * 3

   Fr_Cart.Height = Hei - Fr_Cart.Top
   Gr_Cart.Height = Fr_Cart.Height - Gr_Cart.Top * 1.5

   Call FGrVRows(Gr_Cart)

   wID = Me.Width - W.xFrame * 2
   Fr_Conc.Width = wID
   Fr_Cart.Width = wID

   Gr_PorConcil.Width = (wID - Bt_ToRight.Width) / 2
   Gr_Conciliado.Width = (wID - Bt_ToRight.Width) / 2 - 15

   Bt_ToRight.Left = Gr_PorConcil.Left + Gr_PorConcil.Width
   Bt_ToLeft.Left = Bt_ToRight.Left

   Gr_Conciliado.Left = Bt_ToRight.Left + Bt_ToRight.Width

   Gr_TotConc.Left = Gr_Conciliado.Left

   Tx_Tit(C_CONCIL).Left = Gr_Conciliado.Left

   Bt_Preview(C_CONCIL).Left = Tx_Tit(C_CONCIL).Left + Tx_Tit(C_CONCIL).Width
   Bt_Print(C_CONCIL).Left = Bt_Preview(C_CONCIL).Left + Bt_Preview(C_CONCIL).Width
   Bt_CopyExcel(C_CONCIL).Left = Bt_Print(C_CONCIL).Left + Bt_Print(C_CONCIL).Width
   Bt_VerComp(C_CONCIL).Left = Bt_CopyExcel(C_CONCIL).Left + Bt_CopyExcel(C_CONCIL).Width

End Sub

Private Sub Form_Unload(Cancel As Integer)

#If DATACON = 1 Then
   If Not lDbAnoAnt Is Nothing Then
      Call OpenDbEmp2(lDbAnoAnt, gEmpresa.Rut, gEmpresa.Ano - 1)
      Call CloseDb(lDbAnoAnt)
      Set lDbAnoAnt = Nothing
   End If
#End If

End Sub

Private Sub Gr_Cart_DblClick()
   Dim Row As Integer, r As Integer
   Dim NroDoc As Long, Cargo As Double, Abono As Double
   
   Row = Gr_Cart.Row
   If Gr_Cart.TextMatrix(Row, CC_HFECHA) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Val(Gr_Cart.TextMatrix(Row, CC_IDMOV)) Then
      MsgBox1 "Este detalle de la cartola ya está conciliado.", vbExclamation
      Exit Sub
   End If
   
   r = Gr_PorConcil.Row
   
   NroDoc = Val(Right(Trim(Gr_PorConcil.TextMatrix(r, C_NUMDOC)), 6)) Mod 1000
   NroDoc = Val(Gr_PorConcil.TextMatrix(r, C_NUMDOC)) Mod 1000
   Cargo = vFmt(Gr_PorConcil.TextMatrix(r, C_CARGO))
   Abono = vFmt(Gr_PorConcil.TextMatrix(r, C_ABONO))

   If CalzaMov(NroDoc, Cargo, Abono, Row) Then
      nCartSelMov = 1
      RowCartSel = Row
      Call Bt_ToRight_Click
   Else
      MsgBox1 "No calzan el movimiento por conciliar con el detalle de la cartola.", vbExclamation
   End If

End Sub

Private Sub Gr_Cart_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Gr_Cart, "Cartola: " & Cb_Cartola & " - " & Cb_CartBanco)
   
   End If

End Sub

Private Sub Gr_Conciliado_Click()

'   No se ordena porque no tiene sentido por el saldo

'   Dim Col As Integer
'   Dim Row As Integer
'
'   Row = Gr_Conciliado.MouseRow
'   Col = Gr_Conciliado.MouseCol
'
'   If Col = C_SALDO Then
'      Exit Sub
'   End If
'
'   If Row >= Gr_Conciliado.FixedRows Then
'      Exit Sub
'   End If
'
'   Call OrdenaPorCol(Col)
      
End Sub

Private Sub Gr_Conciliado_DblClick()
   Dim Col As Integer
   Dim Row As Integer
         
   Row = Gr_Conciliado.MouseRow
   Col = Gr_Conciliado.MouseCol
      
   If Row < Gr_Conciliado.FixedRows Then
      Call OrdenaPorCol(Col)
      Exit Sub
   End If
   
   'Call Bt_ToLeft_Click
   Call PostClick(Bt_ToLeft)
End Sub

Private Sub Gr_Conciliado_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Gr_Conciliado, "Comprobantes Conciliados")
   End If

End Sub

Private Sub Gr_Conciliado_Scroll()
    Gr_TotConc.LeftCol = Gr_Conciliado.LeftCol
End Sub

Private Sub Gr_PorConcil_Click()

   'MsgBeep vbExclamation

'   No se ordena porque no tiene sentido por el saldo

'   Dim Col As Integer
'   Dim Row As Integer
'
'   Row = Gr_PorConcil.MouseRow
'   Col = Gr_PorConcil.MouseCol
'
'   If Col = C_SALDO Then
'      Exit Sub
'   End If
'
'   If Row >= Gr_PorConcil.FixedRows Then
'      Exit Sub
'   End If
'
'   Call OrdenaPorCol(Col)
   
End Sub

Private Sub Gr_PorConcil_DblClick()
   Dim Col As Integer
   Dim Row As Integer
         
   Row = Gr_PorConcil.MouseRow
   Col = Gr_PorConcil.MouseCol
      
   If Row < Gr_PorConcil.FixedRows Then
      Call OrdenaPorCol(Col)
      Exit Sub
   End If
   
   Call PostClick(Bt_ToRight)
End Sub

Private Sub Gr_PorConcil_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCopy(KeyCode, Shift) Then
      Call FGr2Clip(Gr_PorConcil, "Comprobantes por Conciliar")
   End If
End Sub

Private Sub Gr_PorConcil_Scroll()
   Gr_TotPorConc.LeftCol = Gr_PorConcil.LeftCol
End Sub

Private Sub Gr_PorConcil_SelChange()

   If Gr_PorConcil.Row < Gr_PorConcil.FixedRows Then
      Exit Sub
   End If

   MousePointer = vbHourglass
   DoEvents

   Call SelMovCartola(Gr_PorConcil.Row)

   MousePointer = vbDefault

End Sub

Private Sub tx_Desde_Change()
   Call FrmEnable(False)
End Sub

Private Sub Tx_Desde_GotFocus()
   If gEmpresa.FCierre <> 0 Then
      Exit Sub
   End If
   Call DtGotFocus(Tx_Desde)
End Sub

Private Sub Tx_Desde_LostFocus()
   If gEmpresa.FCierre <> 0 Then
      Exit Sub
   End If
   
   If Trim$(Tx_Desde) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Desde)
   
End Sub

Private Sub Tx_Desde_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub

Private Sub tx_Hasta_Change()
   Call FrmEnable(False)
End Sub

Private Sub Tx_Hasta_GotFocus()
   If gEmpresa.FCierre <> 0 Then
      Exit Sub
   End If
   
   Call DtGotFocus(Tx_Hasta)
   
End Sub

Private Sub Tx_Hasta_LostFocus()

   If gEmpresa.FCierre <> 0 Then
      Exit Sub
   End If
   
   If Trim$(Tx_Hasta) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Hasta)
      
End Sub

Private Sub Tx_Hasta_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub

Private Sub ChangeGrid(Row As Integer, GrFrom As Control, GrFromTot As Control, Row2 As Integer, GrTo As Control, GrToTot As Control, Optional ByVal vBtLeft As Boolean = False)
   Dim i As Integer, idMov As Long
   Dim Cargo As Double, Abono As Double, IniFrom As Double, IniTo As Double
   
   'AUMENTAR EN GRILLA
   If GrTo.RowHeight(Row2) = 0 Then
      GrTo.RowHeight(Row2) = ROW_HEIGHT
   End If
   
   idMov = Val(GrFrom.TextMatrix(Row, C_IDMOV))
   
   GrTo.TextMatrix(Row2, C_FECHA) = GrFrom.TextMatrix(Row, C_FECHA)
   GrTo.TextMatrix(Row2, C_GLOSA) = GrFrom.TextMatrix(Row, C_GLOSA)
   GrTo.TextMatrix(Row2, C_NUMDOC) = GrFrom.TextMatrix(Row, C_NUMDOC)
   GrTo.TextMatrix(Row2, C_TIPODOC) = GrFrom.TextMatrix(Row, C_TIPODOC)
   GrTo.TextMatrix(Row2, C_CARGO) = GrFrom.TextMatrix(Row, C_CARGO)
   GrTo.TextMatrix(Row2, C_ABONO) = GrFrom.TextMatrix(Row, C_ABONO)
   '3217627
   GrTo.TextMatrix(Row2, C_NUMCOMP) = GrFrom.TextMatrix(Row, C_NUMCOMP)
   GrTo.TextMatrix(Row2, C_FECHACOMP) = GrFrom.TextMatrix(Row, C_FECHACOMP)
   '3217627
   
   'GrTo.TextMatrix(Row2, C_SALDO) = GrFrom.TextMatrix(Row, C_SALDO)
   GrTo.TextMatrix(Row2, C_IDMOV) = idMov
   
   If GrTo.Name = Gr_Conciliado.Name Then ' Conciliar
     '3217627
     If vBtLeft = True Then
      If RowCartSel = 0 Then
       RowCartSel = Gr_Cart.Row
      End If
     End If
     '3217627
      If Year(vFmtTxtDate(GrTo.TextMatrix(Row2, C_FECHA))) = gEmpresa.Ano Then
         GrTo.TextMatrix(Row2, C_IDCART) = ItemData(Cb_Cartola)
         
         Gr_Cart.TextMatrix(RowCartSel, CC_IDMOV) = idMov
      Else
         GrTo.TextMatrix(Row2, C_IDCART) = -1 * ItemData(Cb_Cartola)    'si es del año anterior le ponemos el IdCartola en negativo
         Gr_Cart.TextMatrix(RowCartSel, CC_IDMOV) = -1 * idMov
      End If
      
      'marcamos el que conciliamos
      Call FGrSetRowStyle(Gr_Cart, RowCartSel, "FC", COLOR_VERDEOSCURO)
      
      ' desmarcamos los que no calzan
      
      'Call FGrSetRowStyle(Gr_Cart, RowCartSel, "FC", vbWindowText)
      
   Else  ' Des-Conciliar
      For i = 1 To Gr_Cart.rows - 1
         If Val(Gr_Cart.TextMatrix(i, CC_IDMOV)) = idMov Then
            Call FGrSetRowStyle(Gr_Cart, i, "FC", vbWindowText)
            Gr_Cart.TextMatrix(i, CC_IDMOV) = ""
            Exit For
         End If
      Next i
   
   End If
   
   GrTo.rows = GrTo.rows + 1
   
   Cargo = vFmt(GrFrom.TextMatrix(Row, C_CARGO))
   Abono = vFmt(GrFrom.TextMatrix(Row, C_ABONO))
         
   If GrTo.Name = Gr_Conciliado.Name Then
      IniTo = lSaldoIni
      IniFrom = 0
   Else
      IniTo = 0
      IniFrom = lSaldoIni
   End If
            
   GrToTot.TextMatrix(0, C_CARGO) = Format(vFmt(GrToTot.TextMatrix(0, C_CARGO)) + Cargo, NUMFMT)
   GrToTot.TextMatrix(0, C_ABONO) = Format(vFmt(GrToTot.TextMatrix(0, C_ABONO)) + Abono, NUMFMT)
   GrToTot.TextMatrix(0, C_SALDO) = Format(IniTo + vFmt(GrToTot.TextMatrix(0, C_ABONO)) - vFmt(GrToTot.TextMatrix(0, C_CARGO)), NEGNUMFMT)
            
   'QUITAR EN GRILLA
   
   GrFromTot.TextMatrix(0, C_CARGO) = Format(vFmt(GrFromTot.TextMatrix(0, C_CARGO)) - Cargo, NUMFMT)
   GrFromTot.TextMatrix(0, C_ABONO) = Format(vFmt(GrFromTot.TextMatrix(0, C_ABONO)) - Abono, NUMFMT)
   GrFromTot.TextMatrix(0, C_SALDO) = Format(IniFrom + vFmt(GrFromTot.TextMatrix(0, C_ABONO)) - vFmt(GrFromTot.TextMatrix(0, C_CARGO)), NEGNUMFMT)
   
   GrFrom.RemoveItem Row
   Call FGrVRows(GrFrom)
      
   ModConcil = True
   Call FrmEnable(False)
      
End Sub
Private Sub FrmEnable(ByVal bEnable As Boolean)
   
   If ModConcil = True Then
   
      Bt_Buscar.Enabled = False
      Cb_Cuentas.Enabled = False
      Tx_Desde.Enabled = False
      Tx_Hasta.Enabled = False
      Bt_Fecha(0).Enabled = False
      Bt_Fecha(1).Enabled = False
      
      Cb_CartBanco.Enabled = False
      Cb_Cartola.Enabled = False
      Bt_ImpCart.Enabled = False
      
   Else
      
      Bt_Buscar.Enabled = Not bEnable
      Bt_OK.Enabled = bEnable
      Bt_ToLeft.Enabled = bEnable
      Bt_ToRight.Enabled = bEnable
      
   End If

   
End Sub
Private Sub SaveGrid(Grid As Control, ByVal bConciliado As Boolean)
   Dim Row As Integer
   Dim Q1 As String
   
   For Row = 1 To Grid.rows - 1
      If Trim(Grid.TextMatrix(Row, C_FECHA)) <> "" And Trim(Grid.TextMatrix(Row, C_IDMOV)) <> "" And Grid.RowHeight(Row) > 0 Then
         Q1 = "UPDATE MovComprobante SET "
         If bConciliado Then
            Q1 = Q1 & " idCartola=" & Val(Grid.TextMatrix(Row, C_IDCART))
         Else
            Q1 = Q1 & " idCartola=0"
         End If
         
         Q1 = Q1 & " WHERE idMov=" & Val(Grid.TextMatrix(Row, C_IDMOV))
        
         If Year(vFmtTxtDate(Grid.TextMatrix(Row, C_FECHA))) = gEmpresa.Ano Then
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Call ExecSQL(DbMain, Q1)
            
            '3376884
            Call SeguimientoMovComprobante(0, gEmpresa.id, gEmpresa.Ano, "FrmConciliacion.SaveGrid1", Q1, 1, "WHERE idMov=" & Val(Grid.TextMatrix(Row, C_IDMOV)) & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano, 1, 2)
            'fin 3376884
            
#If DATACON = 1 Then
         ElseIf Not lDbAnoAnt Is Nothing Then   'es del año anterior
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1
            If gEmprSeparadas Then
               Call ExecSQL(lDbAnoAnt, Q1)
            Else
               Call ExecSQL(DbMain, Q1)
               '3376884
                Call SeguimientoMovComprobante(0, gEmpresa.id, gEmpresa.Ano, "FrmConciliacion.SaveGrid2", Q1, 1, "WHERE idMov=" & Val(Grid.TextMatrix(Row, C_IDMOV)) & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1, 1, 2)
               'fin 3376884

            End If
#Else
         Else
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1
            Call ExecSQL(DbMain, Q1)
            
            '3376884
            Call SeguimientoMovComprobante(0, gEmpresa.id, gEmpresa.Ano, "FrmConciliacion.SaveGrid3", Q1, 1, "WHERE idMov=" & Val(Grid.TextMatrix(Row, C_IDMOV)) & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano - 1, 1, 2)
            'fin 3376884
         
#End If
         End If
         
      End If
      
   Next Row
   
End Sub

Private Function SaldoInicial(Conciliado As Boolean, Grid As Control, Row As Integer, ByVal WhereBas As String) As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim F1 As Long, F2 As Long
   Dim Saldo As Double
   Dim WhereCta As String
   
'   WhereCta = ""
'   If ItemData(Cb_Cuentas) <> 0 Then
'      WhereCta = "MovComprobante.idCuenta=" & ItemData(Cb_Cuentas) & " AND "
'   Else
'      WhereCta = "Atrib" & ATRIB_CONCILIACION & "<> 0 AND "
'   End If
   
'   WhereBas = " WHERE Comprobante.Estado IN (" & EC_APROBADO & ", " & EC_PENDIENTE & ")"
'   WhereBas = WhereConcil & " AND Comprobante.TipoAjuste IN ( " & TAJUSTE_FINANCIERO & ", " & TAJUSTE_AMBOS & ")"
'
'   WhereBas = WhereConcil & " AND" & WhereCta
   
   
   F1 = GetTxDate(Tx_Desde) - 1
   Call FirstLastMonthDay(F1, F1, F2)
         
   Q1 = "SELECT SUM(MovComprobante.Debe) as SumaDebe, Sum(MovComprobante.Haber) as SumaHaber "
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.idCuenta = Cuentas.idCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante")
   Q1 = Q1 & " WHERE " & WhereBas & " AND Fecha <=" & F2
   
   'Esto es porque cuando un windows es en español me pone verdadero o falso
   If Conciliado Then
      Q1 = Q1 & " AND IdCartola <> 0"
   Else
      Q1 = Q1 & " AND IdCartola = 0"
   End If
   
   ERR = 0
      
   If Year(GetTxDate(Tx_Desde)) = gEmpresa.Ano - 1 Then
      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano - 1
      
#If DATACON = 1 Then
      If gEmprSeparadas Then
         Call OpenDbEmp2(lDbAnoAnt, gEmpresa.Rut, gEmpresa.Ano - 1)
         Set Rs = OpenRs(lDbAnoAnt, Q1)
      Else
         Set Rs = OpenRs(DbMain, Q1)
      End If
#Else
      Set Rs = OpenRs(DbMain, Q1)
#End If
     
   Else
      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
   End If
   
   If Rs.EOF = False Then
      Saldo = vFld(Rs("SumaDebe")) - vFld(Rs("SumaHaber"))
   End If
   Call CloseRs(Rs)
   
   If Saldo <> 0 Then
      Grid.rows = Row + 1
      Grid.TextMatrix(Row, C_FECHA) = Format(F2, F_SHORTDATE)
      Grid.TextMatrix(Row, C_SALDO) = Format(Saldo, NEGNUMFMT)
      Grid.TextMatrix(Row, C_GLOSA) = "Saldo inicial"
      
      Row = Row + 1
      Grid.rows = Row + 1
   End If
   
End Function

Private Sub OrdenaPorCol(ByVal Col As Integer)
   Dim LstIdCon As String
   Dim LstIdPorCon As String
   Dim Row As Integer
   
   If Col > C_NUMDOC Then
      Exit Sub
   End If
      
   Gr_Conciliado.Redraw = False
   Gr_PorConcil.Redraw = False
   Me.MousePointer = vbHourglass
   
   'Desmarco  columna Ordenada
   Gr_Conciliado.Row = 0
   Gr_Conciliado.Col = lOrdenSel
   Set Gr_Conciliado.CellPicture = LoadPicture()
   
   Gr_PorConcil.Row = 0
   Gr_PorConcil.Col = lOrdenSel
   Set Gr_PorConcil.CellPicture = LoadPicture()
   
   lOrdenSel = Col
   
   For Row = 1 To Gr_PorConcil.rows - 1
      If Trim(Gr_PorConcil.TextMatrix(Row, C_IDMOV)) <> "" Then
         LstIdPorCon = LstIdPorCon & "," & Gr_PorConcil.TextMatrix(Row, C_IDMOV)
      End If
   Next Row
   
   LstIdPorCon = Mid(LstIdPorCon, 2)
   
   If LstIdPorCon = "" Then
      LstIdPorCon = "0"       'para que no seleccione ninguno
   End If
   
   For Row = 1 To Gr_Conciliado.rows - 1
      If Trim(Gr_Conciliado.TextMatrix(Row, C_IDMOV)) <> "" Then
         LstIdCon = LstIdCon & "," & Gr_Conciliado.TextMatrix(Row, C_IDMOV)
      End If
   Next Row
   
   LstIdCon = Mid(LstIdCon, 2)

   If LstIdCon = "" Then
      LstIdCon = "0"       'para que no seleccione ninguno
   End If

   Call LoadAll(LstIdPorCon, LstIdCon)
      
   Gr_Conciliado.Redraw = True
   Gr_PorConcil.Redraw = True
   
   Me.MousePointer = vbDefault
      
End Sub

Private Sub SetupForm()

   Call FGrSetup(Gr_Cart)

   Gr_Cart.TextMatrix(0, CC_FECHA) = "Fecha"
   Gr_Cart.TextMatrix(0, CC_DETALLE) = "Detalle"
   Gr_Cart.TextMatrix(0, CC_NRODOC) = "Nro. Doc."
   Gr_Cart.TextMatrix(0, CC_CARGO) = "Cargo"
   Gr_Cart.TextMatrix(0, CC_ABONO) = "Abono"

   Gr_Cart.ColWidth(CC_FECHA) = FW_FECHA - 100
   Gr_Cart.ColWidth(CC_HFECHA) = 0
   Gr_Cart.ColWidth(CC_DETALLE) = 4300 + 120
   Gr_Cart.ColWidth(CC_NRODOC) = 1000
   Gr_Cart.ColWidth(CC_CARGO) = 1200
   Gr_Cart.ColWidth(CC_ABONO) = 1200
   Gr_Cart.ColWidth(CC_IDMOV) = 0
   Gr_Cart.ColWidth(CC_IDDET) = 0

   La_Concil.ForeColor = COLOR_VERDEOSCURO

End Sub

Private Function SelMovCartola(ByVal Row As Integer) As Integer
   Dim r As Integer, NroDoc As Double, Cargo As Double, Abono As Double
   Dim i As Integer, n As Integer

   nCartSelMov = -1
   RowCartSel = 0

   n = 0
   r = Row
   If Gr_PorConcil.TextMatrix(r, C_FECHA) = "" Then
      Exit Function
   End If
   
   NroDoc = Val(Right(Trim(Gr_PorConcil.TextMatrix(r, C_NUMDOC)), 6)) Mod 1000
   Cargo = vFmt(Gr_PorConcil.TextMatrix(r, C_CARGO))
   Abono = vFmt(Gr_PorConcil.TextMatrix(r, C_ABONO))

   Gr_Cart.Redraw = False
   For i = Gr_Cart.FixedRows To Gr_Cart.rows - 1
      
      If Gr_Cart.TextMatrix(i, CC_HFECHA) = "" Then
         Exit For
      End If
      
      If CalzaMov(NroDoc, Cargo, Abono, i) Then
      'If Val(Gr_Cart.TextMatrix(i, CC_IDMOV)) = 0 And (NroDoc = 0 Or NroDoc = (Val(Gr_Cart.TextMatrix(i, CC_NRODOC)) Mod 1000)) And Cargo = vFmt(Gr_Cart.TextMatrix(i, CC_CARGO)) And Abono = vFmt(Gr_Cart.TextMatrix(i, CC_ABONO)) Then
         ' destacamos lo que calzan
         Call FGrSetRowVisible(Gr_Cart, i)
         Call FGrSetRowStyle(Gr_Cart, i, "FC", vbBlue)
         RowCartSel = i
         n = n + 1
      ElseIf Val(Gr_Cart.TextMatrix(i, CC_IDMOV)) Then ' ya está conciliado
         Call FGrSetRowStyle(Gr_Cart, i, "FC", COLOR_VERDEOSCURO)
      Else
         Call FGrSetRowStyle(Gr_Cart, i, "FC", vbWindowText)
      End If

   Next i
   Gr_Cart.Redraw = True

   nCartSelMov = n
   SelMovCartola = n
   La_nCalzan = n & " "
   
   'Call FGrVRows(Gr_Cart)
   
End Function

Private Function Conciliar(ByVal Row As Integer, ByVal bMsg As Boolean, Optional ByVal vBtLeft As Boolean = False) As Boolean
   Dim Row2 As Integer

   Conciliar = False
   
   If nCartMov > 0 Then  ' La cartola tiene detalle ?
      If nCartSelMov = 0 Then
         If bMsg Then
            If MsgBox1("El movimiento seleccionado no calza con ningún movimiento de la cartola seleccionada." & vbLf & "En la cartola los que calzan se marcan en azul." & vbLf & vbLf & "¿Desea conciliarlo de todas maneras?", vbQuestion + vbYesNo) = vbNo Then
               Exit Function
            End If
         Else
            Exit Function
         End If
      ElseIf nCartSelMov > 1 Then
         If bMsg Then
            MsgBox1 "El movimiento seleccionado calza con más de un movimiento de la cartola, seleccionelo desde la cartola con un doble-clic." & vbLf & "En la cartola los que calzan se marcan en azul.", vbExclamation
         End If
         Exit Function
      End If
      
   End If
      
   Row2 = FGrAddRow(Gr_Conciliado)
   
   Call ChangeGrid(Row, Gr_PorConcil, Gr_TotPorConc, Row2, Gr_Conciliado, Gr_TotConc, vBtLeft)
   Conciliar = True

End Function
Private Function SetupPriv()
   
   If Not ChkPriv(PRV_ADM_CONCIL) Then
      Call EnableForm(Me, False)
   End If
   
End Function

Private Sub SetUpPrtGrid(Grid As Object, GridTot As Object, Tit As String)
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(1) As String
   
   Printer.Orientation = ORIENT_VER
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Tit
   Titulos(1) = "Fecha: " & Tx_Desde & " al " & Tx_Hasta
   gPrtReportes.Titulos = Titulos
         
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
   For i = 0 To GridTot.Cols - 1
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
      
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_FECHA
   gPrtReportes.NTotLines = 1

End Sub


Private Sub Bt_ConvMoneda_Click()
   Dim Frm As FrmConverMoneda
   Dim Col As Integer
   Dim Row As Integer
   Dim Valor As Double
   
'   Col = Grid.Col
'   Row = Grid.Row
'
'   If Col = C_VALOR Then
'      Valor = vFmt(Grid.TextMatrix(Row, Col))
'   End If
   
   Set Frm = New FrmConverMoneda
   Frm.FSelect (Valor)
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


Private Sub CleanCartList()
   Dim c As Integer
   
   Gr_TotConc.Clear
   
   Gr_Conciliado.rows = 1
   Call FGrVRows(Gr_Conciliado)
   
   Gr_Cart.rows = 1
   Call FGrVRows(Gr_Cart)

   Bt_AutoConcil.Enabled = False

End Sub

Private Function CalzaMov(ByVal NroDoc As Double, ByVal Cargo As Double, ByVal Abono As Double, ByVal RowCart As Integer) As Boolean

'3340743 se realiza la equivalencia de MOD 1000 esto ya que con mod se desbordaba
   'If Val(Gr_Cart.TextMatrix(RowCart, CC_IDMOV)) = 0 And (NroDoc = 0 Or NroDoc = (Val(Gr_Cart.TextMatrix(RowCart, CC_NRODOC)) Mod 1000)) And Cargo = vFmt(Gr_Cart.TextMatrix(RowCart, CC_CARGO)) And Abono = vFmt(Gr_Cart.TextMatrix(RowCart, CC_ABONO)) Then
   If Val(Gr_Cart.TextMatrix(RowCart, CC_IDMOV)) = 0 And (NroDoc = 0 Or NroDoc = Val(Gr_Cart.TextMatrix(RowCart, CC_NRODOC)) - (1000 * Fix(Val(Gr_Cart.TextMatrix(RowCart, CC_NRODOC)) / 1000))) And Cargo = vFmt(Gr_Cart.TextMatrix(RowCart, CC_CARGO)) And Abono = vFmt(Gr_Cart.TextMatrix(RowCart, CC_ABONO)) Then
 '3340743
      CalzaMov = True
   Else
      CalzaMov = False
   End If

End Function

Private Sub CorrigeCart()
   Dim Q1 As String, Rs As Recordset, Rc As Long, nRec As Long
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String

   ' Quiebres de enlaces
   
   'Q1 = "SELECT DetCartola.IdCartola, DetCartola.IdMov, MovComprobante.IdCartola, Comprobante.Estado, DetCartola.Cargo, DetCartola.Abono, MovComprobante.Haber, MovComprobante.Debe"
   Q1 = "SELECT Count(*) as N"
   Q1 = Q1 & " FROM ((Cartola INNER JOIN DetCartola ON Cartola.IdCartola = DetCartola.IdCartola "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cartola", "DetCartola") & " )"
   Q1 = Q1 & " INNER JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
   Q1 = Q1 & JoinEmpAno(gDbType, "DetCartola", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE MovComprobante.IdCartola=0 and DetCartola.Cargo = MovComprobante.Haber and DetCartola.Abono = MovComprobante.Debe"
   Q1 = Q1 & " AND Cartola.IdEmpresa = " & gEmpresa.id & " AND Cartola.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   nRec = vFld(Rs("N"))
   Call CloseRs(Rs)

   If nRec > 0 Then
      Call AddLog("Conc: " & gEmpresa.Rut & " - " & gEmpresa.Ano & ": corrige " & nRec & " quiebres de enlace.")
   
      ' corrige enlace
'      Q1 = "UPDATE ((Cartola INNER JOIN DetCartola ON Cartola.IdCartola = DetCartola.IdCartola "
'      Q1 = Q1 & " AND Cartola.IdEmpresa = DetCartola.IdEmpresa AND Cartola.Ano = DetCartola.Ano )"
'      Q1 = Q1 & " INNER JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
'      Q1 = Q1 & " AND DetCartola.IdEmpresa = MovComprobante.IdEmpresa AND DetCartola.Ano = MovComprobante.Ano) "
'      Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
'      Q1 = Q1 & " AND Comprobante.IdEmpresa = MovComprobante.IdEmpresa AND Comprobante.Ano = MovComprobante.Ano "
'      Q1 = Q1 & " SET MovComprobante.IdCartola = DetCartola.IdCartola"
'      Q1 = Q1 & " WHERE MovComprobante.IdCartola=0 and DetCartola.Cargo = MovComprobante.Haber and DetCartola.Abono = MovComprobante.Debe"
'      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
'
'      Rc = ExecSQL(DbMain, Q1)
      

      Tbl = " Cartola "
      sFrom = " ((Cartola INNER JOIN DetCartola ON Cartola.IdCartola = DetCartola.IdCartola "
      sFrom = sFrom & JoinEmpAno(gDbType, "Cartola", "DetCartola") & " )"
      sFrom = sFrom & " INNER JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
      sFrom = sFrom & JoinEmpAno(gDbType, "DetCartola", "MovComprobante") & " )"
      sFrom = sFrom & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
      sFrom = sFrom & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      sSet = " MovComprobante.IdCartola = DetCartola.IdCartola"
      sWhere = " WHERE MovComprobante.IdCartola=0 and DetCartola.Cargo = MovComprobante.Haber and DetCartola.Abono = MovComprobante.Debe"
      sWhere = sWhere & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano

      Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   End If

   
   ' Apuntan a movimientos inexistentes
   'Q1 = "SELECT DetCartola.IdDetCartola, DetCartola.IdMov, MovComprobante.IdMov"
   Q1 = "SELECT Count(*) as N"
   Q1 = Q1 & " FROM DetCartola LEFT JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
   Q1 = Q1 & JoinEmpAno(gDbType, "DetCartola", "MovComprobante")
   Q1 = Q1 & " WHERE MovComprobante.IdMov IS NULL AND DetCartola.idMov > 0"
   Q1 = Q1 & " AND DetCartola.IdEmpresa = " & gEmpresa.id & " AND DetCartola.Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   nRec = vFld(Rs("N"))
   Call CloseRs(Rs)

   If nRec > 0 Then
      Call AddLog("Conc: " & gEmpresa.Rut & " - " & gEmpresa.Ano & ": corrige " & nRec & " enlaces a mov eliminados.")
   
'      Q1 = "UPDATE DetCartola LEFT JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
'      Q1 = Q1 & " AND DetCartola.IdEmpresa = MovComprobante.IdEmpresa AND DetCartola.Ano = MovComprobante.Ano "
'      Q1 = Q1 & " SET DetCartola.idMov = 0"
'      Q1 = Q1 & " WHERE MovComprobante.IdMov IS NULL AND DetCartola.idMov <> 0"
'      Q1 = Q1 & " AND DetCartola.IdEmpresa = " & gEmpresa.id & " AND DetCartola.Ano = " & gEmpresa.Ano
'
'      Rc = ExecSQL(DbMain, Q1)

      Tbl = " DetCartola "
      sFrom = " DetCartola LEFT JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
      sFrom = sFrom & JoinEmpAno(gDbType, "DetCartola", "MovComprobante")
      sSet = " DetCartola.idMov = 0"
      sWhere = " WHERE MovComprobante.IdMov IS NULL AND DetCartola.idMov <> 0"
      sWhere = sWhere & " AND DetCartola.IdEmpresa = " & gEmpresa.id & " AND DetCartola.Ano = " & gEmpresa.Ano

      Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   End If

   ' Apuntan a comprobantes anulados

   'Q1 = "SELECT DetCartola.IdDetCartola, DetCartola.IdMov, Comprobante.Estado"
   Q1 = "SELECT Count(*) as N"
   Q1 = Q1 & " FROM (DetCartola INNER JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
   Q1 = Q1 & JoinEmpAno(gDbType, "DetCartola", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Comprobante.Estado= " & EC_ANULADO
   Q1 = Q1 & " AND DetCartola.IdEmpresa = " & gEmpresa.id & " AND DetCartola.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   nRec = vFld(Rs("N"))
   Call CloseRs(Rs)

   If nRec Then
      Call AddLog("Conc: " & gEmpresa.Rut & " - " & gEmpresa.Ano & ": corrige " & nRec & " enlaces a mov anulados.")

'      Q1 = "UPDATE (DetCartola INNER JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
'      Q1 = Q1 & " AND DetCartola.IdEmpresa = MovComprobante.IdEmpresa AND DetCartola.Ano = MovComprobante.Ano )"
'      Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
'      Q1 = Q1 & " AND Comprobante.IdEmpresa = MovComprobante.IdEmpresa AND Comprobante.Ano = MovComprobante.Ano"
'      Q1 = Q1 & " Set DetCartola.idMov = 0"
'      Q1 = Q1 & " WHERE Comprobante.Estado= " & EC_ANULADO
'      Q1 = Q1 & " AND DetCartola.IdEmpresa = " & gEmpresa.id & " AND DetCartola.Ano = " & gEmpresa.Ano
'      Rc = ExecSQL(DbMain, Q1)
      
      Tbl = " DetCartola "
      sFrom = " (DetCartola INNER JOIN MovComprobante ON DetCartola.IdMov = MovComprobante.IdMov "
      sFrom = sFrom & JoinEmpAno(gDbType, "DetCartola", "MovComprobante") & " )"
      sFrom = sFrom & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
      sFrom = sFrom & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      sSet = " DetCartola.idMov = 0"
      sWhere = " WHERE Comprobante.Estado= " & EC_ANULADO
      sWhere = sWhere & " AND DetCartola.IdEmpresa = " & gEmpresa.id & " AND DetCartola.Ano = " & gEmpresa.Ano
      Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   End If

End Sub
