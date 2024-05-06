VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCalcRazones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Razones Financieras"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "FrmCalcRazones.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11835
      Begin VB.CommandButton Bt_Print 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Picture         =   "FrmCalcRazones.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10560
         TabIndex        =   14
         Top             =   180
         Width           =   1095
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
         Left            =   660
         Picture         =   "FrmCalcRazones.frx":04C6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_CopyExcel 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Picture         =   "FrmCalcRazones.frx":096D
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Copiar Excel"
         Top             =   180
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
         Left            =   3480
         Picture         =   "FrmCalcRazones.frx":0DB2
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   2640
         Picture         =   "FrmCalcRazones.frx":11DB
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   3060
         Picture         =   "FrmCalcRazones.frx":1579
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Calculadora"
         Top             =   180
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
         Left            =   2100
         Picture         =   "FrmCalcRazones.frx":18DA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_EditRazon 
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
         Left            =   120
         Picture         =   "FrmCalcRazones.frx":197E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Detalle razón financiera seleccionada"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   660
      Width           =   11835
      Begin VB.CheckBox Ch_RazFinConfiguradas 
         Caption         =   "Sólo razones financieras configuradas"
         Height          =   255
         Left            =   4380
         TabIndex        =   22
         Top             =   720
         Width           =   3075
      End
      Begin VB.CheckBox Ch_Estado 
         Caption         =   "Sólo comprobantes Aprobados"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   720
         Width           =   2475
      End
      Begin VB.CommandButton Bt_CalcRazFin 
         Caption         =   "Calcular"
         Height          =   705
         Left            =   10560
         Picture         =   "FrmCalcRazones.frx":1D90
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   4380
         TabIndex        =   1
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   1
         Left            =   7260
         Picture         =   "FrmCalcRazones.frx":2313
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   230
      End
      Begin VB.CommandButton Bt_Fecha 
         Height          =   315
         Index           =   0
         Left            =   5400
         Picture         =   "FrmCalcRazones.frx":261D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   230
      End
      Begin VB.ComboBox Cb_Tipo 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   3840
         TabIndex        =   19
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   7
         Left            =   5760
         TabIndex        =   18
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo razón:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5415
      Left            =   0
      TabIndex        =   15
      Top             =   1980
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   20
      Cols            =   9
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
   End
End
Attribute VB_Name = "FrmCalcRazones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_OBLIGATORIA = 0
Const C_IDRAZON = 1
Const C_NOMBRE = 2
Const C_FORMULA = 3
Const C_VALORES = 4
Const C_IGUAL = 5
Const C_TOTAL = 6
Const C_UNIDAD = 7
Const C_CANTDIAS = 8
Const C_OPERADOR = 9
Const C_FMT = 10

Const NCOLS = C_FMT

Const DIVCERO = -2

Private Sub Bt_CalcRazFin_Click()
   Dim F1 As Long, F2 As Long
   
   F1 = GetTxDate(Tx_Desde)
   F2 = GetTxDate(Tx_Hasta)
   
   If F1 > F2 Then
      MsgBeep vbExclamation
      Tx_Hasta.SetFocus
      Exit Sub
      
   End If
   
   MousePointer = vbHourglass
   Call LoadAll
   
   MousePointer = vbDefault
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   
   If Bt_CalcRazFin.Enabled = True Then
      MsgBox1 "Presione el botón Calcular antes de copiar.", vbExclamation
      Exit Sub
   End If
   
   Call FGr2Clip(Grid, "Fecha Inicio: " & Tx_Desde & " Fecha Término: " & Tx_Hasta)
End Sub

Private Sub Bt_EditRazon_Click()
   Dim Frm As FrmParamRaz
   
   If Grid.Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Grid.Row, C_IDRAZON)) <= 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmParamRaz
   If Frm.FEdit(Val(Grid.TextMatrix(Grid.Row, C_IDRAZON))) = vbOK Then
      Call Bt_CalcRazFin_Click
   End If

   Set Frm = Nothing

End Sub

Private Sub Cb_Tipo_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_Estado_Click()
   Call EnableFrm(True)

End Sub

Private Sub Ch_RazFinConfiguradas_Click()
   Call EnableFrm(True)

End Sub

Private Sub Form_Load()
   Dim i As Integer

   Call SetUpGrid
      
   For i = 0 To UBound(gTipoRazFin)
   
      If gTipoRazFin(i).id = 0 Or gTipoRazFin(i).Nombre = "" Then
         Exit For
      End If
      
      Call AddItem(Cb_Tipo, gTipoRazFin(i).Nombre, gTipoRazFin(i).id)
      
   Next i
   
   If Cb_Tipo.ListCount > 0 Then
      Cb_Tipo.ListIndex = 0
   End If

   Call SetTxDate(Tx_Desde, DateSerial(gEmpresa.Ano, 1, 1))
   Call SetTxDate(Tx_Hasta, DateSerial(gEmpresa.Ano, 12, 31))
   
   Ch_Estado = 1
   
   Call LoadAll

End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
   
   Grid.ColWidth(C_IDRAZON) = 0
   Grid.ColWidth(C_NOMBRE) = 3000
   Grid.ColWidth(C_FORMULA) = 3800
   Grid.ColWidth(C_VALORES) = 1700
   Grid.ColWidth(C_IGUAL) = 300
   Grid.ColWidth(C_TOTAL) = 1760
   Grid.ColWidth(C_UNIDAD) = 800
   Grid.ColWidth(C_CANTDIAS) = 0
   Grid.ColWidth(C_OPERADOR) = 0
  
   Grid.ColWidth(C_OBLIGATORIA) = 0
   Grid.ColWidth(C_FMT) = 0
   
   Grid.ColAlignment(C_NOMBRE) = flexAlignLeftCenter
   Grid.ColAlignment(C_FORMULA) = flexAlignCenterCenter
   Grid.ColAlignment(C_VALORES) = flexAlignCenterCenter
   Grid.ColAlignment(C_IGUAL) = flexAlignCenterCenter
   Grid.ColAlignment(C_TOTAL) = flexAlignRightCenter
   Grid.ColAlignment(C_UNIDAD) = flexAlignLeftCenter
      
End Sub
Private Function LoadAll()
   Dim i As Integer
   Dim WhereFecha As String
   Dim WhereTipo As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim F1 As Long
   Dim F2 As Long
   Dim Tipo As Integer
   Dim IdRazon As Long
   Dim Numerador As Double
   Dim Denominador As Double
   Dim Total As Double
   Dim Msg As Boolean
   Dim RcCalc As Integer
   Dim NCtas As Long
   Dim RsCtas As Recordset
   
   Grid.Redraw = False
   Me.MousePointer = vbHourglass
   
   F1 = GetTxDate(Tx_Desde)
   F2 = GetTxDate(Tx_Hasta)
   
   WhereFecha = " Comprobante.Fecha BETWEEN " & F1 & " AND " & F2
   WhereTipo = " Tipo = " & ItemData(Cb_Tipo)
   
   Q1 = "SELECT RazonesFin.IdRazon, Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, ParamRazon.CantDias "
   Q1 = Q1 & " FROM RazonesFin "
   Q1 = Q1 & " LEFT JOIN ParamRazon ON RazonesFin.IdRazon = ParamRazon.IdRazon "
   Q1 = Q1 & " WHERE " & WhereTipo
   Q1 = Q1 & " ORDER BY Tipo, Nombre, RazonesFin.IdRazon "
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   Tipo = 0
   
   Do While Rs.EOF = False
   
      NCtas = 0
   
      If Ch_RazFinConfiguradas <> 0 Then
         Q1 = "SELECT Count(*) as NCtas FROM CuentasRazon WHERE IdRazon = " & vFld(Rs("IdRazon"))
         Set RsCtas = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            NCtas = vFld(RsCtas("NCtas"))
         End If
         
         Call CloseRs(RsCtas)
      End If
      
      If Ch_RazFinConfiguradas = 0 Or (Ch_RazFinConfiguradas <> 0 And NCtas > 0) Then
            
         Grid.rows = Grid.rows + 1
         
         If vFld(Rs("Tipo")) <> Tipo Then
            Grid.TextMatrix(i, C_NOMBRE) = UCase(GetTipoRazFin(vFld(Rs("Tipo"))))
            Grid.TextMatrix(i, C_FMT) = "B"
            Call FGrSetRowStyle(Grid, i, "B")
            Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
            Grid.rows = Grid.rows + 2
            Grid.TextMatrix(i + 1, C_OBLIGATORIA) = "O"
            Grid.TextMatrix(i + 2, C_OBLIGATORIA) = "O"
            i = i + 2
            Tipo = vFld(Rs("Tipo"))
         End If
         
         Grid.TextMatrix(i, C_IDRAZON) = vFld(Rs("IdRazon"))
         Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"))
         Grid.TextMatrix(i, C_FORMULA) = String(5, " ") & vFld(Rs("TxtNumerador")) & String(5, " ")
         Grid.TextMatrix(i, C_OPERADOR) = vFld(Rs("Operador"))
         Grid.TextMatrix(i, C_IGUAL) = "="
         Grid.TextMatrix(i, C_TOTAL) = ""
         Grid.TextMatrix(i, C_UNIDAD) = vFld(Rs("UnidadRes"))
         
         If LCase(Grid.TextMatrix(i, C_UNIDAD)) = "dias" Or LCase(Grid.TextMatrix(i, C_UNIDAD)) = "días" Then
            Grid.TextMatrix(i, C_CANTDIAS) = IIf(vFld(Rs("CantDias")) = 0, 365, vFld(Rs("CantDias")))
         End If
         
         If vFld(Rs("operador")) = "/" Then  'operador típico
            Grid.Col = C_FORMULA
            Grid.Row = i
            Grid.CellFontUnderline = True
            Grid.rows = Grid.rows + 1
            Grid.TextMatrix(i + 1, C_FMT) = "L(" & C_FORMULA & "," & C_FORMULA & ")"
            Grid.TextMatrix(i + 1, C_FORMULA) = IIf(vFld(Rs("TxtDenominador")) <> "", vFld(Rs("TxtDenominador")), 1)
            Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
            Grid.TextMatrix(i + 1, C_OBLIGATORIA) = "O"
            i = i + 1
'         Else
'            Grid.TextMatrix(i, C_FORMULA) = Trim(Grid.TextMatrix(i, C_FORMULA)) & " " & vFld(Rs("operador"), True) & " " & vFld(Rs("TxtDenominador"), True)
'            Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
         End If
         
         If vFld(Rs("TxtDenominador")) = "" Then   'borramos las cuentas del denominador, si las hay, cuando no hay denominador
            Q1 = "DELETE * FROM CuentasRazon "
            Q1 = Q1 & " WHERE IdRazon = " & vFld(Rs("IdRazon"))
            Q1 = Q1 & " AND NumDenom = " & CTA_DENOMINADOR
            
            Call ExecSQL(DbMain, Q1)
         End If
         
         Rs.MoveNext
         
         Grid.rows = Grid.rows + 2
         Grid.TextMatrix(i + 1, C_OBLIGATORIA) = "O"
         Grid.TextMatrix(i + 2, C_OBLIGATORIA) = "O"
         
         i = i + 2
         
      Else
      
         Rs.MoveNext
         
      End If
      
   Loop
   
   Call CloseRs(Rs)
   
   IdRazon = 0
      
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Val(Grid.TextMatrix(i, C_IDRAZON)) <> 0 Then
   
         IdRazon = Val(Grid.TextMatrix(i, C_IDRAZON))
         RcCalc = CalcIdRazon(IdRazon, WhereFecha, Grid.TextMatrix(i, C_OPERADOR), Grid.TextMatrix(i, C_UNIDAD), Numerador, Denominador, Total, vFmt(Grid.TextMatrix(i, C_CANTDIAS)))
         
         If RcCalc <> 0 Then
            'Grid.TextMatrix(i, C_FORMULA) = String(5, " ") & Trim(Grid.TextMatrix(i, C_FORMULA)) & " = " & Format(Numerador, NUMFMT) & String(5, " ")
            Grid.TextMatrix(i, C_FORMULA) = String(5, " ") & Trim(Grid.TextMatrix(i, C_FORMULA)) & String(5, " ")
            Grid.TextMatrix(i, C_VALORES) = String(5, " ") & Format(Numerador, NUMFMT) & String(5, " ")
            
            'Grid.TextMatrix(i + 1, C_FORMULA) = Grid.TextMatrix(i + 1, C_FORMULA) & " = " & Format(Denominador, NUMFMT)
            Grid.TextMatrix(i + 1, C_FORMULA) = Grid.TextMatrix(i + 1, C_FORMULA)
            Grid.TextMatrix(i + 1, C_VALORES) = Format(Denominador, NUMFMT)
            
            Grid.Col = C_VALORES
            Grid.Row = i
            Grid.CellFontUnderline = True
            Grid.TextMatrix(i + 1, C_FMT) = "L(" & C_FORMULA & "," & C_VALORES & ")"

            If RcCalc = DIVCERO Then
               Grid.TextMatrix(i, C_TOTAL) = "Indefinido"
            Else
               Grid.TextMatrix(i, C_TOTAL) = Format(Total, DBLFMT4)
            End If
            
         Else
            'Call MsgBox1("La razón financiera """ & Grid.TextMatrix(i, C_NOMBRE) & """ no está bien definida. Posiblemente falta definir las cuentas que intervienen en su cálculo.", vbExclamation + vbOKOnly)
            If Not Msg Then  'el mensaje se muestra una sola vez
               Call MsgBox1("Hay razones financieras que no están bien configuradas." & vbCrLf & "Posiblemente falta definir las cuentas que intervienen en su cálculo." & vbCrLf & vbCrLf & "Utilice la opción 'Configurar Razones Financieras', bajo el menú Configuración.", vbExclamation + vbOKOnly)
               Msg = True
            End If
         End If
         
      End If
      
   Next i
   
   Call EnableFrm(False)
   Grid.rows = Grid.rows + 1 'por si no tiene lìneas
   Call FGrVRows(Grid)
   Grid.rows = Grid.rows + 2
   Grid.Redraw = True
   Me.MousePointer = vbDefault
   
End Function

Private Sub tx_Desde_Change()
   Call EnableFrm(True)
End Sub

Private Sub Tx_Desde_GotFocus()
   Call DtGotFocus(Tx_Desde)
End Sub

Private Sub Tx_Desde_LostFocus()
   
   If Trim$(Tx_Desde) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Desde)
   
End Sub

Private Sub Tx_Desde_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub

Private Sub tx_Hasta_Change()
   Call EnableFrm(True)
End Sub

Private Sub Tx_Hasta_GotFocus()
   Call DtGotFocus(Tx_Hasta)
   
End Sub


Private Sub Tx_Hasta_LostFocus()
   
   If Trim$(Tx_Hasta) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_Hasta)
      
End Sub

Private Sub Tx_Hasta_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub EnableFrm(bool As Boolean)
   Bt_CalcRazFin.Enabled = bool
   
End Sub
Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid, C_TOTAL)
   
   Set Frm = Nothing

End Sub
Private Sub Bt_ConvMoneda_Click()
   Dim Frm As FrmConverMoneda
   Dim Valor As Double
      
   Set Frm = New FrmConverMoneda
   Frm.FView (Valor)
      
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
Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - 500
   Grid.Width = Me.Width - 230
   
   Call FGrVRows(Grid)
   
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
   
   Call EnableFrm(True)

End Sub

Private Function CalcIdRazon(ByVal IdRazon As Long, ByVal WhFecha As String, ByVal Operador As String, ByVal Unidad As String, Numerador As Double, Denominador As Double, Total As Double, Optional ByVal CantDias As Long = 365) As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim CodCuentasSuma(CTA_DENOMINADOR) As String
   Dim CodCuentasResta(CTA_DENOMINADOR) As String
   Dim StrEstado As String
   Dim WhAjuste As String
   
   
   CalcIdRazon = False
   Numerador = 0
   Denominador = 0
   Total = 0
   
   If Ch_Estado <> 0 Then
      StrEstado = " AND Comprobante.Estado = " & EC_APROBADO
   Else
      StrEstado = " AND Comprobante.Estado IN (" & EC_PENDIENTE & "," & EC_APROBADO & ")"
   End If
   
   WhAjuste = " AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
   
   Q1 = "SELECT CuentasRazon.NumDenom, CuentasRazon.CodCuenta, Cuentas.IdCuenta, CuentasRazon.Operador "
   Q1 = Q1 & " FROM CuentasRazon INNER JOIN Cuentas ON CuentasRazon.CodCuenta = Cuentas.Codigo "
   Q1 = Q1 & " WHERE IdRazon=" & IdRazon
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      If vFld(Rs("Operador"), True) = "+" Then
         CodCuentasSuma(vFld(Rs("NumDenom"))) = CodCuentasSuma(vFld(Rs("NumDenom"))) & " OR " & GenWhereCuentas(FmtCodCuenta(vFld(Rs("CodCuenta"))))
      ElseIf vFld(Rs("Operador"), True) = "-" Then
         CodCuentasResta(vFld(Rs("NumDenom"))) = CodCuentasResta(vFld(Rs("NumDenom"))) & " OR " & GenWhereCuentas(FmtCodCuenta(vFld(Rs("CodCuenta"))))
      End If
      
      Rs.MoveNext
      
   Loop

   Call CloseRs(Rs)
              
   If CodCuentasSuma(CTA_NUMERADOR) <> "" Then
      CodCuentasSuma(CTA_NUMERADOR) = Mid(CodCuentasSuma(CTA_NUMERADOR), 5)
   Else
      Exit Function
   End If
   
   If CodCuentasResta(CTA_NUMERADOR) <> "" Then
      CodCuentasResta(CTA_NUMERADOR) = Mid(CodCuentasResta(CTA_NUMERADOR), 5)
   End If
   
   If CodCuentasSuma(CTA_DENOMINADOR) <> "" Then
      CodCuentasSuma(CTA_DENOMINADOR) = Mid(CodCuentasSuma(CTA_DENOMINADOR), 5)
   End If

   If CodCuentasResta(CTA_DENOMINADOR) <> "" Then
      CodCuentasResta(CTA_DENOMINADOR) = Mid(CodCuentasResta(CTA_DENOMINADOR), 5)
   End If

   Q1 = "SELECT Sum(MovComprobante.Debe - MovComprobante.Haber) as Tot "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta"
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE " & WhFecha & " AND (" & CodCuentasSuma(CTA_NUMERADOR) & ")"
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & StrEstado & WhAjuste
   
   Set Rs = OpenRs(DbMain, Q1)
      
   If Rs.EOF = False Then
      Numerador = vFld(Rs("Tot"))
   End If
   
   Call CloseRs(Rs)
   
   If CodCuentasResta(CTA_NUMERADOR) <> "" Then
   
      Q1 = "SELECT Sum(MovComprobante.Debe-MovComprobante.Haber) as Tot "
      Q1 = Q1 & " FROM (MovComprobante INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
      Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
      Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      Q1 = Q1 & " WHERE " & WhFecha & " AND (" & CodCuentasResta(CTA_NUMERADOR) & ")"
      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
      Q1 = Q1 & StrEstado & WhAjuste
      
      Set Rs = OpenRs(DbMain, Q1)
         
      If Rs.EOF = False Then
         Numerador = Numerador - vFld(Rs("Tot"))
      End If
      
      Call CloseRs(Rs)
      
   End If
    
   Numerador = Abs(Numerador)

   Denominador = 1
   
   If CodCuentasSuma(CTA_DENOMINADOR) <> "" Then
      Q1 = "SELECT Sum(MovComprobante.Debe-MovComprobante.Haber) as Tot "
      Q1 = Q1 & " FROM (MovComprobante INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
      Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
      Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      Q1 = Q1 & " WHERE " & WhFecha & " AND (" & CodCuentasSuma(CTA_DENOMINADOR) & ")"
      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
      Q1 = Q1 & StrEstado & WhAjuste
      
      Set Rs = OpenRs(DbMain, Q1)
         
      If Rs.EOF = False Then
         Denominador = vFld(Rs("Tot"))
      End If
      
      Call CloseRs(Rs)
   
   End If
   
   
   If CodCuentasResta(CTA_DENOMINADOR) <> "" Then
   
      Q1 = "SELECT Sum(MovComprobante.Debe-MovComprobante.Haber) as Tot "
      Q1 = Q1 & " FROM (MovComprobante INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
      Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
      Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      Q1 = Q1 & " WHERE " & WhFecha & " AND (" & CodCuentasResta(CTA_DENOMINADOR) & ")"
      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
      Q1 = Q1 & StrEstado & WhAjuste
      
      Set Rs = OpenRs(DbMain, Q1)
         
      If Rs.EOF = False Then
         Denominador = Denominador - vFld(Rs("Tot"))
      End If
      
      Call CloseRs(Rs)
      
   End If
      
   Denominador = Abs(Denominador)
      
   If Denominador <> 0 Then
      Total = Numerador / Denominador
      CalcIdRazon = True
   Else
      Total = 0
      CalcIdRazon = DIVCERO    'inválido
      Exit Function
   End If

   If LCase(Trim(Unidad)) = "dias" Or LCase(Trim(Unidad)) = "días" Then
      Total = Total * CantDias
   End If
   
   
End Function

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   
   If Bt_CalcRazFin.Enabled = True Then
      MsgBox1 "Presione el botón Calcular antes de seleccionar la vista previa.", vbExclamation
      Exit Sub
   End If
   
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   Call ResetPrtBas(gPrtReportes)
   
End Sub
Private Sub Bt_Print_Click()
   Dim PrtOrient As Integer
   
   If Bt_CalcRazFin.Enabled = True Then
      MsgBox1 "Presione el botón Calcular antes de imprimir.", vbExclamation
      Exit Sub
   End If
   
   Call SetUpPrtGrid
   
   PrtOrient = Printer.Orientation
   Printer.Orientation = ORIENT_VER
   
   Me.MousePointer = vbHourglass
   
   Call gPrtReportes.PrtFlexGrid(Printer)
   
   Me.MousePointer = vbDefault
   
   Printer.Orientation = PrtOrient
   
   Call ResetPrtBas(gPrtReportes)
   
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS)
   Dim Titulos(1) As String
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Me.Caption
   Titulos(1) = "Año " & gEmpresa.Ano
   gPrtReportes.Titulos = Titulos
            
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
     
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_OBLIGATORIA
   
   gPrtReportes.Total = Total
   gPrtReportes.NTotLines = 0
   
   gPrtReportes.FmtCol = C_FMT
   
End Sub

