VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCapitalPropio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capital Propio Tributario"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   Icon            =   "FrmCapitalPropio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   9975
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tx_TotCapPropio 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
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
      Height          =   315
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox Tx_CapPropio 
      BackColor       =   &H8000000F&
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
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Capital Propio Tributario"
      Top             =   6360
      Width           =   7815
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   9855
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
         Left            =   540
         Picture         =   "FrmCapitalPropio.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   8640
         TabIndex        =   8
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
         Left            =   120
         Picture         =   "FrmCapitalPropio.frx":04C6
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Left            =   960
         Picture         =   "FrmCapitalPropio.frx":096D
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   2880
         Picture         =   "FrmCapitalPropio.frx":0DB2
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   2040
         Picture         =   "FrmCapitalPropio.frx":11DB
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   2460
         Picture         =   "FrmCapitalPropio.frx":1579
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   1500
         Picture         =   "FrmCapitalPropio.frx":18DA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9763
      _Version        =   393216
      Rows            =   20
      Cols            =   6
      FixedRows       =   0
   End
End
Attribute VB_Name = "FrmCapitalPropio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_DESC = 0
Const C_VALDET = 1
Const C_MENOS = 2
Const C_TOTAL = 3
Const C_FMT = 4
Const C_OBLIGATORIA = 5

Const NCOLS = C_OBLIGATORIA

'3042010
Public capitalEfectivo As Double
'3042010

Private Sub bt_Cerrar_Click()
   Dim Q1 As String
   Dim Rs As Recordset

   'guardamos el total capital propio
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'CAPPROPIO' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)

   If Rs.EOF = False Then    'ya existe, lo actualizamos
      Q1 = "UPDATE ParamEmpresa SET Valor = '" & vFmt(Tx_TotCapPropio) & "' WHERE Tipo = 'CAPPROPIO' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Else
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES( 'CAPPROPIO', 0, '" & vFmt(Tx_TotCapPropio) & "', " & gEmpresa.id & "," & gEmpresa.Ano & ")"
   End If

   Call ExecSQL(DbMain, Q1)
   Call CloseRs(Rs)
   
   'guardamos el total capital propio en tabla EmpresasAno
   Q1 = "UPDATE EmpresasAno SET CPS_CapPropioTrib = " & vFmt(Tx_TotCapPropio)
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Call LP_FGr2Clip(Grid, Me.Caption)

End Sub

Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid)
   
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

Private Sub Form_Activate()
   MsgBox1 "ATENCIÓN:" & vbNewLine & vbNewLine & "Este informe se genera seleccionando solamente los comprobantes en estado APROBADO.", vbInformation
End Sub

Private Sub Form_Load()

   Me.Caption = Me.Caption & " al 31 de Diciembre " & gEmpresa.Ano & " (según normas art. 2 N° 10 Ley de Renta)"
   Tx_CapPropio = Me.Caption
   
   Call SetUpGrid
   
   Call LoadAll
   
End Sub

Private Sub SetUpGrid()
   
   Grid.ColWidth(C_DESC) = 4400
   Grid.ColWidth(C_VALDET) = 1700
   Grid.ColWidth(C_MENOS) = 1700
   Grid.ColWidth(C_TOTAL) = 1700
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_OBLIGATORIA) = 0
   
   Grid.ColAlignment(C_DESC) = flexAlignLeftCenter
   Grid.ColAlignment(C_VALDET) = flexAlignRightCenter
   Grid.ColAlignment(C_MENOS) = flexAlignRightCenter
   Grid.ColAlignment(C_TOTAL) = flexAlignRightCenter
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   Dim SubTot As Double
   Dim Total As Double
   Dim i As Integer
   Dim MsgINTO As Boolean
   Dim MsgPasEx As Boolean
   Dim WhFecha As String
   
   '3023220 Tema 4
   capitalEfectivo = 0
   WhFecha = " AND Comprobante.Fecha BETWEEN " & CLng(Int(DateSerial(gEmpresa.Ano, 1, 1))) & " AND " & CLng(Int(DateSerial(gEmpresa.Ano, 12, 31)))

   Grid.Redraw = False
   
   Grid.rows = 0
   
   Grid.rows = Grid.rows + 1
   Row = 0
   Grid.TextMatrix(0, C_FMT) = "              .FMT"
   Grid.RowHeight(0) = 0  'Row con el formateo
   
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   
   'Total Activos
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Total Activos"
   
'   Q1 = "SELECT Sum(MovComprobante.Debe - MovComprobante.Haber) As TotalActivos "
   Q1 = "SELECT Sum(MovComprobante.Debe ) As TotalActivosDebe "
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Cuentas.Clasificacion =" & CLASCTA_ACTIVO
   '3042010 se agrega las siguientes condiciones
   Q1 = Q1 & " AND ((Cuentas.Atrib" & ATRIB_CAPITALPROPIO & " IS NULL OR Cuentas.Atrib" & ATRIB_CAPITALPROPIO & " = 0 )"
   Q1 = Q1 & " OR (Cuentas.Atrib" & ATRIB_CAPITALPROPIO & " <> 0 AND Cuentas.TipoCapPropio <> " & CAPPROPIO_ACTIVO_COMPACTIVO & "))"
   '3042010
   Q1 = Q1 & " AND TipoAjuste IN (" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.Estado = " & EC_APROBADO & WhFecha
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Grid.TextMatrix(Row, C_TOTAL) = Format(vFld(Rs("TotalActivosDebe")), NEGNUMFMT)
      Total = vFld(Rs("TotalActivosDebe"))
   End If
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT Sum(MovComprobante.Haber ) As TotalActivosHaber "            'haber sin Capital Propio - Complementario de Activo, a solicitud de Nicolás Catrin - 25 mar 2019
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Cuentas.Clasificacion =" & CLASCTA_ACTIVO
   Q1 = Q1 & " AND ((Cuentas.Atrib" & ATRIB_CAPITALPROPIO & " IS NULL OR Cuentas.Atrib" & ATRIB_CAPITALPROPIO & " = 0 )"
   Q1 = Q1 & " OR (Cuentas.Atrib" & ATRIB_CAPITALPROPIO & " <> 0 AND Cuentas.TipoCapPropio <> " & CAPPROPIO_ACTIVO_COMPACTIVO & "))"
   Q1 = Q1 & " AND TipoAjuste IN (" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.Estado = " & EC_APROBADO & WhFecha
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Grid.TextMatrix(Row, C_TOTAL) = Format(vFmt(Grid.TextMatrix(Row, C_TOTAL)) - vFld(Rs("TotalActivosHaber")), NEGNUMFMT)
      Total = Total - vFld(Rs("TotalActivosHaber"))
   End If
   
   Call CloseRs(Rs)
   
   'Menos valores que disminuyen los Activos
   
   Grid.rows = Grid.rows + 2
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Más valores que disminuyen los Activos"
   
   Q1 = "SELECT Cuentas.Descripcion as DescCta, Sum(MovComprobante.Debe - MovComprobante.Haber) As TotActComp "
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Cuentas.Clasificacion = " & CLASCTA_ACTIVO
   Q1 = Q1 & " AND Cuentas.Atrib" & ATRIB_CAPITALPROPIO & " <> 0 "
   Q1 = Q1 & " AND Cuentas.TipoCapPropio = " & CAPPROPIO_ACTIVO_COMPACTIVO
   Q1 = Q1 & " AND TipoAjuste IN (" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.Estado = " & EC_APROBADO & WhFecha
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY Cuentas.Descripcion"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   SubTot = 0
   Do While Rs.EOF = False
   
      Grid.rows = Grid.rows + 1
      Row = Row + 1
      
      Grid.TextMatrix(Row, C_DESC) = String(10, " ") & vFld(Rs("DescCta"), True)
      Grid.TextMatrix(Row, C_VALDET) = Format(vFld(Rs("TotActComp")), NEGNUMFMT)
      
      SubTot = SubTot + vFld(Rs("TotActComp"))
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
  
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.Row = Row
   Grid.Col = C_DESC
   Grid.CellAlignment = flexAlignRightCenter
   
   Grid.TextMatrix(Row, C_DESC) = "Total Deducciones"
   Grid.TextMatrix(Row, C_MENOS) = Format(SubTot, NEGNUMFMT)

   'Total = Total - SubTot
   Total = Total + SubTot     'ya viene negativo
   
   'Activo Depurado
   Grid.rows = Grid.rows + 2
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Activo Depurado"
   Grid.TextMatrix(Row, C_TOTAL) = Format(Total, NEGNUMFMT)
   
   'menos valores INTO
   
   Grid.rows = Grid.rows + 2
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Menos valores INTO"
   
   Q1 = "SELECT Cuentas.Descripcion as DescCta, Sum(MovComprobante.Debe - MovComprobante.Haber) As TotActComp "
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Cuentas.Clasificacion = " & CLASCTA_ACTIVO
   Q1 = Q1 & " AND Cuentas.Atrib" & ATRIB_CAPITALPROPIO & " <> 0 "
   Q1 = Q1 & " AND Cuentas.TipoCapPropio = " & CAPPROPIO_ACTIVO_VALINTO
   Q1 = Q1 & " AND TipoAjuste IN (" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.Estado = " & EC_APROBADO & WhFecha
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY Cuentas.Descripcion"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   SubTot = 0
   Do While Rs.EOF = False
   
      Grid.rows = Grid.rows + 1
      Row = Row + 1
      
      Grid.TextMatrix(Row, C_DESC) = String(10, " ") & vFld(Rs("DescCta"), True)
      Grid.TextMatrix(Row, C_VALDET) = Format(vFld(Rs("TotActComp")), NEGNUMFMT)
      
      SubTot = SubTot + vFld(Rs("TotActComp"))
      
      If vFld(Rs("TotActComp")) < 0 And MsgINTO = False Then '(esto no debiera ocurrir)
         MsgBox1 "ADVERTENCIA:" & vbNewLine & vbNewLine & "Una de las cuentas de valores INTO tiene saldo Acreedor. Para que esta cuenta no aparezca en este informe, elimine este atributo de la cuenta, utilizando el botón ""Plan de Cuentas"" y luego el botón Editar para la cuenta seleccionada.", vbInformation + vbOKOnly
         MsgINTO = True
      End If

      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
  
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.Row = Row
   Grid.Col = C_DESC
   Grid.CellAlignment = flexAlignRightCenter
   
   Grid.TextMatrix(Row, C_DESC) = "Total Valores INTO"
   Grid.TextMatrix(Row, C_MENOS) = Format(SubTot, NEGNUMFMT)

   If SubTot > 0 Then    'saldo Deudor (es lo que debería ser siempre)
      Total = Total - SubTot
   ElseIf SubTot < 0 Then
      Total = Total + SubTot        'ya vienen negativos (esto no debiera ocurrir)
      'MsgBox1 "ADVERTENCIA:" & vbNewLine & vbNewLine & "Los valores INTO tienen saldo Acreedor. Para que una o más cuentas no aparezcan en este informe, elimine este atributo de la cuenta, utilizando el botón ""Plan de Cuentas"" y luego el botón Editar para la cuenta seleccionada.", vbInformation + vbOKOnly
   End If
   
   'Capital Efectivo
   Grid.rows = Grid.rows + 2
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Capital Efectivo (cód. 102)"
   Grid.TextMatrix(Row, C_TOTAL) = Format(Total, NEGNUMFMT)
   '3023220 Tema 4
   capitalEfectivo = Total
   'menos Pasivo Exigible
   
   Grid.rows = Grid.rows + 2
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "Menos Pasivo Exigible"
   
   Q1 = "SELECT Cuentas.Descripcion as DescCta, Sum(MovComprobante.Debe - MovComprobante.Haber) As TotPasExig "
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Cuentas.Clasificacion = " & CLASCTA_PASIVO
   Q1 = Q1 & " AND Cuentas.Atrib" & ATRIB_CAPITALPROPIO & " <> 0 "
   Q1 = Q1 & " AND Cuentas.TipoCapPropio = " & CAPPROPIO_PASIVO_EXIGIBLE
   Q1 = Q1 & " AND TipoAjuste IN (" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.Estado = " & EC_APROBADO & WhFecha
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY Cuentas.Descripcion"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   SubTot = 0
   Do While Rs.EOF = False
   
      Grid.rows = Grid.rows + 1
      Row = Row + 1
      
      Grid.TextMatrix(Row, C_DESC) = String(10, " ") & vFld(Rs("DescCta"), True)
      Grid.TextMatrix(Row, C_VALDET) = Format(vFld(Rs("TotPasExig")), NEGNUMFMT)
      
      SubTot = SubTot + vFld(Rs("TotPasExig"))
      
      If vFld(Rs("TotPasExig")) > 0 And MsgPasEx = False Then '(esto no debiera ocurrir)
         MsgBox1 "ADVERTENCIA:" & vbNewLine & vbNewLine & "Una de las cuentas de Pasivo Exigible tiene saldo Deudor. Para que esta cuenta no aparezca en este informe, elimine este atributo de la cuenta, utilizando el botón ""Plan de Cuentas"" y luego el botón Editar para la cuenta seleccionada.", vbInformation + vbOKOnly
         MsgPasEx = True
      End If

      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
  
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.Row = Row
   Grid.Col = C_DESC
   Grid.CellAlignment = flexAlignRightCenter
   
   Grid.TextMatrix(Row, C_DESC) = "Total Pasivo Exigible"
   Grid.TextMatrix(Row, C_MENOS) = Format(SubTot, NEGNUMFMT)

   If SubTot > 0 Then      'saldo Deudor (esto no debiera ocurrir nunca)
      'MsgBox1 "ADVERTENCIA:" & vbNewLine & vbNewLine & "El Pasivo Exigible tiene saldo Deudor. Para que una o más cuentas no aparezcan en este informe, elimine este atributo de la cuenta, utilizando el botón ""Plan de Cuentas"" y luego presione el botón Editar para la cuenta seleccionada.", vbInformation + vbOKOnly
      Total = Total - SubTot
   ElseIf SubTot < 0 Then       'saldo Acreedor (es lo que debería ser siempre)
      Total = Total + SubTot        'ya vienen negativos
   End If
   
   'Capital Propio Tributario
   Grid.rows = Grid.rows + 2
   Row = Row + 2
   Call FGrSetRowStyle(Grid, Row, "B")
   Grid.TextMatrix(Row, C_FMT) = "B"
   Grid.TextMatrix(Row, C_DESC) = "CAPITAL PROPIO TRIBUTARIO "
   If Total >= 0 Then
      Grid.TextMatrix(Row, C_DESC) = Grid.TextMatrix(Row, C_DESC) & " (cód. 645)"
   Else
      Grid.TextMatrix(Row, C_DESC) = Grid.TextMatrix(Row, C_DESC) & " (cód. 646)"
   End If
   
   Grid.TextMatrix(Row, C_TOTAL) = Format(Total, NEGNUMFMT)

   Tx_TotCapPropio = Format(Total, NEGNUMFMT)
   
   Tx_CapPropio = "Capital Propio Tributario al 31 de Diciembre " & gEmpresa.Ano & " (Recuadro 3 Form. 22, Código "
   If Total >= 0 Then
      Tx_CapPropio = Tx_CapPropio & "645) "
   Else
      Tx_CapPropio = Tx_CapPropio & "646) "
   End If
      
   For i = Grid.FixedRows To Grid.rows - 1
      Grid.TextMatrix(i, C_OBLIGATORIA) = "O"
   Next i

   If Grid.rows <= 1 Then
      Grid.rows = 2
   End If
   Call FGrVRows(Grid)
   Grid.rows = Grid.rows + 2
         
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.RowSel = Grid.Row
   Grid.Col = 0
   Grid.ColSel = 0
   
   Grid.Redraw = True

End Sub
Public Sub FView()

   Me.Show vbModeless

End Sub

Public Sub GetcapitalEfectivo()

   Call SetUpGrid
   
   Call LoadAll

End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - Tx_CapPropio.Height - 500
   'Grid.Width = Me.Width - 230
   Tx_CapPropio.Top = Grid.Top + Grid.Height + 50
   Tx_TotCapPropio.Top = Tx_CapPropio.Top
   
   Call FGrLocateCntrl(Grid, Tx_TotCapPropio, C_TOTAL)
   Tx_CapPropio.Width = Tx_TotCapPropio.Left - Tx_CapPropio.Left - 50
  
   Call FGrVRows(Grid)

End Sub

Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(0) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = "Capital Propio Tributario"
   gPrtReportes.Titulos = Titulos
   Encabezados(0) = "Al 31 de Diciembre " & gEmpresa.Ano
   gPrtReportes.Encabezados = Encabezados
         
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
   'Total(C_DESC) = "Capital Pripio Tributario"
   'Total(C_TOTAL) = ""
                  
   gPrtReportes.ColWi = ColWi
   'gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_OBLIGATORIA
   gPrtReportes.NTotLines = 0

End Sub
Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   
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
Private Sub Bt_Print_Click()
   Dim OldOrientation As Integer
   
   OldOrientation = Printer.Orientation
   
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
   
   Printer.Orientation = OldOrientation
   
End Sub

