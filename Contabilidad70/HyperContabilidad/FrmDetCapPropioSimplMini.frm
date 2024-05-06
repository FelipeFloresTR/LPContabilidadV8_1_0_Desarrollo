VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmDetCapPropioSimplMini 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Otros Ajustes (Aumentos)"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_DetCapAcum 
      Caption         =   "Base Imponible Acumulada..."
      Height          =   375
      Left            =   4500
      TabIndex        =   1
      Top             =   6600
      Width           =   3735
   End
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   5235
      Left            =   180
      TabIndex        =   0
      Top             =   780
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9234
      Cols            =   2
      Rows            =   4
      FixedCols       =   1
      FixedRows       =   1
      ScrollBars      =   3
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton Bt_VerComp 
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
         Picture         =   "FrmDetCapPropioSimplMini.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Detalle comprobante seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Del 
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
         Left            =   600
         Picture         =   "FrmDetCapPropioSimplMini.frx":0423
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Eliminar registro seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   6360
         TabIndex        =   11
         Top             =   180
         Width           =   1035
      End
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
         Left            =   1500
         Picture         =   "FrmDetCapPropioSimplMini.frx":081F
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   7500
         TabIndex        =   12
         Top             =   180
         Width           =   1035
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
         Left            =   1080
         Picture         =   "FrmDetCapPropioSimplMini.frx":0CD9
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   1920
         Picture         =   "FrmDetCapPropioSimplMini.frx":1180
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   3840
         Picture         =   "FrmDetCapPropioSimplMini.frx":15C5
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   3000
         Picture         =   "FrmDetCapPropioSimplMini.frx":19EE
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   3420
         Picture         =   "FrmDetCapPropioSimplMini.frx":1D8C
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   2460
         Picture         =   "FrmDetCapPropioSimplMini.frx":20ED
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridTot 
      Height          =   315
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6060
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   19
      FixedCols       =   2
      ForeColor       =   0
      ForeColorFixed  =   16711680
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "FrmDetCapPropioSimplMini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDDETCAPPROPIOSIMPL = 0
Const C_ANO = 1
Const C_FECHA = 2
Const C_LNGFECHA = 3
Const C_DESCRIP = 4
Const C_MONTO = 5
Const C_INGRESOMANUAL = 6
Const C_LINEAENBLANCO = 7
Const C_COLOBLIGATORIA = 8
Const C_FMT = 9
Const C_UPDATE = 10

Const NCOLS = C_UPDATE

Dim lRc As Integer
Dim lValorTotal As Double
Dim lValorAnual As Double
Dim lTipoDetCapPropioSimpl As String
Dim lRowTotAno As Integer
Dim lTipoInforme As Integer

Public Function FEdit(ByVal TipoDetCapPropioSimpl As Integer, ByVal TipoInforme As Integer, Valor As Double) As Integer

   lTipoDetCapPropioSimpl = TipoDetCapPropioSimpl
   lTipoInforme = TipoInforme
   
   Me.Show vbModal
   
   If lRc = vbOK Then
      If lTipoInforme = CPS_TIPOINFO_GENERAL Then
         Valor = lValorTotal
      Else
         Valor = lValorAnual
      End If
   End If
   
   FEdit = lRc
   
End Function


Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   
   Unload Me
End Sub

Private Sub Bt_Del_Click()
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   
   Row = Grid.Row
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.Row <> Grid.RowSel Then
      MsgBox1 "Debe eliminar un registro a la vez.", vbExclamation
      Exit Sub
   End If
   
   If Grid.RowHeight(Row) = 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_FECHA) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Val(Grid.TextMatrix(Row, C_INGRESOMANUAL)) = 0 Then
      MsgBox1 "Sólo se pueden eliminar los registros de ingreso manual.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   
   If MsgBox1("¿Está seguro que desea eliminar este registro?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
      
   Call FGrModRow(Grid, Row, FGR_D, C_IDDETCAPPROPIOSIMPL, C_UPDATE)
      
   Grid.rows = Grid.rows + 1
      
   Call CalcTot
End Sub

Private Sub Bt_DetCapAcum_Click()
   Dim Frm As FrmDetCapPropioSimplAcum
   Dim Valor As Double
   Dim Rc As Integer
   
   If Valida() Then
      Call SaveAll

      Set Frm = New FrmDetCapPropioSimplAcum
      Valor = lValorAnual
      Rc = Frm.FEdit(lTipoDetCapPropioSimpl, Valor)
      Set Frm = Nothing
      
      Call LoadAll
      
   End If

End Sub

Private Sub Bt_Ejemplos_Click()
   Dim Frm As FrmAyuda
   
   Set Frm = New FrmAyuda
   
   If lTipoDetCapPropioSimpl = CPS_OTROSAJUSTAUMENTOS Then
      Frm.FViewOtrosAjustesAumentos (Me.Caption)
   Else
      Frm.FViewOtrosAjustesDismin (Me.Caption)
   End If
   
   Set Frm = Nothing
   
End Sub

Private Sub Bt_OK_Click()

   If Valida() Then
      Call SaveAll
   
      lRc = vbOK
      
      Unload Me
   End If
   
End Sub

Private Sub Form_Load()

   If lTipoInforme = CPS_TIPOINFO_VARANUAL Then
      Bt_DetCapAcum.visible = False
   End If
   
   Me.Caption = gTipoDetCapPropioSimpl(lTipoDetCapPropioSimpl)
   
   Bt_DetCapAcum.Caption = "Acum. Anual " & Left(gTipoDetCapPropioSimpl(lTipoDetCapPropioSimpl), 30) & "..."
      
   Call SetUpGrid
   
   Call LoadAll
   
End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Total As Double, TotAno As Double

   Grid.FlxGrid.Redraw = False
   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   If lTipoInforme = CPS_TIPOINFO_GENERAL Then
     'Cargamos los totales acumulados años anteriores
      Q1 = "SELECT AnoValor, Valor FROM CapPropioSimplAnual"
      Q1 = Q1 & " WHERE TipoDetCPS = " & lTipoDetCapPropioSimpl & " AND AnoValor < " & gEmpresa.Ano
      Q1 = Q1 & " And Valor <> 0 AND IdEmpresa = " & gEmpresa.id
      Q1 = Q1 & " ORDER BY AnoValor "
      
      Set Rs = OpenRs(DbMain, Q1)
      
      Do While Not Rs.EOF
         Grid.rows = Grid.rows + 1
         
         Grid.TextMatrix(i, C_IDDETCAPPROPIOSIMPL) = 0
         Grid.TextMatrix(i, C_INGRESOMANUAL) = 0
         Grid.TextMatrix(i, C_ANO) = "Año " & vFld(Rs("AnoValor"))
         Grid.TextMatrix(i, C_MONTO) = Format(Abs(vFld(Rs("Valor"))), NUMFMT)
         Grid.TextMatrix(i, C_COLOBLIGATORIA) = "."
         Grid.TextMatrix(i, C_FMT) = "B"
         Call FGrSetRowStyle(Grid, i, "B")
         
         Total = Total + Abs(vFld(Rs("Valor")))
   
         i = i + 1
         Rs.MoveNext
      Loop
      
      Call CloseRs(Rs)
      
      If i > Grid.FixedRows Then
         Grid.rows = Grid.rows + 1
         Grid.TextMatrix(i, C_LINEAENBLANCO) = -1  'para que no lo edite el usuario (BeforeEdit)
         Grid.TextMatrix(i, C_COLOBLIGATORIA) = "."
      Else
         i = i - 1
      End If
      
      Grid.rows = Grid.rows + 1
      i = i + 1
      
   Else
      Grid.rows = Grid.rows + 1
      
   End If
   
   Grid.TextMatrix(i, C_IDDETCAPPROPIOSIMPL) = 0
   Grid.TextMatrix(i, C_INGRESOMANUAL) = 0
   Grid.TextMatrix(i, C_ANO) = "Año " & gEmpresa.Ano
   Grid.TextMatrix(i, C_COLOBLIGATORIA) = "."
   Grid.TextMatrix(i, C_FMT) = "B"
   Call FGrSetRowStyle(Grid, i, "B")
   lRowTotAno = i
   
   Grid.rows = Grid.rows + 1
   i = i + 1

   'Cargamos el detalle del año desde los comprobantes
   
'   If lTipoDetCapPropioSimpl = CPS_RETDIV Then
'
'      Q1 = "SELECT 0 as IngresoManual, 0 as IdDetCapPropioSimpl, MovComprobante.IdComp, MovComprobante.IdMov, "
'      Q1 = Q1 & " MovComprobante.IdCuenta, Cuentas.Codigo, Cuentas.Descripcion, "
'      Q1 = Q1 & " Comprobante.Tipo as CompTipo, Comprobante.Correlativo as CompCorr, Comprobante.Fecha As Fecha,  "
'      Q1 = Q1 & " MovComprobante.Debe - MovComprobante.Haber as Valor "
'      Q1 = Q1 & " FROM (((Socios INNER JOIN MovComprobante ON Socios.IdCuentaRetiros = MovComprobante.IdCuenta"
'      Q1 = Q1 & JoinEmpAno(gDbType, "Socios", "MovComprobante") & " )"
'      Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
'      Q1 = Q1 & JoinEmpAno(gDbType, "MovComprobante", "Cuentas") & ")"
'      Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
'      Q1 = Q1 & JoinEmpAno(gDbType, "MovComprobante", "Comprobante") & ")"
'      Q1 = Q1 & " WHERE Comprobante.Fecha BETWEEN " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " AND " & CLng(DateSerial(gEmpresa.Ano, 12, 31))
'      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
'      Q1 = Q1 & " AND ( Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste = " & TAJUSTE_AMBOS & ")"
'      Q1 = Q1 & " AND Comprobante.Estado = " & EC_APROBADO
'      Q1 = Q1 & " AND Comprobante.Tipo = " & TC_EGRESO
'
'   Else
'
'      Q1 = "SELECT 0 as IngresoManual, 0 as IdDetCapPropioSimpl, MovComprobante.IdComp, MovComprobante.IdMov, "
'      Q1 = Q1 & " MovComprobante.IdCuenta, Cuentas.Codigo, Cuentas.Descripcion, "
'      Q1 = Q1 & " Comprobante.Tipo as CompTipo, Comprobante.Correlativo as CompCorr, Comprobante.Fecha As Fecha,  "
'      Q1 = Q1 & " MovComprobante.Debe - MovComprobante.Haber as Valor "
'      Q1 = Q1 & " FROM ((MovComprobante INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
'      Q1 = Q1 & JoinEmpAno(gDbType, "MovComprobante", "Cuentas") & ")"
'      Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
'      Q1 = Q1 & JoinEmpAno(gDbType, "MovComprobante", "Comprobante") & ")"
'      Q1 = Q1 & " WHERE Comprobante.Fecha BETWEEN " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " AND " & CLng(DateSerial(gEmpresa.Ano, 12, 31))
'      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
'      Q1 = Q1 & " AND ( Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste = " & TAJUSTE_AMBOS & ")"
'      Q1 = Q1 & " AND Comprobante.Estado = " & EC_APROBADO
'      Q1 = Q1 & " AND Cuentas.CodF22_14Ter = " & lCodF22_14Ter
'
'      If lTipoDetCapPropioSimpl = CPS_PARTICIPACIONES Then
'         Q1 = Q1 & " AND Comprobante.Tipo = " & TC_INGRESO
'
'      ElseIf lTipoDetCapPropioSimpl = CPS_DISMINUCIONES Then
'         Q1 = Q1 & " AND Comprobante.Tipo IN ( " & TC_EGRESO & "," & TC_TRASPASO & ")"
'
'      ElseIf lTipoDetCapPropioSimpl = CPS_GASTOSRECHAZADOS Then
'         Q1 = Q1 & " AND Comprobante.Tipo = " & TC_EGRESO
'
'      End If
'
'   End If
'
'   Q1 = Q1 & " UNION "
   
   'Más el detalle de ingreso manual
'   Q1 = " SELECT 1 as IngresoManual, IdDetCapPropioSimpl, 0 as IdComp, 0 as IdMov, 0 as IdCuenta, ' ' As Codigo, ' ' As Descripcion, "
'   Q1 = Q1 & " 0 as CompTipo, 0 as CompCorr, Fecha, Valor  "
   Q1 = " SELECT 1 as IngresoManual, IdDetCapPropioSimpl, "
   Q1 = Q1 & " Fecha, Descrip, Valor  "
   Q1 = Q1 & " FROM DetCapPropioSimpl "
   Q1 = Q1 & " WHERE IngresoManual = 1 AND TipoDetCPS = " & lTipoDetCapPropioSimpl
   Q1 = Q1 & " AND DetCapPropioSimpl.IdEmpresa = " & gEmpresa.id & " AND DetCapPropioSimpl.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Fecha, IdDetCapPropioSimpl"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   
   Do While Not Rs.EOF
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDDETCAPPROPIOSIMPL) = vFld(Rs("IdDetCapPropioSimpl"))
      Grid.TextMatrix(i, C_INGRESOMANUAL) = vFld(Rs("IngresoManual"))
      Grid.TextMatrix(i, C_FECHA) = IIf(vFld(Rs("Fecha")) > 0, Format(vFld(Rs("Fecha")), EDATEFMT), "")
      Grid.TextMatrix(i, C_LNGFECHA) = vFld(Rs("Fecha"))
      Grid.TextMatrix(i, C_DESCRIP) = vFld(Rs("Descrip"))
      Grid.TextMatrix(i, C_MONTO) = Format(Abs(vFld(Rs("Valor"))), NUMFMT)
      Grid.TextMatrix(i, C_COLOBLIGATORIA) = "."
      
      Total = Total + Abs(vFld(Rs("Valor")))
      TotAno = TotAno + Abs(vFld(Rs("Valor")))
      

      i = i + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid)
   Grid.Row = Grid.FixedRows
   Grid.Col = C_FECHA
   
   Grid.FlxGrid.Redraw = True
   
   Grid.TextMatrix(lRowTotAno, C_MONTO) = Format(TotAno, NUMFMT)
   GridTot.TextMatrix(0, C_MONTO) = Format(Total, NUMFMT)

End Sub

Private Sub SaveAll()
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   
'   'eliminamos todos los registros de detalle de ingreso automático (información de detalle del año en la contabilidad)  y los volvemos a agregar
'   Call DeleteSQL(DbMain, "DetCapPropioSimpl", " WHERE IngresoManual = 0 AND TipoDetCPS = " & lTipoDetCapPropioSimpl & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano)
'
'   'Y ahora los insertamos nuevamente
'   For i = Grid.FixedRows To Grid.Rows - 1
'
'      If Val(Grid.TextMatrix(i, C_INGRESOMANUAL)) = 0 And Grid.TextMatrix(i, C_CUENTA) <> "" Then
'
'         Q1 = "INSERT INTO DetCapPropioSimpl "
'         Q1 = Q1 & " (IdEmpresa, Ano, TipoDetCPS, IngresoManual, IdCuenta, CodCuenta, Fecha, IdMovComp, Valor) VALUES "
'         Q1 = Q1 & " (" & gEmpresa.id & ", " & gEmpresa.Ano
'         Q1 = Q1 & ", " & lTipoDetCapPropioSimpl
'         Q1 = Q1 & ", 0 "
'         Q1 = Q1 & ", " & vFmt(Grid.TextMatrix(i, C_IDCUENTA))
'         Q1 = Q1 & ", '" & VFmtCodigoCta(Grid.TextMatrix(i, C_CODCUENTA)) & "'"
'         Q1 = Q1 & ", " & Grid.TextMatrix(i, C_LNGFECHA)
'         Q1 = Q1 & ", " & VFmtDate(Grid.TextMatrix(i, C_IDMOVCOMP))
'         Q1 = Q1 & ", " & vFmt(Grid.TextMatrix(i, C_MONTO)) & ")"
'         Call ExecSQL(DbMain, Q1)
'
'      End If
'
'   Next i
   
   'Ahora los de ingreso manual
   
    For i = Grid.FixedRows To Grid.rows - 1
      
      If Val(Grid.TextMatrix(i, C_INGRESOMANUAL)) <> 0 Then
  
         If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then
         
            Q1 = "INSERT INTO DetCapPropioSimpl "
            Q1 = Q1 & " (IdEmpresa, Ano, TipoDetCPS, IngresoManual, IdCuenta, CodCuenta, Fecha, Descrip, IdMovComp, Valor) VALUES "
            Q1 = Q1 & " (" & gEmpresa.id & ", " & gEmpresa.Ano
            Q1 = Q1 & ", " & lTipoDetCapPropioSimpl
            Q1 = Q1 & ", 1 "
            Q1 = Q1 & ", 0 "
            Q1 = Q1 & ", ' '"
            Q1 = Q1 & ", " & Grid.TextMatrix(i, C_LNGFECHA)
            Q1 = Q1 & ", '" & ParaSQL(Grid.TextMatrix(i, C_DESCRIP)) & "'"
            Q1 = Q1 & ", 0 "
            Q1 = Q1 & ", " & vFmt(Grid.TextMatrix(i, C_MONTO)) & ")"
            Call ExecSQL(DbMain, Q1)
        
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then
        
            Q1 = "UPDATE  DetCapPropioSimpl "
            Q1 = Q1 & " SET Fecha = " & Grid.TextMatrix(i, C_LNGFECHA)
            Q1 = Q1 & ", Descrip = '" & ParaSQL(Grid.TextMatrix(i, C_DESCRIP)) & "'"
            Q1 = Q1 & ", Valor = " & vFmt(Grid.TextMatrix(i, C_MONTO))
            Q1 = Q1 & " WHERE IdDetCapPropioSimpl = " & Val(Grid.TextMatrix(i, C_IDDETCAPPROPIOSIMPL))
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Call ExecSQL(DbMain, Q1)
            
         ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_D Then
        
            Q1 = " WHERE IdDetCapPropioSimpl = " & Val(Grid.TextMatrix(i, C_IDDETCAPPROPIOSIMPL))
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
            Call DeleteSQL(DbMain, "DetCapPropioSimpl", Q1)
            
         End If
         
      End If
      
   Next i
         
   'Actualizamos los totales en la tabla EmpresaAno
   
   Q1 = "UPDATE EmpresasAno SET "
   
   Select Case lTipoDetCapPropioSimpl
      Case CPS_OTROSAJUSTAUMENTOS
         Q1 = Q1 & " CPS_OtrosAjustesAumentos = "
      Case CPS_OTROSAJUSTDISMIN
         Q1 = Q1 & " CPS_OtrosAjustesDisminuciones = "
   End Select
   
   Q1 = Q1 & vFmt(GridTot.TextMatrix(0, C_MONTO))        'total acumulado con años anteriores, que es el que va a RAB
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
   
   lValorAnual = vFmt(Grid.TextMatrix(lRowTotAno, C_MONTO))

   'Actualizamos el total anual en la tabla CapPropioSimplAnual
   Q1 = "SELECT IdCapPropioSimplAnual FROM CapPropioSimplAnual "
   Q1 = Q1 & " WHERE TipoDetCPS = " & lTipoDetCapPropioSimpl & " AND AnoValor = " & gEmpresa.Ano
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Q1 = "UPDATE CapPropioSimplAnual SET Valor = " & lValorAnual & ", IngresoManual = 0 "
      Q1 = Q1 & " WHERE TipoDetCPS = " & lTipoDetCapPropioSimpl & " AND AnoValor = " & gEmpresa.Ano
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
   Else
      Q1 = "INSERT INTO CapPropioSimplAnual (TipoDetCPS, IngresoManual, AnoValor, Valor, IdEmpresa )"
      Q1 = Q1 & " VALUES( " & lTipoDetCapPropioSimpl
      Q1 = Q1 & ", 0"
      Q1 = Q1 & ", " & gEmpresa.Ano
      Q1 = Q1 & ", " & lValorAnual
      Q1 = Q1 & ", " & gEmpresa.id & ") "
      
   End If
   
   Call ExecSQL(DbMain, Q1)
   
   Call CloseRs(Rs)
   
   lValorTotal = vFmt(GridTot.TextMatrix(0, C_MONTO))
   
End Sub

Private Function Valida() As Boolean
   Dim i As Integer

   Valida = False
   
   For i = Grid.FixedRows To Grid.rows - 1
         
      If Val(Grid.TextMatrix(i, C_INGRESOMANUAL)) <> 0 Then

         If Grid.TextMatrix(i, C_FECHA) = "" Then
            MsgBox1 "Falta ingresar la fecha.", vbExclamation
            Grid.RowSel = i
            Grid.ColSel = C_FECHA
            Exit Function
         End If
         
         If Grid.TextMatrix(i, C_MONTO) = "" Then
            MsgBox1 "Falta ingresar la valor.", vbExclamation
            Grid.Row = i
            Grid.Col = C_MONTO
            Exit Function
         End If
         
      End If
      
   Next i

   Valida = True

End Function
Private Sub SetUpGrid()
   Dim i As Integer, WCodCuenta As Integer, WCuenta As Integer
   
   Grid.Cols = NCOLS + 1
      
   Call FGrSetup(Grid, True)

   Grid.ColWidth(C_IDDETCAPPROPIOSIMPL) = 0
   Grid.ColWidth(C_INGRESOMANUAL) = 0
   Grid.ColWidth(C_ANO) = 1300
   Grid.ColWidth(C_FECHA) = 1200
   Grid.ColWidth(C_LNGFECHA) = 0
   Grid.ColWidth(C_DESCRIP) = 4000
   Grid.ColWidth(C_MONTO) = 1500
   Grid.ColWidth(C_LINEAENBLANCO) = 0
   Grid.ColWidth(C_COLOBLIGATORIA) = 0
   Grid.ColWidth(C_UPDATE) = 0
   Grid.ColWidth(C_FMT) = 0
   
   Grid.ColAlignment(C_FECHA) = flexAlignRightCenter
   Grid.ColAlignment(C_MONTO) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_ANO) = "Año"
   Grid.TextMatrix(0, C_FECHA) = "Fecha"
   Grid.TextMatrix(0, C_DESCRIP) = "Descripción"
   Grid.TextMatrix(0, C_MONTO) = "Monto"
   
   Call FGrVRows(Grid)
   Call FGrTotales(Grid, GridTot)
      
   GridTot.TextMatrix(0, C_ANO) = "Total"

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
   
   Call ResetPrtBas(gPrtReportes)
  
End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Clip As String
   
'   Call LP_FGr2Clip(Grid, Me.Caption)
   Clip = LP_FGr2String(Grid, Me.Caption & vbTab & "Año " & gEmpresa.Ano, False, C_FECHA)
   
   If Clip <> "" Then
      Clip = Clip & FGr2String(GridTot)
      
      Clipboard.Clear
      Clipboard.SetText Clip
   End If
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

Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(1) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Me.Caption
   Titulos(1) = "Año " & gEmpresa.Ano
   gPrtReportes.Titulos = Titulos
'   gPrtReportes.Encabezados = Encabezados
         
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
   Total(C_ANO) = GridTot.TextMatrix(0, C_ANO)
   Total(C_MONTO) = GridTot.TextMatrix(0, C_MONTO)
                  
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_ANO
   gPrtReportes.NTotLines = 1

End Sub



Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim Fecha As Long
   
   Action = vbOK

   If Col = C_FECHA Then
      If Value <> "" Then
         Fecha = GetDate(Value, "dmy")
         
         If Year(Fecha) <> gEmpresa.Ano Then
            If MsgBox1("Advertencia: esta fecha no pertenece al año actual." & vbCrLf & "Desea continuar?", vbExclamation + vbYesNo) = vbNo Then
               Action = vbRetry
               Exit Sub
            End If
         End If
         
         Value = Format(Fecha, EDATEFMT)
         Grid.TextMatrix(Row, Col) = Value
         Grid.TextMatrix(Row, C_LNGFECHA) = Fecha
         
         Grid.rows = Grid.rows + 1
      End If
   
   ElseIf Col = C_DESCRIP Then
      Value = Trim(Value)

   ElseIf Col = C_MONTO Then
      Value = Format(vFmt(Value), NUMFMT)
      Grid.TextMatrix(Row, Col) = Value
      Call CalcTot
   End If

   If Action = vbOK Then
      Grid.TextMatrix(Row, C_INGRESOMANUAL) = 1
      Call FGrModRow(Grid, Row, FGR_U, C_IDDETCAPPROPIOSIMPL, C_UPDATE)
   End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   
   If Grid.TextMatrix(Row, C_LINEAENBLANCO) <> "" Or Grid.TextMatrix(Row, C_ANO) <> "" Then  'fila acum años o en blanco
      Exit Sub
   End If
      
   If Grid.TextMatrix(Row - 1, C_ANO) = "" And (Grid.TextMatrix(Row - 1, C_FECHA) = "" Or Grid.TextMatrix(Row - 1, C_MONTO) = "") Then
      MsgBox1 "Debe completar la fila anterior antes de continuar.", vbExclamation
      Exit Sub
   End If

   If Col = C_FECHA Then
   
      EdType = FEG_Edit
      Grid.TxBox.MaxLength = 10
  
   ElseIf Col = C_DESCRIP Then
   
      If Grid.TextMatrix(Row, C_FECHA) = "" Then
         MsgBox1 "Debe ingresar primero la fecha.", vbExclamation
         Exit Sub
      End If

      EdType = FEG_Edit
      Grid.TxBox.MaxLength = 80
      
   ElseIf Col = C_MONTO Then
   
      If Grid.TextMatrix(Row, C_FECHA) = "" Then
         MsgBox1 "Debe ingresar primero la fecha.", vbExclamation
         Exit Sub
      End If

      EdType = FEG_Edit
      Grid.TxBox.MaxLength = 12
   End If
            
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   
   If Grid.Col = C_MONTO Then
      Call KeyNumPos(KeyAscii)
   ElseIf Grid.Col = C_FECHA Then
      Call KeyDate(KeyAscii)
   ElseIf Grid.Col = C_DESCRIP Then
      Call KeyName(KeyAscii)
   End If
   
End Sub

Private Sub Grid_Scroll()
   GridTot.LeftCol = Grid.LeftCol

End Sub

Private Sub CalcTot()
   Dim Total As Double
   Dim i As Integer
   Dim TotAno As Double

   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.RowHeight(i) > 0 Then
      
         If Grid.TextMatrix(i, C_FECHA) <> "" Then
            TotAno = TotAno + vFmt(Grid.TextMatrix(i, C_MONTO))
         End If
            
         If i <> lRowTotAno Then
            Total = Total + vFmt(Grid.TextMatrix(i, C_MONTO))
         End If
      End If
      
   Next i
   
   Grid.TextMatrix(lRowTotAno, C_MONTO) = Format(TotAno, NUMFMT)
   GridTot.TextMatrix(0, C_MONTO) = Format(Total, NUMFMT)
   
End Sub


