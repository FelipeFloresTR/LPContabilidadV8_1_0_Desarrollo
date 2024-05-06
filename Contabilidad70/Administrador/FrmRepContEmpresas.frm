VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmRepContEmpresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Control Empresas"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Buttons 
      Height          =   645
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   9735
      Begin VB.TextBox Tx_Ano 
         Height          =   315
         Left            =   3960
         TabIndex        =   8
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton Bt_Listar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         Picture         =   "FrmRepContEmpresas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Buscar número de cartola"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   8400
         TabIndex        =   6
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   7200
         TabIndex        =   5
         Top             =   180
         Width           =   1095
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
         Left            =   990
         Picture         =   "FrmRepContEmpresas.frx":043E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Copiar Excel"
         Top             =   180
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
         Left            =   90
         Picture         =   "FrmRepContEmpresas.frx":0883
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
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
         Left            =   495
         Picture         =   "FrmRepContEmpresas.frx":0D2A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Año:"
         Height          =   195
         Left            =   3480
         TabIndex        =   9
         Top             =   240
         Width           =   435
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6135
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10821
      _Version        =   393216
      Rows            =   30
      Cols            =   24
      FixedRows       =   2
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "FrmRepContEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDEMP = 0
Const C_UPDATE = 1
Const C_RUT = 2
Const C_RSOCIAL = 3
Const C_M1 = 4
Const C_M12 = C_M1 + 12 - 1
Const C_CALC_PROPIVA = C_M12 + 1
Const C_AF_DEP = C_M12 + 2
Const C_AF_CM = C_M12 + 3
Const C_AF_33BIS = C_M12 + 4
Const C_CM_ACTIVOS = C_M12 + 5
Const C_CM_PASIVOS = C_M12 + 6
Const C_AJUSTES_IFRS = C_M12 + 7
Const C_BALDEF = C_M12 + 8
Const C_CPT_MUN = C_M12 + 9
Const C_F22RENTA = C_M12 + 10

Const NCOLS = C_F22RENTA

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Bt_Listar_Click()
   Call LoadAll
End Sub

Private Sub bt_OK_Click()
   Call SaveAll
   
   Unload Me
   
End Sub

Private Sub Form_Load()

   Call SetUpGrid
   
   Tx_Ano = Year(Now)
   
   Call LoadAll
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_IDEMP) = 0
   Grid.ColWidth(C_UPDATE) = 0
   Grid.ColWidth(C_RUT) = 1100
   Grid.ColWidth(C_RSOCIAL) = 2750
   
   For i = 1 To 12
      Grid.ColWidth(C_M1 + i - 1) = 300
      Grid.TextMatrix(1, C_M1 + i - 1) = Left(gNomMes(i), 2)
   Next i
   
   Grid.ColWidth(C_AF_DEP) = 1030
   For i = C_AF_CM + 1 To Grid.Cols - 1
      Grid.ColWidth(i) = 900
   Next i
   
   Grid.TextMatrix(1, C_RUT) = "Rut"
   Grid.TextMatrix(1, C_RSOCIAL) = "Razón Social"

   Grid.TextMatrix(0, C_CALC_PROPIVA) = "Cálculo"
   Grid.TextMatrix(1, C_CALC_PROPIVA) = "Prop. IVA"
   Grid.TextMatrix(0, C_AF_DEP) = "Act. Fijo"
   Grid.TextMatrix(1, C_AF_DEP) = "Depreciación"
   Grid.TextMatrix(0, C_AF_CM) = "Act. Fijo"
   Grid.TextMatrix(1, C_AF_CM) = "C.M."
   Grid.TextMatrix(0, C_AF_33BIS) = "Act. Fijo"
   Grid.TextMatrix(1, C_AF_33BIS) = "33 Bis LIR"
   Grid.TextMatrix(0, C_CM_ACTIVOS) = "C.M."
   Grid.TextMatrix(1, C_CM_ACTIVOS) = "Activos"
   Grid.TextMatrix(0, C_CM_PASIVOS) = "C.M."
   Grid.TextMatrix(1, C_CM_PASIVOS) = "Pasivos"
   Grid.TextMatrix(0, C_AJUSTES_IFRS) = "Ajustes"
   Grid.TextMatrix(1, C_AJUSTES_IFRS) = "IFRS"
   Grid.TextMatrix(0, C_BALDEF) = "Balance"
   Grid.TextMatrix(1, C_BALDEF) = "Definitivo"
   Grid.TextMatrix(0, C_CPT_MUN) = "CPT"
   Grid.TextMatrix(1, C_CPT_MUN) = "Municip."
   Grid.TextMatrix(0, C_F22RENTA) = "F22"
   Grid.TextMatrix(1, C_F22RENTA) = "Renta"
   
   
   For i = 0 To Grid.Cols - 1
      If i >= C_M1 Then
         Grid.ColAlignment(i) = flexAlignCenterCenter
      Else
         Grid.ColAlignment(i) = flexAlignLeftCenter
      End If
   Next i
   
   
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - 500
   Grid.Width = Me.Width - 230
   
   Fr_Buttons.Width = Me.Width - Fr_Buttons.Left - 180
   Bt_Cancel.Left = Me.Width - 380 - Bt_Cancel.Width
   Bt_OK.Left = Bt_Cancel.Left - 150 - Bt_OK.Width
   
   Call FGrVRows(Grid)

End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer, n As Integer
   Dim j As Integer
   
   If vFmt(Tx_Ano) <= 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   
   Call InsertRegEmp(vFmt(Tx_Ano))
   
   Q1 = "SELECT ControlEmpresa.IdEmpresa, ControlEmpresa.RazonSocial, ControlEmpresa.RUT, "
   
   For i = 1 To 12
      Q1 = Q1 & " Mes" & i & ","
   Next i
   
   Q1 = Q1 & " AF_Depreciacion, AF_CM, AF_33BisLir, CM_Activos, CM_Pasivos, BalDefinitivo, CPT_Municip, F22Renta, AjustesIFRS, CalcPropIVA "
   Q1 = Q1 & " FROM ControlEmpresa INNER JOIN Empresas ON ControlEmpresa.IdEmpresa = Empresas.IdEmpresa "
   Q1 = Q1 & " WHERE Ano=" & vFmt(Tx_Ano)
   Q1 = Q1 & " AND Empresas.Estado = 0"
   If gAppCode.Demo Then
      Q1 = Q1 & " AND Empresas.RUT IN ('1','2','3')"
   End If

   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.Rows = Grid.FixedRows
   i = Grid.Rows
   n = 0
   
   Do While Rs.EOF = False
  
      Grid.Rows = Grid.Rows + 1
      n = n + 1
      
      Grid.TextMatrix(i, C_IDEMP) = vFld(Rs("IdEmpresa"))
      Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("RUT")))
      Grid.TextMatrix(i, C_RSOCIAL) = vFld(Rs("RazonSocial"), True)
      
      For j = 1 To 12
         Grid.TextMatrix(i, C_M1 + j - 1) = IIf(vFld(Rs("Mes" & j)) <> 0, "x", "")
      Next j
      
      Grid.TextMatrix(i, C_CALC_PROPIVA) = IIf(vFld(Rs("CalcPropIVA")) <> 0, "Si", "")
      Grid.TextMatrix(i, C_AF_DEP) = IIf(vFld(Rs("AF_Depreciacion")) <> 0, "Si", "")
      Grid.TextMatrix(i, C_AF_CM) = IIf(vFld(Rs("AF_CM")) <> 0, "Si", "")
      Grid.TextMatrix(i, C_AF_33BIS) = IIf(vFld(Rs("AF_33BisLir")) <> 0, "Si", "")
      Grid.TextMatrix(i, C_CM_ACTIVOS) = IIf(vFld(Rs("CM_Activos")) <> 0, "Si", "")
      Grid.TextMatrix(i, C_CM_PASIVOS) = IIf(vFld(Rs("CM_Pasivos")) <> 0, "Si", "")
      
      Grid.TextMatrix(i, C_AJUSTES_IFRS) = IIf(vFld(Rs("AjustesIFRS")) <> 0, "Si", "")
      Grid.TextMatrix(i, C_BALDEF) = IIf(vFld(Rs("BalDefinitivo")) <> 0, "OK", "")
      Grid.TextMatrix(i, C_CPT_MUN) = IIf(vFld(Rs("CPT_Municip")) <> 0, "OK", "")
      Grid.TextMatrix(i, C_F22RENTA) = IIf(vFld(Rs("F22Renta")) <> 0, "Presentado", "")
      
      If gAppCode.NivProd = VER_5EMP And n >= 5 Then
         Exit Do
      End If
      
      i = i + 1
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Me.MousePointer = vbDefault

   Call FGrVRows(Grid)
   Grid.Rows = Grid.Rows + 1
      
End Sub
Private Sub InsertRegEmp(Ano As Integer)
   Dim Rs As Recordset
   Dim Q1 As String
#If DATACON = 1 Then
   Dim Db As Database
#End If
   Dim DbPath As String
   Dim Rs2 As Recordset
   
   Me.MousePointer = vbHourglass
   
   Q1 = "SELECT EmpresasAno.IdEmpresa, EmpresasAno.Ano, Empresas.Rut, ControlEmpresa.RazonSocial "
   Q1 = Q1 & " FROM (EmpresasAno "
   Q1 = Q1 & " INNER JOIN Empresas ON EmpresasAno.idEmpresa = Empresas.idEmpresa) "
   Q1 = Q1 & " LEFT JOIN ControlEmpresa ON (EmpresasAno.idEmpresa = ControlEmpresa.IdEmpresa) AND (EmpresasAno.Ano = ControlEmpresa.Ano)"
   Q1 = Q1 & " WHERE EmpresasAno.Ano = " & Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
      
      If vFld(Rs("RazonSocial")) = "" Then 'no está el registro en la tabla ControlEmpresa, lo agregamos
         
#If DATACON = 1 Then

         DbPath = gDbPath & "\Empresas\" & Ano
         DbPath = DbPath & "\" & vFld(Rs("Rut")) & ".mdb"
         
         If ExistFile(DbPath) Then
         
            Q1 = "PWD=" & PASSW_PREFIX & vFld(Rs("Rut")) & ";"
            Call LinkMdbTable(DbMain, DbPath, "Empresa", "LnkEmpresa", , , Q1)
            
            Q1 = "SELECT RazonSocial FROM LnkEmpresa WHERE Id = " & vFld(Rs("IdEmpresa"))
            Set Rs2 = OpenRs(DbMain, Q1)

            If Rs2.EOF = False Then
               Q1 = "INSERT INTO ControlEmpresa "
               Q1 = Q1 & " (IdEmpresa, Ano, RazonSocial, RUT)"
               Q1 = Q1 & " VALUES(" & vFld(Rs("IdEmpresa")) & "," & Ano & ",'" & vFld(Rs2("RazonSocial")) & "'," & vFld(Rs("Rut")) & ")"
               Call ExecSQL(DbMain, Q1)
            End If
            
            Call CloseRs(Rs2)
            Call UnLinkTable(DbMain, "LnkEmpresa")
            
         End If
#Else
         Q1 = "SELECT RazonSocial FROM Empresa WHERE Id = " & vFld(Rs("IdEmpresa")) & " AND Ano = " & Ano
         Set Rs2 = OpenRs(DbMain, Q1)
         If Rs2.EOF = False Then
            Q1 = "INSERT INTO ControlEmpresa "
            Q1 = Q1 & " (IdEmpresa, Ano, RazonSocial, RUT)"
            Q1 = Q1 & " VALUES(" & vFld(Rs("IdEmpresa")) & "," & Ano & ",'" & vFld(Rs2("RazonSocial")) & "'," & vFld(Rs("Rut")) & ")"
            Call ExecSQL(DbMain, Q1)
         End If
         
         Call CloseRs(Rs2)
#End If
         
      End If
      
      Rs.MoveNext
            
   Loop
   
   Call CloseRs(Rs)
      
   Me.MousePointer = vbHourglass

End Sub

Private Sub Grid_DblClick()
   Dim Row As Integer
   Dim Col As Integer
   
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Col >= C_M1 And Col <= C_M12 Then
      
      If Trim(Grid.TextMatrix(Row, Col)) = "" Then
         Grid.TextMatrix(Row, Col) = "x"
      Else
         Grid.TextMatrix(Row, Col) = ""
      End If
      
      Call FGrModRow(Grid, Row, FGR_U, C_IDEMP, C_UPDATE)
   
   Else
   
      Select Case Col

         Case C_AF_DEP, C_AF_CM, C_AF_33BIS, C_CM_ACTIVOS, C_CM_PASIVOS, C_CALC_PROPIVA, C_AJUSTES_IFRS
   
            If Trim(Grid.TextMatrix(Row, Col)) = "" Then
               Grid.TextMatrix(Row, Col) = "Si"
            Else
               Grid.TextMatrix(Row, Col) = ""
            End If
            
            Call FGrModRow(Grid, Row, FGR_U, C_IDEMP, C_UPDATE)
      
         Case C_BALDEF, C_CPT_MUN
         
            If Trim(Grid.TextMatrix(Row, Col)) = "" Then
               Grid.TextMatrix(Row, Col) = "Ok"
            Else
               Grid.TextMatrix(Row, Col) = ""
            End If
            
            Call FGrModRow(Grid, Row, FGR_U, C_IDEMP, C_UPDATE)
            
         Case C_F22RENTA
         
            If Trim(Grid.TextMatrix(Row, Col)) = "" Then
               Grid.TextMatrix(Row, Col) = "Presentado"
            Else
               Grid.TextMatrix(Row, Col) = ""
            End If
            
            Call FGrModRow(Grid, Row, FGR_U, C_IDEMP, C_UPDATE)
            
      End Select
      
   End If
   
   If Grid.TextMatrix(Row, C_UPDATE) <> "" Then
      Call SetRO(Tx_Ano, True)
      Bt_Listar.Enabled = False
   End If

End Sub

Private Sub SaveAll()
   Dim i As Integer
   Dim Q1 As String
   Dim j As Integer
      
   For i = Grid.FixedRows To Grid.Rows - 1
   
      If Grid.TextMatrix(i, C_IDEMP) = "" Then
         Exit Sub
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) <> "" Then
      
         Q1 = "UPDATE ControlEmpresa SET "
            
         For j = 1 To 12
         
            Q1 = Q1 & "Mes" & j & "=" & IIf(Grid.TextMatrix(i, C_M1 + j - 1) <> "", 1, 0) & ","
            
         Next j
            
         Q1 = Q1 & "  AF_Depreciacion = " & IIf(Grid.TextMatrix(i, C_AF_DEP) <> "", 1, 0)
         Q1 = Q1 & ", AF_CM = " & IIf(Grid.TextMatrix(i, C_AF_CM) <> "", 1, 0)
         Q1 = Q1 & ", AF_33BisLir = " & IIf(Grid.TextMatrix(i, C_AF_33BIS) <> "", 1, 0)
         Q1 = Q1 & ", CM_Activos = " & IIf(Grid.TextMatrix(i, C_CM_ACTIVOS) <> "", 1, 0)
         Q1 = Q1 & ", CM_Pasivos = " & IIf(Grid.TextMatrix(i, C_CM_PASIVOS) <> "", 1, 0)
         Q1 = Q1 & ", BalDefinitivo = " & IIf(Grid.TextMatrix(i, C_BALDEF) <> "", 1, 0)
         Q1 = Q1 & ", CPT_Municip = " & IIf(Grid.TextMatrix(i, C_CPT_MUN) <> "", 1, 0)
         Q1 = Q1 & ", F22Renta = " & IIf(Grid.TextMatrix(i, C_F22RENTA) <> "", 1, 0)
         Q1 = Q1 & ", CalcPropIVA = " & IIf(Grid.TextMatrix(i, C_CALC_PROPIVA) <> "", 1, 0)
         Q1 = Q1 & ", AjustesIFRS = " & IIf(Grid.TextMatrix(i, C_AJUSTES_IFRS) <> "", 1, 0)
                  
         Q1 = Q1 & " WHERE IdEmpresa= " & Val(Grid.TextMatrix(i, C_IDEMP)) & " AND Ano=" & vFmt(Tx_Ano)
         
         Call ExecSQL(DbMain, Q1)

      End If

   Next i
      
End Sub
Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim PrtOrient As Integer
   
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   
   PrtOrient = Printer.Orientation
   Printer.Orientation = cdlLandscape
   
   Call gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   Set Frm = Nothing
   
   Printer.Orientation = PrtOrient
   
   Me.MousePointer = vbDefault

End Sub

Private Sub Bt_Print_Click()
   Dim i As Integer
   Dim PrtOrient As Integer
      
   If Grid.TextMatrix(Grid.FixedRows, C_RUT) = "" Then
      Exit Sub
   End If
   
   Call SetUpPrtGrid
   
   PrtOrient = Printer.Orientation
   Printer.Orientation = cdlLandscape
   
   MousePointer = vbHourglass
   Call gPrtReportes.PrtFlexGrid(Printer)
   Me.MousePointer = vbDefault
      
   Printer.Orientation = PrtOrient
   MousePointer = vbDefault
   
End Sub
Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Caption & " Año:" & Tx_Ano)
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(1) As String
   Dim Encabezados(0) As String
   
   Me.MousePointer = vbHourglass
   
   Set gPrtReportes.Grid = Grid
      
   Titulos(0) = Me.Caption
   Titulos(1) = " Año " & Tx_Ano
   gPrtReportes.Titulos = Titulos
   
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize

   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   ColWi(C_RSOCIAL) = ColWi(C_RSOCIAL) - 100
   ColWi(C_AF_DEP) = ColWi(C_AF_DEP) + 100
   ColWi(C_AF_CM) = ColWi(C_AF_CM) - 200
   ColWi(C_CM_ACTIVOS) = ColWi(C_CM_ACTIVOS) - 200
   ColWi(C_CM_PASIVOS) = ColWi(C_CM_PASIVOS) - 200
   ColWi(C_CPT_MUN) = ColWi(C_CPT_MUN) - 200
      
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_RUT
   gPrtReportes.NTotLines = 0
   
End Sub

