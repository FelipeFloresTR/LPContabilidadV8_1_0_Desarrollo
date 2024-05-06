VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmAsistImpPrimCat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asistente de cálculo Impuesto de Primera Categoría"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   13230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13155
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
         Left            =   2040
         Picture         =   "FrmAsistImpPrimCat.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Sumar movimientos seleccionados"
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
         Left            =   3000
         Picture         =   "FrmAsistImpPrimCat.frx":00A4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Calculadora"
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
         Left            =   2580
         Picture         =   "FrmAsistImpPrimCat.frx":0405
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Convertir moneda"
         Top             =   180
         Width           =   375
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
         Picture         =   "FrmAsistImpPrimCat.frx":07A3
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Detalle comprobante seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   10440
         TabIndex        =   6
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
         Left            =   3420
         Picture         =   "FrmAsistImpPrimCat.frx":0C08
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Calendario"
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
         Left            =   1500
         Picture         =   "FrmAsistImpPrimCat.frx":1031
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
         Left            =   660
         Picture         =   "FrmAsistImpPrimCat.frx":1476
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   11760
         TabIndex        =   2
         Top             =   180
         Width           =   1275
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
         Left            =   1080
         Picture         =   "FrmAsistImpPrimCat.frx":191D
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   3435
      Left            =   60
      TabIndex        =   11
      Top             =   840
      Width           =   13130
      _ExtentX        =   23151
      _ExtentY        =   6059
      Cols            =   2
      Rows            =   2
      FixedCols       =   1
      FixedRows       =   0
      ScrollBars      =   0
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
End
Attribute VB_Name = "FrmAsistImpPrimCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_ID = 0
Const C_IDITEM = 1
Const C_CONCEPTO = 2
Const C_REMEJANTNOMINAL = 3
Const C_REMEJANTACT = 4
Const C_GENERADOANO = 5
Const C_CREDUTILIZADO = 6
Const C_REMEJSGTE = 7
Const C_FMT = 8
Const C_COLOBLIGATORIA = 9
Const C_UPD = 10

Const NCOLS = C_UPD

Dim lIDPCBaseImp As Double
Dim lGeneradoAno As Double
Dim lCred33bisUtilizado As Double
Dim lMayorValorEnajenación As Double

Dim lRowCredIng As Integer
Dim lRowCredRetiros As Integer
Dim lRowIDPCNetoPagar As Integer
Dim lRowIDPCPagar As Integer
Dim lRowMayorValorEnajenacion As Integer

Private Sub Bt_DetDoc_Click()
   Dim Row As Integer, Col As Integer
   Dim IdItem As Integer
   Dim Frm As Form
   
   Row = Grid.Row
   Col = Grid.Col
   
   IdItem = Val(Grid.TextMatrix(Row, C_IDITEM))
   
   Select Case IdItem
      Case 1, 7
         If Col = C_CREDUTILIZADO Then  'IDPC sobre Base imponible
            Set Frm = New FrmBaseImponible
            Frm.Show vbModal
            Set Frm = Nothing
         End If
         
      Case 3
         If Col = C_GENERADOANO Then  'Crédito 33 bis
            Set Frm = New FrmAjustesExtraLibCaja
            Frm.Show vbModal
            Set Frm = Nothing
         End If
   End Select
   
End Sub

Private Sub bt_OK_Click()

   If Valida() Then
      SaveAll
      Unload Me
   End If
   
End Sub
Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub SetUpGrid()
   Dim i As Integer

   Grid.Cols = NCOLS + 1
   Grid.rows = 12
      
   Call FGrSetup(Grid)
      
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_IDITEM) = 0
   Grid.ColWidth(C_CONCEPTO) = 5000
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_COLOBLIGATORIA) = 0
   Grid.ColWidth(C_UPD) = 0
   
   For i = C_REMEJANTNOMINAL To C_REMEJSGTE
      Grid.ColWidth(i) = 1600
      Grid.ColAlignment(i) = flexAlignRightCenter
   Next i
   
   Grid.FlxGrid.BackColor = vbButtonFace
   
End Sub

Private Sub SaveAll()
   Dim Row As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxId As Long
   Dim Col As Integer
   
   For Row = Grid.FixedRows To Grid.rows - 1
            
      If Val(Grid.TextMatrix(Row, C_IDITEM)) > 0 Then
      
         If Val(Grid.TextMatrix(Row, C_ID)) <> 0 Then
            Q1 = "UPDATE AsistImpPrimCat SET "
            Q1 = Q1 & "  RemEjAntNominal = " & vFmt(Grid.TextMatrix(Row, C_REMEJANTNOMINAL))
            Q1 = Q1 & ", RemEjAntAct = " & vFmt(Grid.TextMatrix(Row, C_REMEJANTACT))
            Q1 = Q1 & ", GeneradoAno = " & vFmt(Grid.TextMatrix(Row, C_GENERADOANO))
            Q1 = Q1 & ", CredUtilizado = " & vFmt(Grid.TextMatrix(Row, C_CREDUTILIZADO))
            Q1 = Q1 & ", RemEjSgte = " & vFmt(Grid.TextMatrix(Row, C_REMEJSGTE))
            Q1 = Q1 & " WHERE IdAsistImpPrimCat = " & Grid.TextMatrix(Row, C_ID)
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Else
            MaxId = 0
            Q1 = "SELECT Max(IdAsistImpPrimCat) FROM AsistImpPrimCat"
            Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               MaxId = vFld(Rs(0)) + 1
            End If
            Call CloseRs(Rs)

            Q1 = "INSERT INTO AsistImpPrimCat (IdAsistImpPrimCat, IdItem, RemEjAntNominal, RemEjAntAct, GeneradoAno, CredUtilizado, RemEjSgte, IdEmpresa, Ano )"
            Q1 = Q1 & " VALUES(" & MaxId & ", " & Grid.TextMatrix(Row, C_IDITEM)
            Q1 = Q1 & ", " & vFmt(Grid.TextMatrix(Row, C_REMEJANTNOMINAL))
            Q1 = Q1 & ", " & vFmt(Grid.TextMatrix(Row, C_REMEJANTACT))
            Q1 = Q1 & ", " & vFmt(Grid.TextMatrix(Row, C_GENERADOANO))
            Q1 = Q1 & ", " & vFmt(Grid.TextMatrix(Row, C_CREDUTILIZADO))
            Q1 = Q1 & ", " & vFmt(Grid.TextMatrix(Row, C_REMEJSGTE))
            Q1 = Q1 & ", " & gEmpresa.id
            Q1 = Q1 & ", " & gEmpresa.Ano & ")"
            
         End If
         
         Call ExecSQL(DbMain, Q1)
      End If
      
   Next Row

End Sub
Private Function Valida() As Boolean
   Valida = False
   Valida = True
End Function
Private Sub Bt_Print_Click()
   Dim PrtOrient As Integer
   
   PrtOrient = Printer.Orientation
   
   Call SetUpPrtGrid
   
   Me.MousePointer = vbHourglass
   
   Call gPrtReportes.PrtFlexGrid(Printer)
   
   Me.MousePointer = vbDefault
   
   Printer.Orientation = PrtOrient
   gPrtReportes.FmtCol = -1
   Call ResetPrtBas(gPrtReportes)

End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim PrtOrient As Integer
   Dim Pag As Integer
      
   PrtOrient = Printer.Orientation
      
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
   Pag = gPrtReportes.PrtFlexGrid(Frm)
   Set Frm.PrtControl = Bt_Print
   Me.MousePointer = vbDefault
   
   If Pag >= 0 Then
      Call Frm.FView(Caption)
   End If
   
   Set Frm = Nothing
   
   Printer.Orientation = PrtOrient
   gPrtReportes.FmtCol = -1
   Call ResetPrtBas(gPrtReportes)
   
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer, j As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Titulos(0) As String
   
   Printer.Orientation = ORIENT_VER
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Me.Caption
   gPrtReportes.Titulos = Titulos
            
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i) * 0.8
   Next i
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_COLOBLIGATORIA
   gPrtReportes.FmtCol = C_FMT
   gPrtReportes.NTotLines = 0
      
   
End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Clip As String

   Clip = LP_FGr2String(Grid, Me.Caption)
   Clipboard.Clear
   Clipboard.SetText Clip
      
End Sub


Private Sub Form_Load()

   Call SetUpGrid
   Call LoadAll

End Sub
Private Sub LoadAll()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Row As Integer
   Dim i As Integer
   Dim Valor As Double
   Dim TopeUTM As Double, ValUTM As Double
   Dim Fecha As Long
   Dim RemEjAntNominal As Double, GeneradoAno As Double
   
   Grid.FlxGrid.Redraw = False
   Grid.rows = 0
   Row = -1
   
   For i = 1 To C_MAX_ASISTIMPPRIMCAT
   
      If gStrAsistImpPrimCat(i) = "" Then
         Exit For
      End If
      
      Row = Row + 1
      Grid.rows = Row + 1
      Grid.TextMatrix(Row, C_IDITEM) = i
      Grid.TextMatrix(Row, C_CONCEPTO) = gStrAsistImpPrimCat(i)
      
      'obtenemos el ID y el valor guardado
      Q1 = "SELECT IdAsistImpPrimCat, RemEjAntNominal, GeneradoAno FROM AsistImpPrimCat "
      Q1 = Q1 & " WHERE IdItem = " & Grid.TextMatrix(Row, C_IDITEM)
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         Grid.TextMatrix(Row, C_ID) = vFld(Rs("IdAsistImpPrimCat"))
         RemEjAntNominal = vFld(Rs("RemEjAntNominal"))
         GeneradoAno = vFld(Rs("GeneradoAno"))
      End If
      Call CloseRs(Rs)
      
      Select Case i
                           
         Case 1      'IDPC Sobre Base Imponible
            'obtenemos el valor guardado en Base Imponible
            Q1 = "SELECT Valor FROM BaseImponible14Ter "
            Q1 = Q1 & " WHERE TipoBaseImp = " & BASEIMP_TOTALES & " AND IdItemBaseImp = 0"
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               lIDPCBaseImp = vFld(Rs("Valor")) * gImpPrimCategoria
               Grid.TextMatrix(Row, C_CREDUTILIZADO) = Format(lIDPCBaseImp, NUMFMT)
               
               lIDPCBaseImp = vFmt(Grid.TextMatrix(Row, C_CREDUTILIZADO))           'Para evitar problemas de redondeo
               
            End If
            Call CloseRs(Rs)
            
            Grid.TextMatrix(Row, C_FMT) = "B"
            Call FGrSetRowStyle(Grid, Row, "B")
            Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
            
            Row = Row + 2
            Grid.rows = Row + 2
            Grid.TextMatrix(Row - 1, C_FMT) = "L"
            
            'obtenemos el valor guardado en Base Imponible para Mayor Valor Enajenación, para usarlo en el cálculo
            Q1 = "SELECT Valor FROM BaseImponible14Ter "
            Q1 = Q1 & " WHERE TipoBaseImp = " & BASEIMP_TOTALES & " AND IdItemBaseImp = 1"
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               lMayorValorEnajenación = vFld(Rs("Valor"))
            End If
            Call CloseRs(Rs)
         
         
         Case 2      'Créditos contra Impuesto de Primera Categoría
            
            Grid.TextMatrix(Row - 1, C_REMEJANTNOMINAL) = "Rem. Ejer. Ant."
            Grid.TextMatrix(Row, C_REMEJANTNOMINAL) = "Nominal"
            Grid.TextMatrix(Row - 1, C_REMEJANTACT) = "Rem. Ejer. Ant."
            Grid.TextMatrix(Row, C_REMEJANTACT) = "Actualizado"
            Grid.TextMatrix(Row - 1, C_GENERADOANO) = "Generado"
            Grid.TextMatrix(Row, C_GENERADOANO) = "en el Año"
            Grid.TextMatrix(Row - 1, C_CREDUTILIZADO) = "Crédito"
            Grid.TextMatrix(Row, C_CREDUTILIZADO) = "Utilizado"
            Grid.TextMatrix(Row - 1, C_REMEJSGTE) = "Rem. Ejer."
            Grid.TextMatrix(Row, C_REMEJSGTE) = "Siguiente"
            
            Grid.TextMatrix(Row - 1, C_FMT) = "B"
            Grid.TextMatrix(Row, C_FMT) = "B"
            Call FGrSetRowStyle(Grid, Row - 1, "B")
            Call FGrSetRowStyle(Grid, Row, "B")
            Grid.Row = Row - 1
            Grid.Col = C_REMEJANTNOMINAL
            Grid.CellAlignment = flexAlignCenterCenter
            Grid.Col = C_REMEJANTACT
            Grid.CellAlignment = flexAlignCenterCenter
            Grid.Col = C_GENERADOANO
            Grid.CellAlignment = flexAlignCenterCenter
            Grid.Col = C_CREDUTILIZADO
            Grid.CellAlignment = flexAlignCenterCenter
            Grid.Col = C_REMEJSGTE
            Grid.CellAlignment = flexAlignCenterCenter
            Grid.Row = Row
            Grid.Col = C_REMEJANTNOMINAL
            Grid.CellAlignment = flexAlignCenterCenter
            Grid.Col = C_REMEJANTACT
            Grid.CellAlignment = flexAlignCenterCenter
            Grid.Col = C_GENERADOANO
            Grid.CellAlignment = flexAlignCenterCenter
            Grid.Col = C_CREDUTILIZADO
            Grid.CellAlignment = flexAlignCenterCenter
            Grid.Col = C_REMEJSGTE
            Grid.CellAlignment = flexAlignCenterCenter
            
            
         Case 3
            Grid.TextMatrix(Row, C_FMT) = "L"
            
            Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT, C_REMEJANTNOMINAL, C_REMEJANTACT)
            Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT, C_REMEJSGTE, C_REMEJSGTE)
            Valor = GetValAjustesELC(TAEC_AGREGADOS, 6)     'Crédito 33 bis ELC
            
            'topamos  valor máx: 500 UTM
            Fecha = DateSerial(gEmpresa.Ano, 12, 31)
               
            If GetValMoneda("UTM", ValUTM, Fecha, False) = True Then     'obtiene la última UTM ingresada en el sistema que sea a lo más del 31 dic del año actual
               TopeUTM = vFmt(Format(500 * ValUTM, NUMFMT))   'redondeamos a 0 decimales
         
               If Valor > TopeUTM Then
                  Valor = TopeUTM
               End If
            Else
               MsgBox1 "No se encontró el valor de la UTM para calcular tope de Crédito 33 bis.", vbExclamation
         
            End If
            
            lGeneradoAno = Valor
            Grid.TextMatrix(Row, C_GENERADOANO) = Format(lGeneradoAno, NUMFMT)
            
            'min entre IDPC sobre Base Imponible y Generado Año
            If lIDPCBaseImp < lGeneradoAno Then
               Valor = lIDPCBaseImp
            Else
               Valor = lGeneradoAno
            End If
            
            lCred33bisUtilizado = Valor
            
            Grid.TextMatrix(Row, C_CREDUTILIZADO) = Format(lCred33bisUtilizado, NUMFMT)
            
            lCred33bisUtilizado = vFmt(Grid.TextMatrix(Row, C_CREDUTILIZADO))     'para evitar problemas de redondeo
            
            
         Case 4
         
            Grid.TextMatrix(Row, C_REMEJANTNOMINAL) = Format(RemEjAntNominal, NUMFMT)
            Grid.TextMatrix(Row, C_GENERADOANO) = Format(GeneradoAno, NUMFMT)
            
            Call FGrSetRowStyle(Grid, Row, "BC", vbWhite, C_REMEJANTNOMINAL, C_REMEJANTNOMINAL)
            Call FGrSetRowStyle(Grid, Row, "BC", vbWhite, C_GENERADOANO, C_GENERADOANO)
            lRowCredIng = Row
            
         Case 5
            
            Grid.TextMatrix(Row, C_REMEJANTNOMINAL) = Format(RemEjAntNominal, NUMFMT)
            Grid.TextMatrix(Row, C_GENERADOANO) = Format(GeneradoAno, NUMFMT)
            
            Call FGrSetRowStyle(Grid, Row, "BC", vbWhite, C_REMEJANTNOMINAL, C_REMEJANTNOMINAL)
            Call FGrSetRowStyle(Grid, Row, "BC", vbWhite, C_GENERADOANO, C_GENERADOANO)
            lRowCredRetiros = Row
            Row = Row + 1
            Grid.rows = Row + 1
            
         Case 6
            Grid.TextMatrix(Row - 1, C_FMT) = "L"
            
            lRowIDPCNetoPagar = Row
            Grid.TextMatrix(Row, C_FMT) = "B"
            Call FGrSetRowStyle(Grid, Row, "B")
            Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
            Row = Row + 1
            Grid.rows = Row + 1
            
         Case 7
            Grid.TextMatrix(Row - 1, C_FMT) = "L"
            
            lRowMayorValorEnajenacion = Row
            Grid.TextMatrix(Row, C_FMT) = "B"
            Call FGrSetRowStyle(Grid, Row, "B")
            Row = Row + 1
            Grid.rows = Row + 1
            
         Case 8
            lRowIDPCPagar = Row
            Grid.TextMatrix(Row, C_FMT) = "B"
            Call FGrSetRowStyle(Grid, Row, "B")
            Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
            
            
      End Select
      
      Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
      
   Next i
      
   Call CalcTot
   
   For i = 0 To Grid.rows - 1
      Grid.TextMatrix(i, C_COLOBLIGATORIA) = "."
   Next i

   Call FGrVRows(Grid)
   Grid.rows = Grid.rows - 2
   
   Grid.Col = C_CONCEPTO
   Grid.Row = 0
   Grid.FlxGrid.Redraw = True



End Sub


Private Sub CalcTot()
   Dim Valor As Double
   Dim Factor As Single
   Dim Val1 As Double, Val2 As Double
   
   Factor = GetFactorCM(DateSerial(gEmpresa.Ano - 1, 12, 1))
   
   'Remanente Ejercicio Anterior Actualizado - Crédito Ingreso Diferido
   Valor = vFmt(Grid.TextMatrix(lRowCredIng, C_REMEJANTNOMINAL)) * Factor
   Grid.TextMatrix(lRowCredIng, C_REMEJANTACT) = Format(Valor, NUMFMT)
   
   'Remanente Ejercicio Anterior Actualizado - Crédito Retiros
   Valor = vFmt(Grid.TextMatrix(lRowCredRetiros, C_REMEJANTNOMINAL)) * Factor
   Grid.TextMatrix(lRowCredRetiros, C_REMEJANTACT) = Format(Valor, NUMFMT)
   
   'Crédito Utilizado - Crédito Ingreso Diferido
   Val1 = lIDPCBaseImp - lCred33bisUtilizado
   Val2 = vFmt(Grid.TextMatrix(lRowCredIng, C_REMEJANTACT)) + vFmt(Grid.TextMatrix(lRowCredIng, C_GENERADOANO))
   
   If Val1 < Val2 Then
      Valor = Val1
   Else
      Valor = Val2
   End If
   Grid.TextMatrix(lRowCredIng, C_CREDUTILIZADO) = Format(Valor, NUMFMT)
   
   'Crédito Utilizado - Crédito Retiros
   Val1 = lIDPCBaseImp - lCred33bisUtilizado - vFmt(Grid.TextMatrix(lRowCredIng, C_CREDUTILIZADO))
   Val2 = vFmt(Grid.TextMatrix(lRowCredRetiros, C_REMEJANTACT)) + vFmt(Grid.TextMatrix(lRowCredRetiros, C_GENERADOANO))
   
   If Val1 < Val2 Then
      Valor = Val1
   Else
      Valor = Val2
   End If
   Grid.TextMatrix(lRowCredRetiros, C_CREDUTILIZADO) = Format(Valor, NUMFMT)
   
   'Remanente Ejercicio Siguiente - Credito Retiros
   Valor = vFmt(Grid.TextMatrix(lRowCredRetiros, C_REMEJANTACT)) + vFmt(Grid.TextMatrix(lRowCredRetiros, C_GENERADOANO)) - vFmt(Grid.TextMatrix(lRowCredRetiros, C_CREDUTILIZADO))
   Grid.TextMatrix(lRowCredRetiros, C_REMEJSGTE) = Format(Valor, NUMFMT)
   
   'Cred. Utilizado - Neto a Pagar
   Valor = lIDPCBaseImp - lCred33bisUtilizado - vFmt(Grid.TextMatrix(lRowCredIng, C_CREDUTILIZADO)) - vFmt(Grid.TextMatrix(lRowCredRetiros, C_CREDUTILIZADO))
   Grid.TextMatrix(lRowIDPCNetoPagar, C_CREDUTILIZADO) = Format(Valor, NUMFMT)
  
   'Cred. Utilizado - Mayor Valor Enajenación
   Valor = lMayorValorEnajenación * gImpPrimCategoria
   Grid.TextMatrix(lRowMayorValorEnajenacion, C_CREDUTILIZADO) = Format(Valor, NUMFMT)
  
   'Cred. Utilizado - IDPG a Pagar
   Valor = Valor * gImpPrimCategoria
   Grid.TextMatrix(lRowIDPCPagar, C_CREDUTILIZADO) = Format(Valor, NUMFMT)
   
End Sub


Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)

   'Crédito asociado a ingreso diferido o Crédito asociado a retiros, dividendos y participaciones percibidas
   Value = Format(vFmt(Value), NUMFMT)
   Grid.TextMatrix(Row, Col) = Value
   Call CalcTot

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid2.FEG2_EdType)
   
   'Crédito asociado a ingreso diferido o Crédito asociado a retiros, dividendos y participaciones percibidas
   If (Row = lRowCredIng Or Row = lRowCredRetiros) And (Col = C_REMEJANTNOMINAL Or Col = C_GENERADOANO) Then
      EdType = FEG_Edit
   End If
      
End Sub

Private Sub Grid_DblClick()
   Call Bt_DetDoc_Click
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
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

