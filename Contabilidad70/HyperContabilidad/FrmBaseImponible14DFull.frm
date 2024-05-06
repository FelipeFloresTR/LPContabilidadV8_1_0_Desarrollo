VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmBaseImponible14DFull 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base Imponible Primera Categoría Reg 14 D"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11895
   StartUpPosition =   1  'CenterOwner
   Begin FlexEdGrid3.FEd3Grid Grid 
      Height          =   7875
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   13891
      Cols            =   2
      Rows            =   3
      FixedCols       =   1
      FixedRows       =   1
      ScrollBars      =   3
      AllowUserResizing=   1
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
      TabIndex        =   11
      Top             =   0
      Width           =   11835
      Begin VB.CommandButton Bt_SaldosVig 
         Caption         =   "Saldos Vigentes"
         Height          =   315
         Left            =   5940
         TabIndex        =   9
         Top             =   180
         Width           =   1515
      End
      Begin VB.CommandButton Bt_Expand 
         Caption         =   "Expandir Todo"
         Height          =   315
         Left            =   4140
         TabIndex        =   8
         Top             =   180
         Width           =   1515
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10560
         TabIndex        =   10
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
         Left            =   540
         Picture         =   "FrmBaseImponible14DFull.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
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
         Left            =   120
         Picture         =   "FrmBaseImponible14DFull.frx":04BA
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
         Picture         =   "FrmBaseImponible14DFull.frx":0961
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
         Picture         =   "FrmBaseImponible14DFull.frx":0DA6
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
         Picture         =   "FrmBaseImponible14DFull.frx":11CF
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
         Picture         =   "FrmBaseImponible14DFull.frx":156D
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
         Picture         =   "FrmBaseImponible14DFull.frx":18CE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmBaseImponible14DFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDTBLBASEIMP14D = 0
Const C_IDARRBASEIMP14D = 1
Const C_REGIMEN = 2
Const C_NIVEL = 3
Const C_TIPO = 4
Const C_CODIGO = 5
Const C_FORMAINGRESO = 6
Const C_OPENCLOSE = 7
Const C_DESCRIP = 8
Const C_VALOR = 9
Const C_FMT = 10
Const C_UPDATE = 11

Const NCOLS = C_UPDATE

Dim lRc As Integer
Dim lBaseImponible As Double
Dim lRowCod8300 As Integer


Public Function FEdit(BaseImponible As Double) As Integer

   Me.Show vbModal
   
   BaseImponible = 0
   
   If lRc = vbOK Then
      BaseImponible = lBaseImponible
   End If
   
   FEdit = lRc
   
End Function
'
'
'Private Sub Bt_BaseImpAcum_Click()
'   Dim Frm As FrmDetCapPropioSimplAcum
'   Dim Rc As Integer
'   Dim Valor As Double
'
'   If Valida() Then
'      Call SaveAll
'
'      Set Frm = New FrmDetCapPropioSimplAcum
'      Rc = Frm.FEdit(CPS_BASEIMPONIBLE, Valor)
'
'      If Rc = vbOK Then
'
'         If gEmpresa.ProPymeGeneral <> 0 Then
'            Grid.TextMatrix(R_14DN3, C_VALOR) = Format(Valor, NUMFMT)
'         Else
'            Grid.TextMatrix(R_14DN8, C_VALOR) = Format(Valor, NUMFMT)
'         End If
'
'      End If
'
'      Set Frm = Nothing
'   End If
'
'End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   
   Unload Me
End Sub

Private Sub Bt_Expand_Click()
   Call ExpandAll
End Sub

Private Sub Bt_OK_Click()

   If valida() Then
      Call SaveAll
   
      lRc = vbOK
      
      Unload Me
   End If
   
End Sub

Private Sub Bt_SaldosVig_Click()
   Call SaldosVigentes
End Sub

Private Sub Form_Load()
   
   Call SetUpGrid
   
   
   Call LoadBase
   Call LoadAll
   
   If gEmpresa.FCierre <> 0 Then
      MsgBox1 "Esta período ya ha sido cerrado. No podrá modificar ingresos manuales.", vbExclamation + vbOKOnly
      Exit Sub
   End If

   
End Sub
Private Sub LoadBase()
   Dim i As Integer
   Dim Row As Integer

   Grid.Redraw = False
   
   Row = Grid.FixedRows
   Grid.rows = Grid.FixedRows
   
   
   For i = 1 To UBound(gBaseImponible14D)
      
      If gBaseImponible14D(i).Nivel = 0 Then
         Exit For
      End If
      
      If (gBaseImponible14D(i).Regimen = FTE_14DN3 And Not gEmpresa.ProPymeGeneral) Or (gBaseImponible14D(i).Regimen = FTE_14DN8 And Not gEmpresa.ProPymeTransp) Then
         If gBaseImponible14D(i).Codigo <> 1700 Then
            If gBaseImponible14D(i).Nivel = BIMP14D_MAXNIV Then
               ClearDetBaseImp14D (gBaseImponible14D(i).Codigo)
            End If
            GoTo NextRow
         End If
         
      End If
      
      'condición en DURO dada la especificidad
      If (gBaseImponible14D(i).Codigo = 4200 Or gBaseImponible14D(i).Codigo = 4300) Then
         If gEmpresa.Ano = 2020 And Not gEmpresa.ProPymeTransp Then
            If gBaseImponible14D(i).Nivel = BIMP14D_MAXNIV Then
               ClearDetBaseImp14D (gBaseImponible14D(i).Codigo)
            End If
            GoTo NextRow
         End If
      End If
      
      
      Grid.rows = Grid.rows + 1
      If (gBaseImponible14D(i).Codigo = 1700 Or gBaseImponible14D(i).Codigo = 4700) Then
         If gEmpresa.Ano > 2021 And gEmpresa.ProPymeTransp Then
            Grid.RowHeight(Row) = 0
         End If
      End If
      
      
      
      If gBaseImponible14D(i).Nivel <= 2 And Row > Grid.FixedRows Then
         Grid.rows = Grid.rows + 1
         Row = Row + 1
      End If

      Grid.TextMatrix(Row, C_IDARRBASEIMP14D) = i
      Grid.TextMatrix(Row, C_REGIMEN) = gBaseImponible14D(i).Regimen
      Grid.TextMatrix(Row, C_TIPO) = gBaseImponible14D(i).Tipo
      Grid.TextMatrix(Row, C_NIVEL) = gBaseImponible14D(i).Nivel
      Grid.TextMatrix(Row, C_FORMAINGRESO) = gBaseImponible14D(i).FormaIngreso
      Grid.TextMatrix(Row, C_CODIGO) = gBaseImponible14D(i).Codigo
      
      If gBaseImponible14D(i).Nivel <= 4 And gBaseImponible14D(i).Nivel > 1 Then
         Grid.TextMatrix(Row, C_OPENCLOSE) = "-"
         Call FGrFontBold(Grid, Row, C_OPENCLOSE, True)
      End If
      Grid.TextMatrix(Row, C_DESCRIP) = String((gBaseImponible14D(i).Nivel - 1) * 4, " ") & gBaseImponible14D(i).Nombre
      
      If gBaseImponible14D(i).Nivel <= 2 Then
         Grid.TextMatrix(Row, C_DESCRIP) = UCase(Grid.TextMatrix(Row, C_DESCRIP))
      End If
   
      If gBaseImponible14D(i).Nivel <= 3 Then
         Call FGrFontBold(Grid, Row, -1, True)
         Grid.TextMatrix(Row, C_FMT) = "B"
      Else
         Grid.TextMatrix(Row, C_DESCRIP) = String((gBaseImponible14D(i).Nivel - 1) * 2, " ") & Grid.TextMatrix(Row, C_DESCRIP)
      End If
      
      If gBaseImponible14D(i).Nivel = 4 Then
         Call FGrForeColor(Grid, Row, -1, vbBlue)
         Grid.TextMatrix(Row, C_FMT) = "FCELL"
      End If
      
      If gBaseImponible14D(i).Nivel = 5 And gBaseImponible14D(i).FormaIngreso <> ING_MANUAL Then
         Call FGrBackColor(Grid, Row, C_VALOR, COLOR_GRISLTLT)
      End If
      
      Grid.TextMatrix(Row, C_VALOR) = 0
      
      Row = Row + 1
NextRow:

   Next i

'   Call OcultarSegunRegimen
   
   Grid.rows = Grid.rows + 1
   
   Grid.Redraw = True
   
End Sub


Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Total As Double
   Dim Valor As Double
   Dim TipoComp As Integer, TipoAjuste As Integer
   Dim Row As Integer
   Dim Val7400 As Double
   Dim Val7500 As Double
   Dim Val7600 As Double
   Dim Val7700 As Double
   Dim Val7800 As Double
   Dim Val7900 As Double
   Dim Val8100 As Double
   Dim TipoOperCaja As Integer
   Dim Regimen As Integer
   Dim porcEgresoExistOServi As Integer
   Dim valorExistente, valorPorcExistente As Long
   
   
   '3405135 se agregar indice para mejorar tiempo de respuesta en script solo en sql
    If gDbType = SQL_SERVER Then
   Q1 = ""
   Q1 = Q1 & "if not exists (select name from sysindexes  where name = 'idx_documentosCuenTotal') "
   Q1 = Q1 & "CREATE NONCLUSTERED INDEX idx_documentosCuenTotal ON Documento "
   Q1 = Q1 & "(IdDoc ASC,IdEmpresa ASC,Ano ASC,IdCuentaTotal ASC,IdCuentaExento ASC,IdCuentaAfecto ASC); "
   
   Call ExecSQL(DbMain, Q1)
   End If
   '3405135

   valorExistente = 0
   porcEgresoExistOServi = 100
' 2738256 Correccion SQL Server
'   Q1 = "SELECT PorcExisteOServ "
'   Q1 = Q1 & " FROM Empresa "
'   Q1 = Q1 & " WHERE Id = " & gEmpresa.id
'
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF = False Then
'      porcEgresoExistOServi = IIf(vFld(Rs("PorcExisteOServ")) = "", 100, vFld(Rs("PorcExisteOServ")))
'   End If
'   Call CloseRs(Rs)


   Q1 = "SELECT IdBaseImponible14D, Tipo, Nivel, Codigo, Valor "
   Q1 = Q1 & " FROM BaseImponible14D "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   ' ADO 2747807 Tema 1 y 5 Oculta servicios pagados a partir del 2021
   If gEmpresa.Ano >= 2021 Then
    If gEmpresa.Ano > 2021 And gEmpresa.ProPymeTransp Then
       If gEmpresa.Ano >= 2023 Then
        Q1 = Q1 & " AND Codigo not in (5300,10000) "
       Else
        Q1 = Q1 & " AND Codigo not in (5300,10420) "
       End If
    Else
      '3402617
      If gEmpresa.Ano >= 2023 And gEmpresa.ProPymeGeneral Then
        Q1 = Q1 & " AND Codigo not in (5300,10000) "
      Else
        Q1 = Q1 & " AND Codigo not in (5300) "
      End If
      '3402617
    End If
   Else
     Q1 = Q1 & " AND Codigo not in (401, 5201) "
   End If
   Q1 = Q1 & " ORDER BY Codigo"

   Set Rs = OpenRs(DbMain, Q1)

   Grid.FlxGrid.Redraw = False
   i = Grid.FixedRows


   'Cargamos los datos en la taba
   Do While Not Rs.EOF And i < Grid.rows
      If vFld(Rs("Codigo")) = 1700 Then
        i = i
      End If
      
      
      Do While Not Val(Grid.TextMatrix(i, C_CODIGO)) = vFld(Rs("Codigo")) And i < Grid.rows
         i = i + 1
      Loop

      If Val(Grid.TextMatrix(i, C_CODIGO)) = vFld(Rs("Codigo")) Then
         If Val(Grid.TextMatrix(i, C_NIVEL)) = vFld(Rs("Nivel")) Then
            Grid.TextMatrix(i, C_IDTBLBASEIMP14D) = vFld(Rs("IdBaseImponible14D"))
            If Val(Grid.TextMatrix(i, C_TIPO)) = BIMP14D_INGRESO Then
               Grid.TextMatrix(i, C_VALOR) = Format(Abs(vFld(Rs("Valor"))), NUMFMT)
            Else
               Grid.TextMatrix(i, C_VALOR) = Format(Abs(vFld(Rs("Valor"))) * -1, NUMFMT)
            End If
         End If
         Rs.MoveNext
      End If

   Loop

   Call CloseRs(Rs)

   'Ahora actualizamos los traspasos

   For Row = Grid.FixedRows To Grid.rows - 1

      If Val(Grid.TextMatrix(Row, C_TIPO)) = BIMP14D_INGRESO Then
         TipoComp = TC_INGRESO
         TipoAjuste = TAEC_AGREGADOS
         TipoOperCaja = TOPERCAJA_INGRESO
      Else
         TipoComp = TC_EGRESO
         TipoAjuste = TAEC_DEDUCCIONES
         TipoOperCaja = TOPERCAJA_EGRESO
      End If
      
      If Val(Grid.TextMatrix(Row, C_CODIGO)) = 1700 Then
        i = i
      End If

      Select Case Grid.TextMatrix(Row, C_FORMAINGRESO)

         Case ING_TRASPASOAJUSTE, ING_AMBOSAJUSTE    'ING_AMBOSAJUSTE sólo implica Traspaso dado que no se puede dar la opción de traspaso o ingreso manual
            Valor = LoadValCuentasAjustes14D(TipoAjuste, gBaseImponible14D(Val(Grid.TextMatrix(Row, C_IDARRBASEIMP14D))).IdItemCtasAsociadasAjustes, TipoComp)
            If Val(Grid.TextMatrix(Row, C_TIPO)) = BIMP14D_INGRESO Then
               Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            Else
               Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)
            End If

         Case ING_TRASPASOLIBCAJA

            If Val(Grid.TextMatrix(Row, C_TIPO)) = BIMP14D_INGRESO Then

               If InStr(Grid.TextMatrix(Row, C_DESCRIP), "Ingresos percibidos del Giro") > 0 Then
                  Regimen = FTE_14DN3
                  Valor = GetPercibidosPagados(TipoOperCaja, Regimen)
                  Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
               ElseIf InStr(Grid.TextMatrix(Row, C_DESCRIP), "Ingresos devengados o percibidos") > 0 Then
                  Regimen = FTE_14A
                  Valor = GetPercibidosPagados(TipoOperCaja, Regimen)
                  Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
                 ' ADO 2747807 Tema 3
               ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 401 Then
                  Regimen = FTE_14DN3
                  Valor = GetDevengadosAnteriores(TipoOperCaja, Regimen)
                  Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
               End If

            Else   'Egresos
               Valor = ExistOInsumPagados(TipoOperCaja)
               Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)

               ' ADO 2747807 Tema 5
               If Val(Grid.TextMatrix(Row, C_CODIGO)) = 5201 Then
                  Valor = ExistOInsumPagados(TipoOperCaja, False)
                  Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)
               End If

               If InStr(Grid.TextMatrix(Row, C_DESCRIP), "Gastos Adeudados asociados a ingresos") > 0 Then
                  Regimen = FTE_14DN3
                  Valor = GetPercibidosPagadosCompras(TipoOperCaja, Regimen)
                  Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
               End If


            End If

         Case ING_TRASPASO

            'INGRESOS
            If Trim(Grid.TextMatrix(Row, C_DESCRIP)) = "Otros ingresos percibidos" Then
               Valor = GetTotCta_CodF22_14Ter(651, "C") + LoadValCuentasAjustes14D(TAEC_AGREGADOS, 15) + LoadValCuentasAjustes14D(TAEC_AGREGADOS, 19)
               Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)

            ElseIf Trim(Grid.TextMatrix(Row, C_DESCRIP)) = "Ingresos percibidos por la enajenación de bienes depreciables" Then
               Valor = LoadValCuentasAjustes14D(TAEC_AGREGADOS, 14)
               Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)

            ElseIf Trim(Grid.TextMatrix(Row, C_DESCRIP)) = "Monto Liquido recibido" Then
               Valor = GetTotCta_CodF22_14Ter(629, "C") + GetValAjustesELC(TAEC_AGREGADOS, 8) + GetValAjustesELC(TAEC_AGREGADOS, 9)
               Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)

            'ElseIf InStr(Grid.TextMatrix(Row, C_DESCRIP), "Ingresos devengados con empresas relacionadas") > 0 Then
            '   Valor = LoadValCuentasAjustes14D(TAEC_AGREGADOS, 16) + LoadValCuentasAjustes14D(TAEC_AGREGADOS, 17) + LoadValCuentasAjustes14D(TAEC_AGREGADOS, 18)
            '   Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)

            ElseIf Trim(Grid.TextMatrix(Row, C_DESCRIP)) = "Crédito art. 33 Bis LIR" Then
               Valor = GetVal33Bis()
               Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)

            'EGRESOS
            ElseIf Trim(Grid.TextMatrix(Row, C_DESCRIP)) = "Compra de activos fijos depreciables" Then
'               Valor = GetTotCta_CodF22_14Ter(632, "D") + LoadValCuentasAjustes14D(TAEC_DEDUCCIONES, 13) + LoadValCuentasAjustes14D(TAEC_DEDUCCIONES, 14)
               'Valor = GetTotCta_CodF22_14Ter(632, "D") + GetValAjustesELC(TAEC_DEDUCCIONES, 13) + GetValAjustesELC(TAEC_DEDUCCIONES, 14)
               Valor = GetActivosFijosDepreciables(TipoOperCaja)
               Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)

            ElseIf InStr(Grid.TextMatrix(Row, C_DESCRIP), "Remuneraciones pagadas") > 0 Then
               Valor = GetTotCta_CodF22_14Ter(631, "D")
               Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)

             ElseIf Trim(Grid.TextMatrix(Row, C_DESCRIP)) = "Intereses pagados por préstamos" Then
               Valor = GetTotCta_CodF22_14Ter(633, "D")
               Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)

             ElseIf Trim(Grid.TextMatrix(Row, C_DESCRIP)) = "Impuestos que no sean de la LIR" Then
               Valor = GetIVAIrrecSaldoPagar() + LoadValCuentasAjustes14D(TipoAjuste, gBaseImponible14D(Val(Grid.TextMatrix(Row, C_IDARRBASEIMP14D))).IdItemCtasAsociadasAjustes, TipoComp)
               Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)

             ElseIf Trim(Grid.TextMatrix(Row, C_DESCRIP)) = "Arriendos pagados" Then
               Valor = GetTotCta_CodF22_14Ter(1140, "D")
               Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)

             ElseIf InStr(Grid.TextMatrix(Row, C_DESCRIP), "Partidas pagadas del inciso 1° del art 21") > 0 Then
               Valor = GetTotCta_CodF22_14Ter(1144, "D")
               Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)

             ElseIf InStr(Grid.TextMatrix(Row, C_DESCRIP), "Gastos adeudados asociados a ingresos devengados") > 0 Then
               'Valor = GetTotCta_CodF22_14Ter(630, "D") - LoadValCuentasAjustes14D(TAEC_AGREGADOS, 1) - GetTotCta_CodF22_14Ter_NC(630, "C")
               '2699582
               If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
                 Regimen = FTE_14A
                  Valor = GetPercibidosPagadosCompras(TipoOperCaja, Regimen)
               Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)

               Else
                 Valor = GetTotCta_CodF22_14Ter(630, "D") - LoadValCuentasAjustes14D(TAEC_AGREGADOS, 1) - GetTotCta_CodF22_14Ter_NC(630, "C")
               Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)
               End If

                'fin 2699582

             ElseIf Trim(Grid.TextMatrix(Row, C_DESCRIP)) = "Otras deducciones a la RLI" Then
               Valor = GetTotCta_CodF22_14Ter(635, "D") + GetValAjustesELC(TAEC_DEDUCCIONES, 17) + GetValAjustesELC(TAEC_DEDUCCIONES, 5) + GetValAjustesELC(TAEC_DEDUCCIONES, 15) + GetValAjustesELC(TAEC_DEDUCCIONES, 8)
               Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor) * -1, NUMFMT)

            End If

      End Select

      'traspaso de ajustes espejo
      If Val(Grid.TextMatrix(Row, C_CODIGO)) = 5200 Then
         valorExistente = vFmt(Grid.TextMatrix(Row, C_VALOR)) + (Abs(GetNCVParaExistencias()) * -1)
'         valorPorcExistente = vFmt(Grid.TextMatrix(Row, C_VALOR)) * (porcEgresoExistOServi / 100)
         Grid.TextMatrix(Row, C_VALOR) = Format(valorExistente, NUMFMT)
'      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 5300 Then
'        If valorExistente < 0 Then
'            Grid.TextMatrix(Row, C_VALOR) = "-" & Abs(Abs(valorExistente) - Abs(valorPorcExistente))
'         Else
'            Grid.TextMatrix(Row, C_VALOR) = Abs(Abs(valorExistente) - Abs(valorPorcExistente))
'         End If
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 7400 Then
         Val7400 = vFmt(Grid.TextMatrix(Row, C_VALOR))
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 7500 Then
         Val7500 = vFmt(Grid.TextMatrix(Row, C_VALOR))
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 7600 Then
         Val7600 = vFmt(Grid.TextMatrix(Row, C_VALOR))
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 7700 Then
         Val7700 = vFmt(Grid.TextMatrix(Row, C_VALOR))
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 7800 Then
         Val7800 = vFmt(Grid.TextMatrix(Row, C_VALOR))
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 7900 Then
         Val7900 = vFmt(Grid.TextMatrix(Row, C_VALOR))

      'Estos ítems, a pesar de estar en los EGRESOS, suman a la base imponible
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 10800 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Val7400), NUMFMT)
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 10900 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Val7500), NUMFMT)
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 11000 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Val7600), NUMFMT)
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 11100 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Val7700), NUMFMT)
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 11200 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Val7800), NUMFMT)
      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 11300 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Val7900), NUMFMT)

      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 8100 Then
         Val8100 = vFmt(Grid.TextMatrix(Row, C_VALOR))

      ElseIf Val(Grid.TextMatrix(Row, C_CODIGO)) = 8300 Then
         Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Val8100) * -1, NUMFMT)
         lRowCod8300 = Row

    '2699582
      ElseIf Trim(Grid.TextMatrix(Row, C_DESCRIP)) = "Reajuste de PPM" Then
           If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
               Valor = GetValPpm()
               Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
                     End If

      ElseIf Trim(Grid.TextMatrix(Row, C_CODIGO)) = 9600 Then
        Valor = 0
            If gEmpresa.Ano >= 2022 Then

               Valor = GetPerdidaAnoAnterior()
          End If
               Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
     'fin 2699582

      End If

   Next Row

   Grid.Row = Grid.FixedRows
   Grid.Col = C_DESCRIP

   Call CalcGrid
   Grid.FlxGrid.Redraw = True

End Sub
Private Function GetIVAIrrecSaldoPagar() As Double
   Dim Valor As Double
   Dim Rs As Recordset
   Dim Q1 As String
   Dim TmpTbl As String
   Dim Rc As Integer
   
   '2882638
   Dim Rs2 As Recordset
   Dim Q2 As String
   Dim Rs3 As Recordset
   Dim Q3 As String
   'fin 2882638
   
    '2970958
   Dim Rs4 As Recordset
   Dim Q4 As String
   '2970958
   
   Valor = 0
   
   TmpTbl = DbGenTmpName2(gDbType, "TBImp_")
   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)


   Q1 = "SELECT IdDoc, Sum(Pagado) as TotPagado, Afecto, Exento, IVA, IVAIrrec "
   Q1 = Q1 & " INTO " & TmpTbl
   Q1 = Q1 & " FROM LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc"
   Q1 = Q1 & " WHERE LibroCaja.TipoLib = " & LIB_COMPRAS
   Q1 = Q1 & " AND TipoDocs.Diminutivo IN( 'FAC', 'NDC', 'IMP')"
    '2853377
   'Q1 = Q1 & " AND Year(FechaIngresoLibro) = " & gEmpresa.Ano
     '3051436
   If gDbType = SQL_SERVER Then
    Q1 = Q1 & " AND Year(FechaIngresoLibro -2) in ( " & gEmpresa.Ano & "," & gEmpresa.Ano - 1 & ")"
   Else
     Q1 = Q1 & " AND Year(FechaIngresoLibro) in ( " & gEmpresa.Ano & "," & gEmpresa.Ano - 1 & ")"
   End If
    'fin 2853377
   
   'Q1 = Q1 & " AND IVAIrrec > 0"
   '2853377
   'Q1 = Q1 & " AND (otroimp > 0 or IVAIrrec > 0) and MontoAfectaBaseImp > 0 "
   Q1 = Q1 & " AND (otroimp > 0 or IVAIrrec > 0)"
   'fin 2853377
   
   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   '2805543
   Q1 = Q1 & " AND tipoOper = 2 "
   'fin 2805543
   
    '2852276
   Q1 = Q1 & " AND LibroCaja.pagado > 0 "
   'fin 2852276
     
   Q1 = Q1 & " GROUP BY IdDoc, Afecto, Exento, IVA, IVAIrrec"


   Rc = ExecSQL(DbMain, Q1)

 Call CloseRs(Rs)
 
   ' ***** ADO 2746994 Se suman los valores
   'Q1 = "SELECT  Iif( Iif(TotPagado - Afecto - Exento - IVA > 0, TotPagado - Afecto - Exento - IVA, 0) > IVAIrrec, IVAIrrec, Iif(TotPagado - Afecto - Exento - IVA > 0, TotPagado - Afecto - Exento - IVA, 0) ) As Total "
   'Q1 = "SELECT  Sum(Iif( Iif(TotPagado - Afecto - Exento - IVA > 0, TotPagado - Afecto - Exento - IVA, 0) > IVAIrrec, IVAIrrec, Iif(TotPagado - Afecto - Exento - IVA > 0, TotPagado - Afecto - Exento - IVA, 0) )) As Total "
   
   '2853377
   'Q1 = "SELECT  sum(TotPagado - ((afecto) + (iva) + (Exento))) As Total "
   Q1 = "SELECT Sum (IIf(TotPagado - ((Afecto) + (IVA) + (Exento)) < 0, 0, TotPagado - ((Afecto) + (IVA) + (Exento)))) As Total "
   'fin 2853377
   Q1 = Q1 & " FROM " & TmpTbl

'2752418
'   Q1 = "SELECT sum(pagado - ((afecto) + (iva))) As Total  "
'   Q1 = Q1 & " FROM LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc"
'   Q1 = Q1 & " WHERE LibroCaja.TipoLib = " & LIB_COMPRAS
'   Q1 = Q1 & " AND TipoDocs.Diminutivo IN( 'FAC', 'NDC', 'IMP')"
'   Q1 = Q1 & " AND Year(FechaIngresoLibro) = " & gEmpresa.Ano
'   Q1 = Q1 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
'   Q1 = Q1 & "  AND otroimp > 0 or IVAIrrec > 0 and MontoAfectaBaseImp >0 "
   

   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
    
    Valor = vFld(Rs("Total"))
     
     Call CloseRs(Rs)
     
      '2970958
    Q4 = "SELECT Sum(Afecto+ Exento+IVA+ IVAIrrec) as Valor "
   Q4 = Q4 & " FROM LibroCaja INNER JOIN TipoDocs ON LibroCaja.TipoLib = TipoDocs.TipoLib AND LibroCaja.TipoDoc = TipoDocs.TipoDoc"
   Q4 = Q4 & " WHERE LibroCaja.TipoLib = " & LIB_COMPRAS
   Q4 = Q4 & " AND TipoDocs.Diminutivo IN('NCC')"
    Q4 = Q4 & " AND Year(FechaIngresoLibro) in ( " & gEmpresa.Ano & "," & gEmpresa.Ano - 1 & ")"
   Q4 = Q4 & " AND (otroimp > 0 or IVAIrrec > 0)"
   
   Q4 = Q4 & " AND LibroCaja.IdEmpresa = " & gEmpresa.id & " AND LibroCaja.Ano = " & gEmpresa.Ano
   Q4 = Q4 & " AND tipoOper = 1 "
   '3056447
   'Q4 = Q4 & " AND LibroCaja.pagado = 0 "
   Q4 = Q4 & " AND LibroCaja.pagado > 0 "
   '3056447
   'Q4 = Q4 & " GROUP BY IdDoc, Afecto, Exento, IVA, IVAIrrec"
     
     
     Set Rs4 = OpenRs(DbMain, Q4)
     If Not Rs4.EOF Then
        Valor = Valor + vFld(Rs4("Valor"))
     End If
     
     Call CloseRs(Rs4)
     '2970958
     
     
     '2882638
     Q2 = "SELECT IdDoc "
     Q2 = Q2 & " FROM " & TmpTbl
        
        Set Rs2 = OpenRs(DbMain, Q2)
        
       Do While Rs2.EOF = False
           
            'Q3 = "select Cuentas.Codigo,otroimp from Documento,MovDocumento,TipoValor,Cuentas "
' 3132660
Q3 = "select Cuentas.Codigo, IIF(Movdocumento.Debe > 0, Movdocumento.Debe, Movdocumento.Haber) AS Valor from Documento,MovDocumento,TipoValor,Cuentas "
' fin 3132660
            Q3 = Q3 & "where Documento.IdDoc = MovDocumento.IdDoc and MovDocumento.IdTipoValLib = TipoValor.Codigo "
            Q3 = Q3 & " and MovDocumento.IdCuenta = cuentas.idCuenta " 'and TipoValor.Atributo = 'OTROSIMP' " ' se comenta segun ado 3012320
            Q3 = Q3 & " and Documento.iddoc =" & vFld(Rs2("IDDOC"))
            Q3 = Q3 & " and Documento.IdEmpresa  = " & gEmpresa.id
            Q3 = Q3 & " and Documento.tipolib = TipoValor.tipolib "
            Q3 = Q3 & " and MovDocumento.tasa > 0 "
            Q3 = Q3 & " order by MovDocumento.tasa desc"
           
            Set Rs3 = OpenRs(DbMain, Q3)
            
             Do While Rs3.EOF = False
                If Left(vFld(Rs3("codigo")), 1) = 1 Then ' si es activo (cuenta comienza con 1) no ira en los impuestos que no sean lir
                 Valor = Valor - vFld(Rs3("Valor"))
                End If
                Rs3.MoveNext
             Loop
            
            Call CloseRs(Rs3)
                                  
        Rs2.MoveNext
        
        Loop
   
       Call CloseRs(Rs2)
       'fin 2882638
       
      'Valor = vFld(Rs("Total"))
      
      '2852276
      'If Valor < 0 Then
        'Valor = 0
     ' End If
     'fin 2852276
   End If
   
   'Call CloseRs(Rs)

   Q1 = "DROP TABLE " & TmpTbl
   Rc = ExecSQL(DbMain, Q1)

   
   GetIVAIrrecSaldoPagar = Valor
   
End Function

Private Sub SaveAll()
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   
   lBaseImponible = 0
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Val(Grid.TextMatrix(i, C_NIVEL)) = BIMP14D_MAXNIV Then
      
'         If Grid.TextMatrix(i, C_UPDATE) <> "" Then
            If Grid.TextMatrix(i, C_IDTBLBASEIMP14D) <> "" Then
   
               Q1 = "UPDATE BaseImponible14D SET"
               Q1 = Q1 & " Valor = " & vFmt(Grid.TextMatrix(i, C_VALOR))
               Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND IdBaseImponible14D = " & Grid.TextMatrix(i, C_IDTBLBASEIMP14D)
               Call ExecSQL(DbMain, Q1)

            Else
            
               Q1 = "INSERT INTO BaseImponible14D (IdEmpresa, Ano, Tipo, Nivel, Codigo, Fecha, Valor)"
               Q1 = Q1 & " VALUES(" & gEmpresa.id & "," & gEmpresa.Ano & "," & Grid.TextMatrix(i, C_TIPO)
               Q1 = Q1 & ", " & Grid.TextMatrix(i, C_NIVEL) & ", " & Grid.TextMatrix(i, C_CODIGO) & ", 0, " & vFmt(Grid.TextMatrix(i, C_VALOR)) & ")"
               Call ExecSQL(DbMain, Q1)
   
            End If
            
'         End If
         
      End If
      
   Next i
   
   Q1 = "UPDATE EmpresasAno SET "

   lBaseImponible = vFmt(Grid.TextMatrix(Grid.FixedRows, C_VALOR))

   If gEmpresa.ProPymeGeneral <> 0 Then
      Q1 = Q1 & "  CPS_BaseImpPrimCat_14DN3 = " & lBaseImponible
      Q1 = Q1 & ", CPS_BaseImpPrimCat_14DN8 = 0 "

   ElseIf gEmpresa.ProPymeTransp <> 0 Then
      Q1 = Q1 & "  CPS_BaseImpPrimCat_14DN3 = 0 "
      Q1 = Q1 & ", CPS_BaseImpPrimCat_14DN8 = " & lBaseImponible

   End If

   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
  

End Sub

Private Function valida() As Boolean

   valida = True
   
End Function
Private Sub SetUpGrid()
   Dim i As Integer, WCodCuenta As Integer, WCuenta As Integer
   
   Grid.Cols = NCOLS + 1
      
   Call FGrSetup(Grid, True)

   Grid.ColWidth(C_IDTBLBASEIMP14D) = 0 '500
   Grid.ColWidth(C_IDARRBASEIMP14D) = 0 '500
   Grid.ColWidth(C_REGIMEN) = 0 ' 500
   Grid.ColWidth(C_TIPO) = 0
   Grid.ColWidth(C_NIVEL) = 0 '500
   Grid.ColWidth(C_CODIGO) = 0 '500
   Grid.ColWidth(C_FORMAINGRESO) = 0 '500
   Grid.ColWidth(C_OPENCLOSE) = 300
   Grid.ColWidth(C_DESCRIP) = 7600 + 1800
   Grid.ColWidth(C_VALOR) = 1630
   Grid.ColWidth(C_UPDATE) = 0
   Grid.ColWidth(C_FMT) = 0
   
   Grid.ColAlignment(C_OPENCLOSE) = flexAlignCenterCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_DESCRIP) = "Base Imponible 14D"
   Grid.TextMatrix(0, C_VALOR) = "Monto"
   
      
End Sub
Private Sub OpenCloseGrid()
   Dim Col As Integer
   Dim Row As Integer
   Dim Nivel As Integer
   Dim OpenClose As String
   Dim i As Integer
   
   Col = Grid.Col
   Row = Grid.Row
   Nivel = Val(Grid.TextMatrix(Row, C_NIVEL))
   OpenClose = Trim(Grid.TextMatrix(Row, C_OPENCLOSE))
   
   If Col <> C_OPENCLOSE Or Trim(Grid.TextMatrix(Row, C_OPENCLOSE)) = "" Or Nivel = 1 Then
      Exit Sub
   End If
   
   
   For i = Row + 1 To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_NIVEL)) > Nivel Then
         If OpenClose = "-" Then
            Grid.RowHeight(i) = 0
         Else
            Grid.RowHeight(i) = Grid.RowHeight(0)
            If Val(Grid.TextMatrix(i, C_NIVEL)) <= 4 Then
               Grid.TextMatrix(i, C_OPENCLOSE) = "-"
            End If
         End If
      ElseIf Val(Grid.TextMatrix(i, C_NIVEL)) <= Nivel Then
         Exit For
      End If
   Next i
      
   If OpenClose = "-" Then
      Grid.TextMatrix(Row, C_OPENCLOSE) = "+"
   Else
      Grid.TextMatrix(Row, C_OPENCLOSE) = "-"
   End If
      
End Sub
Private Sub ExpandAll()
   Dim Col As Integer
   Dim Row As Integer
   Dim Nivel As Integer
   Dim OpenClose As String
   Dim i As Integer
   
   Grid.Redraw = False
   
   Col = C_OPENCLOSE
   Row = Grid.FixedRows
   Nivel = 1
   OpenClose = "+"
      
   For i = Row + 1 To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_NIVEL)) > Nivel Then
         If OpenClose = "-" Then
            Grid.RowHeight(i) = 0
         Else
            Grid.RowHeight(i) = Grid.RowHeight(0)
            If Val(Grid.TextMatrix(i, C_NIVEL)) <= 4 Then
               Grid.TextMatrix(i, C_OPENCLOSE) = "-"
            End If
         End If
      End If
   Next i
   
   Call OcultarSegunRegimen
      
   Grid.Redraw = True
   
End Sub
Private Sub SaldosVigentes()
   Dim i As Integer
   
   Call ExpandAll
   
   Grid.Redraw = False
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Val(Grid.TextMatrix(i, C_NIVEL)) > 2 Then
         If vFmt(Grid.TextMatrix(i, C_VALOR)) = 0 Then
            Grid.RowHeight(i) = 0
         Else
            Grid.RowHeight(i) = Grid.RowHeight(0)
         End If
      End If
   Next i
   
   Call OcultarSegunRegimen

   Grid.Redraw = True

End Sub

Private Sub OcultarSegunRegimen()
   Dim i As Integer
      
   For i = Grid.FixedRows To Grid.rows - 1
   
      If (Val(Grid.TextMatrix(i, C_REGIMEN)) = FTE_14DN3 And Not gEmpresa.ProPymeGeneral) Or (Val(Grid.TextMatrix(i, C_REGIMEN)) = FTE_14DN8 And Not gEmpresa.ProPymeTransp) Then
         If vFmt(Grid.TextMatrix(i, C_VALOR)) <> 0 And Val(Grid.TextMatrix(i, C_NIVEL)) = BIMP14D_MAXNIV Then
            ClearDetBaseImp14D (Val(Grid.TextMatrix(i, C_CODIGO)))
            Grid.TextMatrix(i, C_VALOR) = 0
         End If
         Grid.RowHeight(i) = 0
      End If
      
      'condición en DURO dada la especificidad
      If (Val(Grid.TextMatrix(i, C_CODIGO)) = 4200 Or Val(Grid.TextMatrix(i, C_CODIGO)) = 4300) And gEmpresa.Ano = 2021 And Not gEmpresa.ProPymeTransp Then
         If vFmt(Grid.TextMatrix(i, C_VALOR)) <> 0 And Val(Grid.TextMatrix(i, C_NIVEL)) = BIMP14D_MAXNIV Then
            ClearDetBaseImp14D (Val(Grid.TextMatrix(i, C_CODIGO)))
            Grid.TextMatrix(i, C_VALOR) = 0
         End If
         Grid.RowHeight(i) = 0
      End If
      
   Next i
   
   
   Call CalcGrid
   
End Sub
Private Sub ClearDetBaseImp14D(ByVal Codigo As Integer)
   
   Call DeleteSQL(DbMain, "BaseImponible14D", " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND Codigo = " & Codigo)
   
End Sub
Private Sub CalcGrid()
   Dim Col As Integer
   Dim Row As Integer
   Dim Nivel As Integer
   Dim i As Integer
   Dim Tot As Double
     
   For Nivel = 4 To 1 Step -1
      i = Grid.FixedRows
      
      Do While i < Grid.rows - 1
      
         If Val(Grid.TextMatrix(i, C_NIVEL)) = Nivel Then
            
            Row = i
            i = i + 1
            Tot = 0
            Do While i < Grid.rows - 1 And (Grid.TextMatrix(i, C_NIVEL) = "" Or Val(Grid.TextMatrix(i, C_NIVEL)) >= Nivel + 1)
               If Grid.RowHeight(i) > 0 Then
                If Val(Grid.TextMatrix(i, C_NIVEL)) = Nivel + 1 Then
                   Tot = Tot + vFmt(Grid.TextMatrix(i, C_VALOR))
                End If
               End If
               i = i + 1
            Loop
            
            Grid.TextMatrix(Row, C_VALOR) = Format(Tot, NUMFMT)
            
         Else
            i = i + 1
         End If
         
      Loop
      
   Next Nivel
      
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

Private Sub Bt_CopyExcel_Click()
   Call LP_FGr2Clip(Grid, Me.Caption & vbTab & "Año " & gEmpresa.Ano)

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
'   Encabezados(0) = "Al 31 de Diciembre " & gEmpresa.Ano
'   gPrtReportes.Encabezados = Encabezados
         
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = 0
   Next i
                  
   ColWi(C_DESCRIP) = Grid.ColWidth(C_DESCRIP) - 200
   ColWi(C_VALOR) = Grid.ColWidth(C_VALOR) - 100
   
                  
   'Total(C_DESC) = "Capital Pripio Tributario"
   'Total(C_TOTAL) = ""
                  
   gPrtReportes.ColWi = ColWi
   'gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_REGIMEN
   gPrtReportes.FmtCol = C_FMT
   gPrtReportes.NTotLines = 0

End Sub


Private Sub Form_Resize()

   Grid.Height = Me.Height - Grid.Top - 700
   
End Sub

Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   
   Action = vbOK


'   If Col = C_VALOR Then
'      If Value <> "" Then
'         Value = Format(vFmt(Value), NUMFMT)
'         Grid.TextMatrix(Row, Col) = Value
'      End If
'   End If
   
'   If Action = vbOK Then
'      Call FGrModRow(Grid, Row, FGR_U, C_IDSOCIO, C_UPDATE)
'   End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
   
'   If (gEmpresa.ProPymeGeneral <> 0 And Row = FTE_14DN3) Or (gEmpresa.ProPymeTransp <> 0 And Row = FTE_14DN8) Then
'
'      If Col = C_VALOR Then
'         EdType = FEG_Edit
'         Grid.TxBox.MaxLength = 12
'      End If
'
'   End If
   
End Sub

Private Sub Grid_DblClick()
   Dim Col As Integer
   Dim Row As Integer
   Dim Frm As FrmDetBaseImponible14DFull
   Dim Valor As Double
   
   Col = Grid.Col
   Row = Grid.Row
   
   If Col = C_OPENCLOSE Then
      Call OpenCloseGrid
      
   ElseIf Col = C_VALOR Then
   
      If gEmpresa.FCierre <> 0 Then
         Exit Sub
      End If
   
      If Grid.TextMatrix(Row, C_FORMAINGRESO) = ING_MANUAL Or (Grid.TextMatrix(Row, C_FORMAINGRESO) = ING_AMBOS And vFmt(Grid.TextMatrix(Row, C_VALOR)) = 0) Then
         Set Frm = New FrmDetBaseImponible14DFull
         
         If Frm.FEdit(Val(Grid.TextMatrix(Row, C_TIPO)), Val(Grid.TextMatrix(Row, C_CODIGO)), Grid.TextMatrix(Row, C_DESCRIP), Valor) = vbOK Then
            
            If Val(Grid.TextMatrix(Row, C_TIPO)) = BIMP14D_INGRESO Then
               Grid.TextMatrix(Row, C_VALOR) = Format(Valor, NUMFMT)
            Else
               Grid.TextMatrix(Row, C_VALOR) = Format(Valor * -1, NUMFMT)
            End If
            Call FGrModRow(Grid, Row, FGR_U, C_IDTBLBASEIMP14D, C_UPDATE)

            Call CalcGrid
         End If
         
         Set Frm = Nothing
         
         If Val(Grid.TextMatrix(Row, C_CODIGO)) = 8100 Then
            Grid.TextMatrix(lRowCod8300, C_VALOR) = Format(Abs(vFmt(Grid.TextMatrix(Row, C_VALOR))), NUMFMT)
         End If
         
      End If
      
   End If
   
   '2699582
   If gEmpresa.Ano >= 2022 And gEmpresa.ProPymeGeneral = True Or gEmpresa.ProPymeTransp = True Then
   
       If Val(Grid.TextMatrix(Row, C_CODIGO)) = 2800 Then
         Dim FrmPPM As FrmAsisPPM
    
    
         Set FrmPPM = New FrmAsisPPM
         'FrmAsisPPM.LoadAll
         FrmAsisPPM.Show vbModal
    
         Set Frm = Nothing
         Call LoadAll
       End If
       
      
   End If
   'fin 2699582
   
   
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
   
   Call KeyNum(KeyAscii)

End Sub
