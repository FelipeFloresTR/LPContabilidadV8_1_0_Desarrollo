VERSION 5.00
Object = "{D08E2972-AC68-4923-8490-23F41A1304FD}#1.1#0"; "FlexEdGrid3.ocx"
Begin VB.Form FrmPercepciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Percepciones"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   9690
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   25495
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   11055
      Begin VB.Frame Frame 
         Height          =   735
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   9255
         Begin VB.TextBox Tx_IdPerc 
            Height          =   375
            Left            =   3840
            TabIndex        =   27
            Top             =   240
            Visible         =   0   'False
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
            Left            =   2280
            Picture         =   "FrmPercepciones.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Calendario"
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton Bt_OK 
            Caption         =   "Guardar"
            Height          =   315
            Left            =   6840
            TabIndex        =   25
            Top             =   240
            Width           =   885
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
            Left            =   360
            Picture         =   "FrmPercepciones.frx":0429
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Sumar movimientos seleccionados"
            Top             =   180
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
            Left            =   1320
            Picture         =   "FrmPercepciones.frx":04CD
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Imprimir"
            Top             =   180
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
            Left            =   840
            Picture         =   "FrmPercepciones.frx":0987
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Vista previa de la impresión"
            Top             =   180
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
            Left            =   1800
            Picture         =   "FrmPercepciones.frx":0E2E
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Copiar Excel"
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton Bt_Salir 
            Caption         =   "Cerrar"
            Height          =   315
            Left            =   8160
            TabIndex        =   20
            Top             =   240
            Width           =   885
         End
         Begin VB.Frame Fr_Doc 
            BorderStyle     =   0  'None
            Height          =   470
            Left            =   3660
            TabIndex        =   19
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Detalle de las Cantidades Percibidas"
         Height          =   2295
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   9135
         Begin VB.TextBox Tx_TEX 
            Height          =   315
            Left            =   6720
            MaxLength       =   9
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1680
            Width           =   2115
         End
         Begin VB.TextBox Tx_TEF 
            Height          =   315
            Left            =   2400
            MaxLength       =   9
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1680
            Width           =   2115
         End
         Begin VB.ComboBox Cb_Contabiliza 
            Height          =   315
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1080
            Width           =   2175
         End
         Begin VB.ComboBox Cb_RegEmpresa 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox Tx_Rut 
            Height          =   315
            Left            =   7080
            MaxLength       =   12
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   360
            Width           =   1755
         End
         Begin VB.TextBox Txt_NCertificado 
            Height          =   285
            Left            =   3720
            TabIndex        =   7
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Tx_Fecha 
            Height          =   315
            Left            =   780
            TabIndex        =   4
            Top             =   360
            Width           =   1155
         End
         Begin VB.CommandButton Bt_SelFecha 
            Height          =   315
            Left            =   1920
            Picture         =   "FrmPercepciones.frx":1273
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tasa TEX :"
            Height          =   195
            Left            =   4920
            TabIndex        =   17
            Top             =   1680
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Efectiva de Crédito - TEF:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   1740
            Width           =   2235
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contabilización:"
            Height          =   195
            Index           =   3
            Left            =   4920
            TabIndex        =   13
            Top             =   1140
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Regimen Empresa Fuente:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   1140
            Width           =   1875
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "RUT Empresa Fuente:"
            Height          =   195
            Left            =   5400
            TabIndex        =   9
            Top             =   420
            Width           =   1590
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N° Certificado:"
            Height          =   195
            Index           =   0
            Left            =   2640
            TabIndex        =   6
            Top             =   420
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   420
            Width           =   495
         End
      End
      Begin FlexEdGrid3.FEd3Grid Grid 
         Height          =   6960
         Left            =   240
         TabIndex        =   1
         Top             =   3360
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   12277
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
   End
End
Attribute VB_Name = "FrmPercepciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const C_CODIGO = 0
Const C_DESCRIP = 1
Const C_VALOR = 2
Const C_NIVEL = 3
Const NCOLS = C_NIVEL

Public CodCta As String
Public GIdPerc As Long
Public Fecha As String
Public orden As Long
Public idcomp As Long

Private Sub Bt_Calendar_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
   
   Call Frm.SelDate(Fecha)
   
   Set Frm = Nothing
End Sub

Private Sub Bt_CopyExcel_Click()
'Call LP_FGr2Clip(Grid, Me.Caption & vbTab & "Año " & gEmpresa.Ano)
Call LP_FGr2Clip(Grid, "  Descripción  " & vbTab & "  Valor ")
End Sub

Private Sub Bt_OK_Click()
Call CalcGrid
If Valida() Then
    Call SaveAll
End If

End Sub

Private Function Valida() As Boolean

    Valida = False
    
    If Trim(Tx_Fecha.Text) = "" Then
      MsgBox "Favor ingresar una Fecha", vbExclamation, "Percepciones"
      Tx_Fecha.SetFocus
      Exit Function
    End If
    
    If Not validaFecha() Then
        Tx_Fecha.Text = ""
        Exit Function
   End If
    
    If Trim(Txt_NCertificado.Text) = "" Or Val(Trim(Txt_NCertificado.Text)) < 1 Then
      MsgBox "El Numero de certificado es obligatorio y mayor a 0", vbExclamation, "Percepciones"
      Txt_NCertificado.SetFocus
      Exit Function
    End If
    
    If Trim(Tx_Rut.Text) = "" Then
      MsgBox "Favor ingresar un Rut Empresa Fuente", vbExclamation, "Percepciones"
      Tx_Rut.SetFocus
      Exit Function
    End If

   If Not ValidRut(Me.Tx_Rut.Text) Then
      MsgBox "Rut No válido, Favor volver a ingresar", vbExclamation, "Percepciones"
      Me.Tx_Rut.Text = ""
      Exit Function
   End If
   
   If vFmtRut(Tx_Rut) = gEmpresa.Rut Then
      MsgBox "El Rut Empresa Fuente No puede ser igual al Rut de la empresa con la que esta trabajando", vbExclamation, "Percepciones"
      Me.Tx_Rut.Text = ""
      Exit Function
   End If

    If CbItemData(Cb_RegEmpresa) < 0 Then
      MsgBox "Favor ingresar un Regimen Empresa Fuente", vbExclamation, "Percepciones"
      Cb_RegEmpresa.SetFocus
      Exit Function
    End If

    If CbItemData(Cb_Contabiliza) < 0 Then
      MsgBox "Favor ingresar una Contabilización", vbExclamation, "Percepciones"
      Cb_Contabiliza.SetFocus
      Exit Function
    End If
    
    If Trim(Tx_TEF.Text) <> "" Then
        If CDbl(Me.Tx_TEF) > 100 Then
         MsgBox "El Valor TEF No puede exceder el 100%", vbExclamation, "Percepciones"
         Tx_TEF.SetFocus
         Exit Function
        End If
    End If
    
    If Trim(Tx_TEX.Text) <> "" Then
        If CDbl(Me.Tx_TEX) > 100 Then
         MsgBox "El Valor TEX No puede exceder el 100%", vbExclamation, "Percepciones"
         Tx_TEX.SetFocus
         Exit Function
        End If
    End If
    
    If CDbl(Grid.TextMatrix(1, C_VALOR)) = 0 Then
        MsgBox "El Valor de las Percepciones debe ser mayor a 0", vbExclamation, "Percepciones"
        Exit Function
    End If

    Valida = True

End Function

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
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(2) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = ORIENT_VER
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Me.Caption
   Titulos(1) = "N° Certificado: " & Txt_NCertificado
   Titulos(2) = "RUT Empresa Fuente: " & Tx_Rut
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
   'gPrtReportes.ColObligatoria = C_REGIMEN
   gPrtReportes.FmtCol = C_FMT
   gPrtReportes.NTotLines = 0

End Sub

Public Sub ResetPrtBas(PrtCls As ClsPrtFlxGrid)
   Dim ColWi(0) As Integer
   Dim Total(0) As String
   Dim Titulos(0) As String
   Dim FntTitulos(0) As FontDef_t
   Dim FntEncabezados(0) As FontDef_t
   Dim Encabezados(0) As String
   Dim EncabezadosCont(0) As String

   PrtCls.CallEndDoc = True
   PrtCls.ColObligatoria = 1
   PrtCls.PrintHeader = True
   PrtCls.EsContinuacion = False
   
   PrtCls.GrFontName = ""
   PrtCls.GrFontSize = -1
   PrtCls.TotFntBold = True
   
   PrtCls.InitPag = -1
   PrtCls.CellHeight = 0
   
   PrtCls.FmtCol = -1
   PrtCls.PrintFecha = ChkNoPrtFecha()
   
   PrtCls.ColWi = ColWi
   PrtCls.Titulos = Titulos
   PrtCls.Encabezados = Encabezados
   PrtCls.EncabezadosCont = EncabezadosCont
   Call PrtCls.FntTitulos(FntTitulos)
   Call PrtCls.FntEncabezados(FntEncabezados)

End Sub

Private Sub Bt_Salir_Click()
Unload Me
End Sub

Private Sub Bt_SelFecha_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   
   Set Frm = New FrmCalendar
  
   Call Frm.TxSelDate(Tx_Fecha)
   
   Set Frm = Nothing
   
   If Not validaFecha() Then
        Tx_Fecha.Text = ""
   End If
   
End Sub

Private Function validaFecha() As Boolean
   Dim desde As Long
   Dim Hasta As Long
   Dim fechaSel As Long
   
   validaFecha = True
   desde = DateSerial(gEmpresa.Ano, 1, 1) 'DateSerial((gEmpresa.Ano - 1), 31, 12)
   Hasta = DateSerial(gEmpresa.Ano + 1, 1, 1)
   fechaSel = GetTxDate(Tx_Fecha)
   
   If GetTxDate(Tx_Fecha) < desde Or GetTxDate(Tx_Fecha) > Hasta Then
        MsgBox "La Fecha tiene que estar dentro de el año que esta trabajando"
        validaFecha = False
   End If


End Function

Private Sub Bt_Sum_Click()
Dim Frm As FrmSumSimple
   
   Set Frm = New FrmSumSimple
   
   Call Frm.FViewSum(Grid)
   
   Set Frm = Nothing
End Sub

Private Sub Form_Load()

If Fecha <> "" Then
    Me.Tx_Fecha = Format(Fecha, DATEFMT)
End If
Call SetupRegimenEmpFuente
Call SetupContabilizacion
Call FillParam(Cb_RegEmpresa, "REGEMPREFUE")
Call FillParam(Cb_Contabiliza, "CONTABILIZA")
Call SetUpGrid
Call LoadBase
Tx_IdPerc = 0
If GIdPerc > 0 Then
    Me.Tx_IdPerc.Text = GIdPerc
    Call CargarValores
    Call Grid_SelChange
Else
   Call CargarValores
    Call Grid_SelChange
End If
'Call CreateTblPercepciones
'
'Call CreateTblDetPercepciones


End Sub


Private Sub Grid_BeforeEdit(ByVal row As Integer, ByVal Col As Integer, EdType As FlexEdGrid3.FEG3_EdType)
If Col = C_VALOR Then
    If Grid.TextMatrix(row, C_NIVEL) > 4 Then
    EdType = FEG_Edit
    End If
End If
End Sub

Private Sub Grid_EditKeyPress(KeyAscii As Integer)
Call KeyNumPos(KeyAscii)
End Sub

Private Sub Grid_SelChange()
Call CalcGrid
End Sub

Private Sub Tx_Fecha_GotFocus()
Call DtGotFocus(Tx_Fecha)
End Sub

Private Sub Tx_Fecha_LostFocus()
Call DtLostFocus(Tx_Fecha)
End Sub


Private Sub Tx_RUT_Validate(Cancel As Boolean)
If Tx_Rut = "" Then
      Exit Sub
   End If
   
   If Not ValidRut(Me.Tx_Rut.Text) Then
      MsgBox1 "Rut No Válido, Favor volver a ingresar", vbInformation
      Me.Tx_Rut.Text = ""
      Exit Sub
   End If
   
   If vFmtRut(Tx_Rut) = gEmpresa.Rut Then
      MsgBox1 "El Rut Empresa Fuente No puede ser igual al Rut de la empresa con la que esta trabajando", vbInformation
      Me.Tx_Rut.Text = ""
      Exit Sub
   End If
   
End Sub

Private Sub Tx_TEF_KeyPress(KeyAscii As Integer)
Call KeyNumDecimal(KeyAscii)
End Sub


Private Sub Tx_TEX_KeyPress(KeyAscii As Integer)
Call KeyNumDecimal(KeyAscii)
End Sub

Private Sub Txt_NCertificado_KeyPress(KeyAscii As Integer)
Call KeyNumPos(KeyAscii)
End Sub

Private Sub SetUpGrid()
   Dim i As Integer, WCodCuenta As Integer, WCuenta As Integer
   
   Grid.Cols = NCOLS + 1
      
   Call FGrSetup(Grid, True)

   Grid.ColWidth(C_CODIGO) = 0
'   Grid.ColWidth(C_IDARRBASEIMP14D) = 0 '500
'   Grid.ColWidth(C_REGIMEN) = 0 ' 500
'   Grid.ColWidth(C_TIPO) = 0
   Grid.ColWidth(C_NIVEL) = 0
'   Grid.ColWidth(C_CODIGO) = 0 '500
'   Grid.ColWidth(C_FORMAINGRESO) = 0 '500
'   Grid.ColWidth(C_OPENCLOSE) = 300
   Grid.ColWidth(C_DESCRIP) = 5600
   Grid.ColWidth(C_VALOR) = 1630
   
   
'   Grid.ColWidth(C_UPDATE) = 0
'   Grid.ColWidth(C_FMT) = 0
'
'   Grid.ColAlignment(C_OPENCLOSE) = flexAlignCenterCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
'
'   Grid.TextMatrix(0, C_DESCRIP) = "Base Imponible 14D"
'   Grid.TextMatrix(0, C_VALOR) = "Monto"
   
      
End Sub


Private Sub LoadBase()
   Dim i As Integer
   Dim row As Integer

   'Grid.Redraw = False
   
   row = Grid.FixedRows
   Grid.rows = Grid.FixedRows
   
   
   For i = 1 To UBound(Percepciones)
      
      If Percepciones(i).Nivel = 0 Then
         Exit For
      End If
      

      Grid.rows = Grid.rows + 1
      
      If (Percepciones(i).Codigo = 300 Or Percepciones(i).Codigo = 1800 Or Percepciones(i).Codigo = 3100 Or Percepciones(i).Codigo = 4400) Then
       Grid.RowHeight(row) = 0
      End If
'
'      If gBaseImponible14D(i).Nivel <= 2 And Row > Grid.FixedRows Then
'         Grid.rows = Grid.rows + 1
'         Row = Row + 1
'      End If
'
'      Grid.TextMatrix(Row, C_IDARRBASEIMP14D) = i
'      Grid.TextMatrix(Row, C_REGIMEN) = gBaseImponible14D(i).Regimen
'      Grid.TextMatrix(Row, C_TIPO) = gBaseImponible14D(i).Tipo
      Grid.TextMatrix(row, C_NIVEL) = Percepciones(i).Nivel
'      Grid.TextMatrix(Row, C_FORMAINGRESO) = gBaseImponible14D(i).FormaIngreso
      Grid.TextMatrix(row, C_CODIGO) = Percepciones(i).Codigo
'
'      If gBaseImponible14D(i).Nivel <= 4 And gBaseImponible14D(i).Nivel > 1 Then
'         Grid.TextMatrix(Row, C_OPENCLOSE) = "-"
'         Call FGrFontBold(Grid, Row, C_OPENCLOSE, True)
'      End If
      Grid.TextMatrix(row, C_DESCRIP) = String((Percepciones(i).Nivel - 1) * 4, " ") & Percepciones(i).Nombre
'
      If Percepciones(i).Nivel <= 2 Then
         Grid.TextMatrix(row, C_DESCRIP) = UCase(Grid.TextMatrix(row, C_DESCRIP))
      End If
'
      If Percepciones(i).Nivel <= 3 Then
         Call FGrFontBold(Grid, row, -1, True)
         'Grid.TextMatrix(Row, C_FMT) = "B"
      Else
         Grid.TextMatrix(row, C_DESCRIP) = String((Percepciones(i).Nivel - 1) * 2, " ") & Grid.TextMatrix(row, C_DESCRIP)
      End If
      
      If Percepciones(i).Nivel = 4 Then
         Call FGrForeColor(Grid, row, -1, vbBlue)
         'Grid.TextMatrix(Row, C_FMT) = "FCELL"
      End If
      
      If Percepciones(i).Nivel = 5 And Percepciones(i).FormaIngreso <> ING_MANUAL Then
         Call FGrBackColor(Grid, row, C_VALOR, COLOR_GRISLTLT)
      End If
      
      Grid.TextMatrix(row, C_VALOR) = 0
      
      row = row + 1
NextRow:

   Next i

'   Call OcultarSegunRegimen
   
   Grid.rows = Grid.rows + 1
   
   'Grid.Redraw = True
   
End Sub

Private Sub CargarValores()
Dim Q1 As String
Dim Rs As Recordset

   Q1 = "SELECT Fecha, NumCertificado, RutEmpresa, Regimen, Contabilizacion, TasaTef, TasaTex, Coddet, valor, P.IDPERC "
   Q1 = Q1 & " FROM DETPERCEPCIONES D, PERCEPCIONES P "
   Q1 = Q1 & " WHERE D.IDPerc = P.IDPerc "
   If GIdPerc > 0 Then
        Q1 = Q1 & " AND D.IDPerc = " & GIdPerc
   Else
   
     If idcomp > 0 And orden > 0 And CodCta > 0 Then
        Q1 = Q1 & " AND P.IDCOMP = " & idcomp
        Q1 = Q1 & " AND P.ORDEN = " & orden
        Q1 = Q1 & " AND P.IDCUENTA = " & CodCta
    Else
        Exit Sub
    End If
     
   End If
   
   Q1 = Q1 & " ORDER BY Coddet "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
        Me.Tx_Fecha = Format(vFld(Rs("Fecha")), DATEFMT)
        Me.Txt_NCertificado = vFld(Rs("NumCertificado"))
        Me.Tx_Rut = FmtCID(vFld(Rs("RutEmpresa")))
        Call CbSelItem(Me.Cb_RegEmpresa, vFld(Rs("Regimen")))
        Call CbSelItem(Me.Cb_Contabiliza, vFld(Rs("Contabilizacion")))
        Me.Tx_TEF = vFld(Rs("TasaTef"))
        Me.Tx_TEX = vFld(Rs("TasaTex"))
        GIdPerc = vFld(Rs("IDPERC"))
        Tx_IdPerc = vFld(Rs("IDPERC"))
       Do While Rs.EOF = False
       
          For i = Grid.FixedRows To Grid.rows - 1
          
            If vFmt(Grid.TextMatrix(i, C_CODIGO)) = vFld(Rs("Coddet")) Then
              Grid.TextMatrix(i, C_VALOR) = vFld(Rs("valor"))
            End If
    
          Next i
    
    
          Rs.MoveNext
    
       Loop
   
   End If


End Sub
Private Sub CalcGrid()
   Dim Col As Integer
   Dim row As Integer
   Dim Nivel As Integer
   Dim i As Integer
   Dim Tot As Double
     
   For Nivel = 4 To 1 Step -1
      i = Grid.FixedRows
      
      Do While i < Grid.rows - 1
      
         If Val(Grid.TextMatrix(i, C_NIVEL)) = Nivel Then
            
            row = i
            i = i + 1
            Tot = 0
            Do While i < Grid.rows - 1 And (Grid.TextMatrix(i, C_NIVEL) = "" Or Val(Grid.TextMatrix(i, C_NIVEL)) >= Nivel + 1)
               If Val(Grid.TextMatrix(i, C_NIVEL)) = Nivel + 1 Then
                  Tot = Tot + vFmt(Grid.TextMatrix(i, C_VALOR))
               End If
               i = i + 1
            Loop
            
            Grid.TextMatrix(row, C_VALOR) = Format(Tot, NUMFMT)
            
         Else
            i = i + 1
         End If
         
      Loop
      
   Next Nivel
      
End Sub

Private Sub SaveAll()
   Dim i As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tef, Tex As String
   
   
   lBaseImponible = 0
   
           If gDbType = SQL_ACCESS Then
            Q1 = "SELECT IIF(COUNT(EXISTS(SELECT IDPERC "
            Q1 = Q1 & "                   FROM PERCEPCIONES WHERE IDPERC = " & Me.Tx_IdPerc & ")) = 0, 0, MAX(IDPERC)) AS Id "
        
            Tef = Tx_TEF
            Tex = Tx_TEX
        Else
            Q1 = "SELECT ISNULL(MAX(IDPERC),0) AS ID   "
            
            Tef = Replace(Tx_TEF, ",", ".")
            Tex = Replace(Tx_TEX, ",", ".")
        
        End If
        Q1 = Q1 & " From Percepciones "
   
    If Tx_IdPerc = 0 Then
    
        Set Rs = OpenRs(DbMain, Q1)
    
        If Not Rs.EOF Then
           Tx_IdPerc = vFld(Rs("Id")) + 1
        End If
        Call CloseRs(Rs)
        
            Q1 = "INSERT INTO Percepciones (IDPerc, IdCuenta, IdEmpresa, Ano, Fecha, NumCertificado, RutEmpresa, Regimen, Contabilizacion, TasaTef, TasaTex, Percepciones, orden )"
            Q1 = Q1 & " VALUES(" & Tx_IdPerc & "," & CodCta & "," & gEmpresa.id & ", " & gEmpresa.Ano & ", " & GetTxDate(Tx_Fecha) & ", " & Trim(Txt_NCertificado) & ", " & vFmtRut(Tx_Rut)
            Q1 = Q1 & ", " & IIf(CbItemData(Cb_RegEmpresa) < 0, "", CbItemData(Cb_RegEmpresa)) & ", " & IIf(CbItemData(Cb_Contabiliza) < 0, "", CbItemData(Cb_Contabiliza))
            Q1 = Q1 & ", '" & Tef & "', '" & Tex & "', '" & CDbl(Grid.TextMatrix(1, C_VALOR)) & "', " & orden & ")"
            Call ExecSQL(DbMain, Q1)
        
        
           For i = Grid.FixedRows To Grid.rows - 1
           
               Q1 = "INSERT INTO DetPercepciones (IDPerc, CodDet, Valor )"
               Q1 = Q1 & " VALUES(" & Tx_IdPerc & "," & vFmt(Grid.TextMatrix(i, C_CODIGO)) & ",'" & Trim(vFmt(Grid.TextMatrix(i, C_VALOR))) & "')"
               Call ExecSQL(DbMain, Q1)
      
            Next i
            
            MsgBox "Informacion guardada correctamente", vbInformation, "Percepciones"
            Unload Me
            
    Else
    
        Q1 = "UPDATE Percepciones "
        Q1 = Q1 & " SET Percepciones.IdCuenta = " & CodCta & ", "
        Q1 = Q1 & "        Percepciones.IdEmpresa = " & gEmpresa.id & ", "
        Q1 = Q1 & "        Percepciones.Ano = " & gEmpresa.Ano & ", "
        Q1 = Q1 & "        Percepciones.Fecha = " & GetTxDate(Tx_Fecha) & ", "
        Q1 = Q1 & "        Percepciones.NumCertificado = " & Trim(Txt_NCertificado) & ", "
        Q1 = Q1 & "        Percepciones.RutEmpresa = " & vFmtRut(Tx_Rut) & ", "
        Q1 = Q1 & "        Percepciones.Regimen = " & IIf(CbItemData(Cb_RegEmpresa) < 0, "", CbItemData(Cb_RegEmpresa)) & ", "
        Q1 = Q1 & "        Percepciones.Contabilizacion = " & IIf(CbItemData(Cb_Contabiliza) < 0, "", CbItemData(Cb_Contabiliza)) & ", "
        Q1 = Q1 & "        Percepciones.TasaTef = '" & Tef & "', "
        Q1 = Q1 & "        Percepciones.TasaTex = '" & Tex & "', "
        Q1 = Q1 & "        Percepciones.Percepciones = '" & CDbl(Grid.TextMatrix(1, C_VALOR)) & "' "
        Q1 = Q1 & " Where Percepciones.IDPERC = " & Tx_IdPerc
        Call ExecSQL(DbMain, Q1)
        
        For i = Grid.FixedRows To Grid.rows - 1
        
            Q1 = "UPDATE DetPercepciones "
            Q1 = Q1 & " SET DetPercepciones.Valor = '" & vFmt(Grid.TextMatrix(i, C_VALOR)) & "' "
            Q1 = Q1 & " WHERE DetPercepciones.IDPerc = " & Tx_IdPerc
            Q1 = Q1 & " AND DetPercepciones.CodDet =  " & vFmt(Grid.TextMatrix(i, C_CODIGO))
            Call ExecSQL(DbMain, Q1)
        
        Next i
        
        MsgBox "Informacion guardada correctamente", vbInformation, "Percepciones"
        Unload Me
        
   End If


End Sub

