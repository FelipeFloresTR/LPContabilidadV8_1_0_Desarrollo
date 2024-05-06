VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmAjustesExtraLibCajaRLI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajustes Extra - Contables RLI HR RAB"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   11685
   StartUpPosition =   1  'CenterOwner
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   14420
      Cols            =   2
      Rows            =   2
      FixedCols       =   1
      FixedRows       =   0
      ScrollBars      =   2
      AllowUserResizing=   0
      HighLight       =   1
      SelectionMode   =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   -1  'True
      Locked          =   0   'False
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton Bt_VerSaldosPositivos 
         Caption         =   "Mostrar Partidas con Valores"
         Height          =   315
         Left            =   4500
         TabIndex        =   8
         Top             =   180
         Width           =   2355
      End
      Begin VB.CommandButton Bt_ExportHRRAB 
         Caption         =   "Exportar a HR RAB"
         Height          =   315
         Left            =   8520
         TabIndex        =   9
         Top             =   180
         Width           =   1635
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
         Left            =   1440
         Picture         =   "FrmAjustesExtraLibCajaRLI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   2400
         Picture         =   "FrmAjustesExtraLibCajaRLI.frx":00A4
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   1980
         Picture         =   "FrmAjustesExtraLibCajaRLI.frx":0405
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Convertir moneda"
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
         Left            =   2820
         Picture         =   "FrmAjustesExtraLibCajaRLI.frx":07A3
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   900
         Picture         =   "FrmAjustesExtraLibCajaRLI.frx":0BCC
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   60
         Picture         =   "FrmAjustesExtraLibCajaRLI.frx":1011
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cancelar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10200
         TabIndex        =   10
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
         Left            =   480
         Picture         =   "FrmAjustesExtraLibCajaRLI.frx":14B8
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmAjustesExtraLibCajaRLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_ID = 0
Const C_TIPOAJUSTE = 1
Const C_IDGRUPO = 2
Const C_IDITEM = 3
Const C_TIPOITEM = 4
Const C_CONCEPTO = 5
Const C_VALOR = 6
Const C_FMT = 7
Const C_COLOBLIGATORIA = 8
Const C_UPD = 9

Const NCOLS = C_UPD

Dim lSepGrid As String

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub


Private Sub Bt_DetDoc_Click()
   Dim Row As Integer
   Dim Frm As Form
   
   Row = Grid.Row
   
   
End Sub

Public Sub FEdit()
   Me.Show vbModal
End Sub

Private Sub Bt_ExportHRRAB_Click()
   Dim fname As String
   
   Me.MousePointer = vbHourglass
   Call Export_RLI_HR_RAB(fname)
   Me.MousePointer = vbDefault

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Bt_VerSaldosPositivos_Click()
   Dim i As Integer
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.TextMatrix(i, C_TIPOITEM) <> "" And Val(Grid.TextMatrix(i, C_VALOR)) = 0 Then
         Grid.RowHeight(i) = 0
      End If
   Next i
   
   Call FGrVRows(Grid, 1)
End Sub

Private Sub Form_Load()
   lOrientacion = ORIENT_VER

   Call SetUpGrid
   
   If gEmpresa.Ano >= 2020 Then
      Me.Caption = ReplaceStr(Me.Caption, "RAB", "RAD")
      Bt_ExportHRRAB.Caption = "Exportar a HR RAD"
   End If

   Call LoadAll
End Sub

Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1
   Call FGrSetup(Grid)
      
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_TIPOAJUSTE) = 0
   Grid.ColWidth(C_IDGRUPO) = 0
   Grid.ColWidth(C_IDITEM) = 0
   Grid.ColWidth(C_TIPOITEM) = 0    '400
   Grid.ColWidth(C_CONCEPTO) = 9400
   Grid.ColWidth(C_VALOR) = 1600
   Grid.ColWidth(C_FMT) = 0
   Grid.ColWidth(C_COLOBLIGATORIA) = 0
   Grid.ColWidth(C_UPD) = 0
   
   Grid.ColAlignment(C_TIPOITEM) = flexAlignCenterCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   
   Grid.FlxGrid.BackColor = vbButtonFace
   
   Call FGrVRows(Grid)

End Sub
Private Sub LoadAll()
   Dim i As Integer, Row As Integer
   Dim TipoIngreso As Integer
   Dim Valor As Double
   Dim FmtLine As String
   Dim id As Long
   Dim j As Integer, k As Integer, o As Integer
   
   Grid.FlxGrid.Redraw = False
   Grid.rows = 0
   Row = -1
      
   FmtLine = "L"
   lSepGrid = " | "
    
   For k = 1 To MAX_TIPOAJUSTESECRLI
   
      If gTipoAjustesECRLI(k) = "" Then
         Exit For
      End If
  
      Row = Row + 1
      Grid.rows = Row + 1
      If k > 1 Then
         Row = Row + 1
         Grid.rows = Row + 1
         Grid.TextMatrix(Row - 1, C_COLOBLIGATORIA) = "."

      End If
      
      Grid.TextMatrix(Row, C_CONCEPTO) = UCase(gTipoAjustesECRLI(k))
      Grid.TextMatrix(Row, C_FMT) = "B"
      Call FGrSetRowStyle(Grid, Row, "B")
      Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
      Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
  
   
      For j = 1 To MAX_GRUPOAJUSTESECRLI
      
         If gGrupoAjustesECRLI(k, j) = "" Then
            Exit For
         End If
         
         Row = Row + 1
         Grid.rows = Row + 1
         Grid.TextMatrix(Row, C_CONCEPTO) = String(3, " ") & gGrupoAjustesECRLI(k, j)
         Grid.TextMatrix(Row, C_FMT) = "B"
         Call FGrSetRowStyle(Grid, Row, "B")
         Call FGrSetRowStyle(Grid, Row, "BC", COLOR_GRISLT)
         Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
   
         For o = 1 To MAX_ITEMAJUSTESECRLI
            For i = 1 To MAX_ITEMAJUSTESECRLI
            
               If gAjustesExtraContRLI(k, j, i).Nombre <> "" And gAjustesExtraContRLI(k, j, i).orden = o Then
               
                  Row = Row + 1
                  Grid.rows = Row + 1
                  Grid.TextMatrix(Row, C_TIPOAJUSTE) = k
                  Grid.TextMatrix(Row, C_IDGRUPO) = j
                  Grid.TextMatrix(Row, C_IDITEM) = i
                  Grid.TextMatrix(Row, C_TIPOITEM) = gAjustesExtraContRLI(k, j, i).TipoItem
                  Grid.TextMatrix(Row, C_CONCEPTO) = String(10, " ") & gAjustesExtraContRLI(k, j, i).Nombre
                              
                  Valor = LoadValCuentas(Row, Grid.TextMatrix(Row, C_TIPOITEM), Grid.TextMatrix(Row, C_CONCEPTO), gAjustesExtraContRLI(k, j, i).LstCuentas)
                  Grid.TextMatrix(Row, C_VALOR) = Format(Abs(Valor), NEGNUMFMT)
                              
                  Grid.TextMatrix(Row, C_FMT) = FmtLine
                  FmtLine = ""
                  Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
               
               End If
            Next i
         Next o
      Next j
      
   Next k
         
   
   Row = Row + 1
   Grid.rows = Row + 1
   Grid.TextMatrix(Row, C_FMT) = "L"
   Grid.TextMatrix(Row, C_COLOBLIGATORIA) = "."
         
   Grid.TopRow = 0
   
   Grid.FlxGrid.Redraw = True

End Sub
Private Function LoadValCuentas(ByVal Row As Integer, ByVal TipoItem As String, ByVal NombreItem As String, ByVal LstCuentas As String) As Double
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Tot As Double
   Dim Descrip As String
      
   If LstCuentas = "" Then
      LoadValCuentas = 0
      Exit Function
   End If
   
   LstCuentas = Mid(LstCuentas, 2)
   LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
   
   Q1 = "SELECT Sum(Debe - Haber) as Valor "
   Q1 = Q1 & " FROM MovComprobante INNER JOIN Comprobante ON (MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " WHERE IdCuenta IN (" & LstCuentas & ")"
   Q1 = Q1 & " AND TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.Tipo <> " & TC_APERTURA            'se agrega esta condición por solicitud de Joshua Nicolás Catrin (07/09/2018)
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      Tot = Format(vFld(Rs("Valor")), NEGNUMFMT)
   End If
   
   Call CloseRs(Rs)
   
   LoadValCuentas = Tot
   
End Function
Private Sub Form_Resize()
   
   Grid.Height = Me.Height - Grid.Top - 800

   If Grid.Height > 10930 Then
      Grid.Height = 10930
   End If
   
   Call FGrVRows(Grid)
   
End Sub
Private Sub Bt_Print_Click()
   Dim OldOrientacion As Integer
   Dim Frm As FrmPrtSetup
   Dim nFolio As Integer
   Dim Pag As Integer
   
   OldOrientacion = Printer.Orientation
   
   Me.MousePointer = vbHourglass
   
   Call SetUpPrtGrid
         
   Call gPrtLibros.PrtFlexGrid(Printer)
         
   Me.MousePointer = vbDefault
         
   Printer.Orientation = OldOrientacion
   lInfoPreliminar = False
            
   Call ResetPrtBas(gPrtLibros)
      
End Sub

Private Sub Bt_Preview_Click()
   Dim Frm As FrmPrintPreview
   Dim PrtOrient As Integer
   Dim Pag As Integer
      
   PrtOrient = Printer.Orientation
   
   lPapelFoliado = False
   Call SetUpPrtGrid
   
   Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   Me.MousePointer = vbHourglass
      
   Call gPrtLibros.PrtFlexGrid(Frm)
   
   Set Frm.PrtControl = Bt_Print
   
   Me.MousePointer = vbDefault
      
   Call Frm.FView(Caption)
   Set Frm = Nothing
      
   Call ResetPrtBas(gPrtLibros)
   Printer.Orientation = PrtOrient
   gPrtLibros.FmtCol = -1
           
End Sub
Private Sub SetUpPrtGrid()
   Dim i As Integer, j As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Titulos(1) As String
   
   Set gPrtLibros.Grid = Grid
   
   lOrientacion = ORIENT_VER
   Titulos(0) = Me.Caption
                        
   gPrtLibros.Titulos = Titulos
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i) * 0.9
   Next i
   
   gPrtLibros.ColWi = ColWi
   gPrtLibros.ColObligatoria = C_COLOBLIGATORIA
   gPrtLibros.FmtCol = C_FMT
   gPrtLibros.NTotLines = 0
            
End Sub

Private Sub Bt_CopyExcel_Click()
   Dim Clip As String

   Clip = LP_FGr2String(Grid, Me.Caption)
   Clipboard.Clear
   Clipboard.SetText Clip
      
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
Public Function Export_RLI_HR_RAB(fname As String) As Long
   Dim FPath As String
   Dim LogPath As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String
   Dim i As Integer, j As Integer, k As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Valor As Double
   Dim n As Long, r As Integer
   Dim ExpDir As String
   Dim TipoArchivo As String
   Dim LstCuentas As String
   Dim TipoItem As String
   Dim Descrip As String
   
   On Error Resume Next
      
   Sep = ";"
   
   ExpDir = gHRPath & "\RUTS"
   MkDir ExpDir
      
   ExpDir = ExpDir & "\" & Right("00000000" & gEmpresa.Rut, 8)
   MkDir ExpDir
      
   ExpDir = ExpDir & "\ImpConta"
   MkDir ExpDir

   If gEmpresa.Ano < 2020 Then
      fname = "RLI_HR_RAB"
   Else
      fname = "RLI_HR_RAD"
   End If
      
   fname = fname & "_" & Right(gEmpresa.Ano, 2) & ".csv"
   
   FPath = ExpDir & "\" & fname
   
   Fd = FreeFile
   ERR.Clear
   
   Open FPath For Output As #Fd
   If ERR Then
      MsgErr FPath
      Export_RLI_HR_RAB = -ERR
      Exit Function
   End If

   On Error GoTo 0
   
   Buf = "Tipo Item" & Sep & "Fecha" & Sep & "Descripción" & Sep & "Monto"

   Print #Fd, Buf

   Buf = ""
   n = 0
   
   
   'imprimimos el archivo
   
   For r = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(r, C_TIPOITEM) <> "" And vFmt(Grid.TextMatrix(r, C_VALOR)) <> 0 Then
      
         k = Val(Grid.TextMatrix(r, C_TIPOAJUSTE))
         j = Val(Grid.TextMatrix(r, C_IDGRUPO))
         i = Val(Grid.TextMatrix(r, C_IDITEM))
         
         LstCuentas = gAjustesExtraContRLI(k, j, i).LstCuentas
         LstCuentas = Mid(LstCuentas, 2)
         LstCuentas = Left(LstCuentas, Len(LstCuentas) - 1)
         TipoItem = Grid.TextMatrix(r, C_TIPOITEM)
      
         Q1 = "SELECT Comprobante.Tipo, Comprobante.Correlativo, Comprobante.Fecha, MovComprobante.Debe - MovComprobante.Haber as Valor "
         Q1 = Q1 & " FROM MovComprobante INNER JOIN Comprobante ON (MovComprobante.IdComp = Comprobante.IdComp "
         Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
         Q1 = Q1 & " WHERE IdCuenta IN (" & LstCuentas & ")"
         Q1 = Q1 & " AND TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
         Q1 = Q1 & " AND Comprobante.Tipo <> " & TC_APERTURA
         Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
      
         Set Rs = OpenRs(DbMain, Q1)
      
         Do While Not Rs.EOF
            Descrip = gTipoComp(vFld(Rs("Tipo"))) & " " & vFld(Rs("Correlativo")) & " " & Trim(Grid.TextMatrix(r, C_CONCEPTO))
            Buf = TipoItem & Sep & Format(vFld(Rs("Fecha")), "dd/mm/yyyy") & Sep & Descrip & Sep & Abs(vFld(Rs("Valor")))
            Print #Fd, Buf
            n = n + 1
      
            Call Rs.MoveNext
         Loop
      
         Call CloseRs(Rs)
               
      End If
      
   Next r
      
   Close Fd

   If n = 0 Then
      If gEmpresa.Ano >= 2020 Then
         MsgBox1 "No existen datos para generar archivo  Ajustes Extra - Contables RLI HR RAD." & vbCrLf & vbCrLf & "Verifique si existen movimientos en sus cuentas de Ajustes Extra - Contables RLI HR RAD y si la configuración de las cuentas ha sido realizada.", vbInformation
      Else
         MsgBox1 "No existen datos para generar archivo  Ajustes Extra - Contables RLI HR RAB." & vbCrLf & vbCrLf & "Verifique si existen movimientos en sus cuentas de Ajustes Extra - Contables RLI HR RAB y si la configuración de las cuentas ha sido realizada.", vbInformation
      End If
      fname = ""
   Else
      FPath = ReplaceStr(FPath, "C:\HR\LPContab\..\", "C:\HR\")
      MsgBox1 "Proceso de exportación finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & FPath, vbInformation + vbOKOnly
   End If
      
   Export_RLI_HR_RAB = 0

End Function

