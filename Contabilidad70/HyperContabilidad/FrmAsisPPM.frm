VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAsisPPM 
   Caption         =   "Asistente Reajuste PPM"
   ClientHeight    =   7320
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
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
      Index           =   0
      Left            =   5400
      Picture         =   "FrmAsisPPM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Sumar movimientos seleccionados"
      Top             =   1200
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
      Index           =   1
      Left            =   11640
      Picture         =   "FrmAsisPPM.frx":00A4
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Sumar movimientos seleccionados"
      Top             =   1200
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Width           =   12015
      Begin MSFlexGridLib.MSFlexGrid GridTotTraspaso 
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1085
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         RowHeightMin    =   1
         MergeCells      =   1
      End
   End
   Begin VB.Frame FrmAsisPPM 
      Caption         =   "PPM Voluntario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   6240
      TabIndex        =   9
      Top             =   720
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid GridPpmVoluntario 
         Height          =   3255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   3
      End
      Begin MSFlexGridLib.MSFlexGrid GridTotVolunt 
         Height          =   1155
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3720
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   2037
         _Version        =   393216
         Rows            =   4
         Cols            =   3
         FixedRows       =   3
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   16711680
         ScrollTrack     =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "PPM Obligatorio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid GridPpmObligatorio 
         Height          =   3255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   3
      End
      Begin MSFlexGridLib.MSFlexGrid GridTotOblig 
         Height          =   1155
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3720
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   2037
         _Version        =   393216
         Rows            =   4
         Cols            =   3
         FixedRows       =   3
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   16711680
         ScrollTrack     =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12075
      Begin VB.CommandButton bt_Help 
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
         Left            =   2760
         Picture         =   "FrmAsisPPM.frx":0148
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ver formato archivo para importar cartola"
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
         Left            =   1860
         Picture         =   "FrmAsisPPM.frx":052B
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   1440
         Picture         =   "FrmAsisPPM.frx":088C
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   2280
         Picture         =   "FrmAsisPPM.frx":0C2A
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
         Left            =   960
         Picture         =   "FrmAsisPPM.frx":1053
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
         Left            =   120
         Picture         =   "FrmAsisPPM.frx":1498
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
         Left            =   540
         Picture         =   "FrmAsisPPM.frx":193F
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   10560
         TabIndex        =   1
         Top             =   180
         Width           =   1035
      End
   End
End
Attribute VB_Name = "FrmAsisPPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_FECHA_PAGO = 0
Const C_MONTO = 1
Const C_MONTO_ACTUALIZADO = 2

Const C_NOMBRE = 0
Const C_NUMLIN = 1

Const NCOLS = C_MONTO_ACTUALIZADO

Dim lOrientacion As Integer
Dim lPapelFoliado As Boolean
Dim lInfoPreliminar As Boolean

Private Sub SetUpGridAll()
   Dim iO As Integer
   Dim iV As Integer
      
   'grid ppm obligatorio
   GridPpmObligatorio.ColWidth(C_FECHA_PAGO) = 1200
   GridPpmObligatorio.ColWidth(C_MONTO) = 1300
   GridPpmObligatorio.ColWidth(C_MONTO_ACTUALIZADO) = 2100
   
   For iO = 0 To GridPpmObligatorio.Cols - 1
      GridPpmObligatorio.FixedAlignment(iO) = flexAlignCenterCenter
      GridPpmObligatorio.ColAlignment(iO) = flexAlignRightCenter

   Next iO
   GridPpmObligatorio.ColAlignment(C_FECHA_PAGO) = flexAlignLeftCenter

   GridPpmObligatorio.TextMatrix(0, C_FECHA_PAGO) = "FECHA PAGO"
   GridPpmObligatorio.TextMatrix(0, C_MONTO) = "MONTO"
   GridPpmObligatorio.TextMatrix(0, C_MONTO_ACTUALIZADO) = "MONTO ACTUALIZADO"
      
    
     Call FGrSetup(GridPpmObligatorio)
   'Call FGrTotales(GridPpmObligatorio, GridTotOblig)
   Call FGrVRows(GridPpmObligatorio)
      
      
   'grid ppm voluntario

      
   GridPpmVoluntario.ColWidth(C_FECHA_PAGO) = 1200
   GridPpmVoluntario.ColWidth(C_MONTO) = 1300
   GridPpmVoluntario.ColWidth(C_MONTO_ACTUALIZADO) = 2100
   
   For iV = 0 To GridPpmVoluntario.Cols - 1
      GridPpmVoluntario.FixedAlignment(iV) = flexAlignCenterCenter
      GridPpmVoluntario.ColAlignment(iV) = flexAlignRightCenter
   Next iV
   GridPpmVoluntario.ColAlignment(C_FECHA_PAGO) = flexAlignLeftCenter

   GridPpmVoluntario.TextMatrix(0, C_FECHA_PAGO) = "FECHA PAGO"
   GridPpmVoluntario.TextMatrix(0, C_MONTO) = "MONTO"
   GridPpmVoluntario.TextMatrix(0, C_MONTO_ACTUALIZADO) = "MONTO ACTUALIZADO"

     Call FGrSetup(GridPpmVoluntario)
   'Call FGrTotales(GridPpmVoluntario, GridTotVolunt)
   Call FGrVRows(GridPpmVoluntario)
End Sub

Private Sub LoadAllPpmObli()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Q2 As String
   Dim Rs2 As Recordset
   Dim FechaPPM As Long
   Dim TipoPPM As Boolean
   Dim Mostrar As Boolean
   
   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa WHERE Tipo='PPM'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      TipoPPM = IIf(vFld(Rs("Valor")) <> 0, True, False)
      FechaPPM = DateSerial(gEmpresa.Ano, 1, 20)
   End If
   
   Call CloseRs(Rs)
   
   
   Dim montoActualizado As String
     
   If gDbType = SQL_ACCESS Then
   Q1 = "SELECT Comprobante.FECHA,MovComprobante.DEBE "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp)"
   Q1 = Q1 & "  INNER JOIN ParamEmpresa ON MovComprobante.IdEmpresa = ParamEmpresa.IdEmpresa "
   Q1 = Q1 & " WHERE MovComprobante.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND ParamEmpresa.tipo = 'CTAPPMOBLI' AND  MovComprobante.idCuenta = int(ParamEmpresa.valor) "
   Q1 = Q1 & " AND COMPROBANTE.TIPO = " & TC_EGRESO
   Q1 = Q1 & " AND COMPROBANTE.ESTADO = " & EC_APROBADO
 ElseIf gDbType = SQL_SERVER Then
   Q1 = "SELECT Comprobante.FECHA,MovComprobante.DEBE "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp)"
   Q1 = Q1 & "  INNER JOIN ParamEmpresa ON MovComprobante.IdEmpresa = ParamEmpresa.IdEmpresa "
   Q1 = Q1 & " WHERE MovComprobante.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND ParamEmpresa.tipo = 'CTAPPMOBLI' AND  MovComprobante.idCuenta = Convert(int,ParamEmpresa.valor) "
   Q1 = Q1 & " AND COMPROBANTE.TIPO = " & TC_EGRESO
   Q1 = Q1 & " AND COMPROBANTE.ESTADO = " & EC_APROBADO
 
 End If
    
'    If TipoPPM Then
'        Q1 = Q1 & " AND Comprobante.FECHA  > " & FechaPPM
'    End If
    
    Q1 = Q1 & " ORDER BY Comprobante.FECHA ASC "

   Set Rs = OpenRs(DbMain, Q1)
   i = GridPpmObligatorio.FixedRows
   GridPpmObligatorio.rows = i
   Do While Rs.EOF = False
       If TipoPPM And vFld(Rs("Fecha")) < FechaPPM Then
        bt_Help.visible = True
        Mostrar = True
       Else
           GridPpmObligatorio.rows = i + 1
    
           Q2 = "SELECT Factor "
           Q2 = Q2 & " FROM FactorActAnual"
           Q2 = Q2 & " WHERE Ano = " & Year(vFld(Rs("Fecha")))
           Q2 = Q2 & " AND MesCol = 12 "
           Q2 = Q2 & " AND MesRow = " & month(vFld(Rs("Fecha")))
        
           Set Rs2 = OpenRs(DbMain, Q2)
          If Rs2.EOF = False Then
              
              montoActualizado = Format(vFld(Rs("DEBE")) * vFld(Rs2("Factor")), NUMFMT)
          Else
             
             montoActualizado = Format(vFld(Rs("DEBE")) * 1, NUMFMT)
             
          End If
              
          GridPpmObligatorio.TextMatrix(i, C_FECHA_PAGO) = Format(vFld(Rs("Fecha")), SDATEFMT)
          GridPpmObligatorio.TextMatrix(i, C_MONTO) = Format(vFld(Rs("DEBE")), NUMFMT)
          GridPpmObligatorio.TextMatrix(i, C_MONTO_ACTUALIZADO) = montoActualizado
             
          If vFld(Rs("Fecha")) > FechaPPM And i = 1 And Not Mostrar Then
            bt_Help.visible = False
          End If
             
          i = i + 1
       End If
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
    Call CloseRs(Rs2)
     Call CalcTotOblig
  ' Call FGrVRows(GridPpmObligatorio)
   GridPpmObligatorio.Redraw = True

End Sub

Private Sub LoadAllPpmVolun()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Q2 As String
   Dim Rs2 As Recordset
   Dim montoActualizado As String
  
   If gDbType = SQL_ACCESS Then
   Q1 = "SELECT Comprobante.FECHA,MovComprobante.DEBE "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp)"
   Q1 = Q1 & "  INNER JOIN ParamEmpresa ON MovComprobante.IdEmpresa = ParamEmpresa.IdEmpresa "
   Q1 = Q1 & " WHERE MovComprobante.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND ParamEmpresa.tipo = 'CTAPPMVOLU' AND  MovComprobante.idCuenta = int(ParamEmpresa.valor) "
   Q1 = Q1 & " AND COMPROBANTE.TIPO = " & TC_EGRESO
   Q1 = Q1 & " AND COMPROBANTE.ESTADO = " & EC_APROBADO
  ElseIf gDbType = SQL_SERVER Then
   Q1 = "SELECT Comprobante.FECHA,MovComprobante.DEBE "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp)"
   Q1 = Q1 & "  INNER JOIN ParamEmpresa ON MovComprobante.IdEmpresa = ParamEmpresa.IdEmpresa "
   Q1 = Q1 & " WHERE MovComprobante.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND ParamEmpresa.tipo = 'CTAPPMVOLU' AND  MovComprobante.idCuenta = Convert(int,ParamEmpresa.valor) "
   Q1 = Q1 & " AND COMPROBANTE.TIPO = " & TC_EGRESO
   Q1 = Q1 & " AND COMPROBANTE.ESTADO = " & EC_APROBADO
   
   End If

   Set Rs = OpenRs(DbMain, Q1)
   
   GridPpmVoluntario.Redraw = False
   i = GridPpmVoluntario.FixedRows
   GridPpmVoluntario.rows = i
   Do While Rs.EOF = False
      GridPpmVoluntario.rows = i + 1
      
      If FGrChkMaxSize(GridPpmVoluntario) = True Then
         Exit Do
      End If
      
      
      Q2 = "SELECT Factor "
           Q2 = Q2 & " FROM FactorActAnual"
           Q2 = Q2 & " WHERE Ano = " & Year(vFld(Rs("Fecha")))
           Q2 = Q2 & " AND MesCol = 12 "
           Q2 = Q2 & " AND MesRow = " & month(vFld(Rs("Fecha")))
        
           Set Rs2 = OpenRs(DbMain, Q2)
          If Rs2.EOF = False Then
              
            montoActualizado = Format(vFld(Rs("DEBE")) * vFld(Rs2("Factor")), NUMFMT)
          Else
            montoActualizado = Format(vFld(Rs("DEBE")) * 1, NUMFMT)
          End If
          
      
      GridPpmVoluntario.TextMatrix(i, C_FECHA_PAGO) = Format(vFld(Rs("Fecha")), SDATEFMT)
      GridPpmVoluntario.TextMatrix(i, C_MONTO) = Format(vFld(Rs("DEBE")), NUMFMT)
      GridPpmVoluntario.TextMatrix(i, C_MONTO_ACTUALIZADO) = montoActualizado
     
      
      i = i + 1
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
     Call CloseRs(Rs2)
   'Call FGrVRows(GridPpmVoluntario)
   GridPpmVoluntario.Redraw = True
   
     Call CalcTotVolunt

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

Private Sub bt_Cerrar_Click()
Unload Me
End Sub

Private Sub Bt_ConvMoneda_Click()
 Dim Frm As FrmConverMoneda
   Dim Valor As Double
      
   Set Frm = New FrmConverMoneda
   Frm.FView (Valor)
      
   Set Frm = Nothing
End Sub

Private Sub Bt_CopyExcel_Click()
Dim Clip As String

   'Call FGr2Clip(Grid, "Libro de Retenciones" & vbTab & "Mes: " & Cb_Mes & " " & Val(Cb_Ano))
   Clip = FGr2String(GridPpmObligatorio, Me.Caption + " Obligatorio" & vbTab & "Año " & gEmpresa.Ano, False, C_NUMLIN)
   Clip = Clip & FGr2String(GridTotOblig)
   'Clip = Clip & " Reajuste" & vbTab & GridTotOblig.TextMatrix(2, C_MONTO_ACTUALIZADO)
   Clip = Clip & "" & vbCrLf
   Clip = Clip & FGr2String(GridPpmVoluntario, Me.Caption + " Voluntario" & vbTab & "Año " & gEmpresa.Ano)
   Clip = Clip & FGr2String(GridTotVolunt)
  ' Clip = Clip & "Reajuste" & vbTab & GridTotVolunt.TextMatrix(2, C_MONTO_ACTUALIZADO)
    Clip = Clip & "" & vbCrLf
   Clip = Clip & "Total a Traspasar" & vbTab & GridTotTraspaso.TextMatrix(1, 0)
   
     

' Call LP_FGr2Clip(GridPpmObligatorio, Me.Caption + " Obligatorio" & vbTab & "Año " & gEmpresa.Ano)
  'Call LP_FGr2Clip(GridPpmVoluntario, Me.Caption + " Voluntario" & vbTab & "Año " & gEmpresa.Ano)
  
  Clipboard.Clear
   Clipboard.SetText Clip
End Sub

Private Sub bt_Help_Click()
Dim Q1 As String
If MsgBox1("¿Desea considerar los PPM pagados hasta el 20/Enero para los cálculos posteriores?", vbQuestion Or vbDefaultButton2 Or vbYesNo) = vbYes Then
    Q1 = "UPDATE ParamEmpresa SET Valor = '0' WHERE Tipo='PPM' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1, False)
Else
    Q1 = "UPDATE ParamEmpresa SET Valor = '1' WHERE Tipo='PPM' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1, False)
End If
Call Form_Load
End Sub

Private Sub Bt_Preview_Click()
 Dim Frm As FrmPrintPreview
    Dim Pag As Integer
    Dim OldOrientacion As Integer
     Dim Titulos(2) As String
   Dim Encabezados(6) As String
   Dim EncabezadosCont(6) As String
    Dim EncabezadosCont2(6) As String
   Dim FontTit(2) As FontDef_t
   Dim OldOrient As Integer
   Dim Mes As String
   Dim Idx As Integer
   Dim FntEncabezados(0) As FontDef_t
    
   Set gPrtLibros = New ClsPrtFlxGrid
    
    
   OldOrientacion = Printer.Orientation
    
   Me.MousePointer = vbHourglass
   
 Set Frm = Nothing
   
   Set Frm = New FrmPrintPreview
   
   
   Titulos(0) = "Reajuste PPM Obligatorio y Voluntario"
   
   FontTit(0).FontBold = True
         
   If lInfoPreliminar Then
      Titulos(2) = INFO_PRELIMINAR
      FontTit(2).FontBold = True
   End If
      
   gPrtLibros.Titulos = Titulos
   Call gPrtLibros.FntTitulos(FontTit())
   
'   Encabezados(0) = Lb_Contrib
   Encabezados(0) = "   "
   Encabezados(1) = "   "
   Encabezados(2) = "REAJUSTE PPM OBLIGATORIO"
   
   gPrtLibros.Encabezados = Encabezados
   
   EncabezadosCont(0) = "    "
   EncabezadosCont(1) = "    "
   EncabezadosCont(2) = "REAJUSTE PPM VOLUNTARIO"
   
   gPrtLibros.EncabezadosCont = EncabezadosCont
       
  
   FntEncabezados(0).FontName = "Arial"
   FntEncabezados(0).FontSize = 10
   FntEncabezados(0).FontBold = True
   FntEncabezados(0).FontUnderline = False
   
   
   Call gPrtLibros.FntEncabezados(FntEncabezados)
   
   
   'gPrtLibros.Obs = ""   'para que no ponga las notas
   gPrtLibros.CallEndDoc = False
        
   'Set Frm = Nothing
   Set Frm = New FrmPrintPreview
   
   
   Call SetUpPrtGrid(GridPpmObligatorio, GridTotOblig)
     'gPrtLibros.PermitirMasDe1Franja = True
   Me.MousePointer = vbHourglass
  
   If GridPpmObligatorio.rows > 1 Then
    'gPrtLibros.InitPag = 1
    Pag = gPrtLibros.PrtFlexGrid(Frm)
   'Call gPrtLibros.PrtFlexGrid(Frm)
   'Frm.NewPage
   End If
       gPrtLibros.EsContinuacion = True
   Call SetUpPrtGrid2(GridPpmVoluntario, GridTotVolunt)
   
   If GridPpmVoluntario.rows > 1 Then
   'gPrtLibros.InitPag = 2
   
   Pag = gPrtLibros.PrtFlexGrid(Frm)
   
   'Call gPrtLibros.PrtFlexGrid(Frm)
   
    End If
    'gPrtLibros.EsContinuacion = False
    
   
   EncabezadosCont2(0) = "    "
   EncabezadosCont2(1) = "    "
   EncabezadosCont2(2) = "TOTAL REAJUSTE TRASPASO"
   

   gPrtLibros.EncabezadosCont = EncabezadosCont2
   
   
   Call SetUpPrtGrid3(GridTotTraspaso)
   gPrtLibros.EsContinuacion = True


   Pag = gPrtLibros.PrtFlexGrid1(Frm)

    gPrtLibros.EsContinuacion = False
    
    gPrtLibros.CallEndDoc = True

   Set Frm.PrtControl = Bt_Print
   
   Me.MousePointer = vbDefault
   
   Call Frm.FView(Caption)
   
   Set Frm = Nothing

   
    Printer.Orientation = OldOrientacion
    Call ResetPrtBas(gPrtLibros)
   Me.MousePointer = vbDefault
End Sub

Private Sub Bt_Print_Click()
 Dim Frm As FrmPrtSetup
    Dim Pag As Integer
     Dim Titulos(2) As String
   Dim Encabezados(6) As String
   Dim EncabezadosCont(6) As String
    Dim EncabezadosCont2(6) As String
   Dim FontTit(2) As FontDef_t
   Dim OldOrient As Integer
   Dim Mes As String
   Dim Idx As Integer
   Dim FntEncabezados(0) As FontDef_t
    
   Me.MousePointer = vbHourglass
   
   Set gPrtLibros = New ClsPrtFlxGrid
   
      Set Frm = New FrmPrtSetup
   'gPrtLibros.PermitirMasDe1Franja = True
   
   Titulos(0) = "Reajuste PPM Obligatorio y Voluntario"
   
   FontTit(0).FontBold = True
         
   If lInfoPreliminar Then
      Titulos(2) = INFO_PRELIMINAR
      FontTit(2).FontBold = True
   End If
      
   gPrtLibros.Titulos = Titulos
   Call gPrtLibros.FntTitulos(FontTit())
   
'   Encabezados(0) = Lb_Contrib
   Encabezados(0) = "   "
   Encabezados(1) = "   "
   Encabezados(2) = "REAJUSTE PPM OBLIGATORIO"
   
   gPrtLibros.Encabezados = Encabezados
   
   EncabezadosCont(0) = "    "
   EncabezadosCont(1) = "    "
   EncabezadosCont(2) = "REAJUSTE PPM VOLUNTARIO"
   
   gPrtLibros.EncabezadosCont = EncabezadosCont
       
  
   FntEncabezados(0).FontName = "Arial"
   FntEncabezados(0).FontSize = 10
   FntEncabezados(0).FontBold = True
   FntEncabezados(0).FontUnderline = False
   
   
   Call gPrtLibros.FntEncabezados(FntEncabezados)
      gPrtLibros.Obs = ""   'para que no ponga las notas
   gPrtLibros.CallEndDoc = False
   
   Call SetUpPrtGrid(GridPpmObligatorio, GridTotOblig)
   
   If Frm.FEdit(lOrientacion, lPapelFoliado, lInfoPreliminar) <> vbOK Then
      Call ResetPrtBas(gPrtLibros)
      Set Frm = Nothing
      Exit Sub
   End If
  
   
   Set Frm = Nothing
   Me.MousePointer = vbHourglass
  
   If GridPpmObligatorio.rows > 1 Then
    'gPrtLibros.InitPag = 1
    Pag = gPrtLibros.PrtFlexGrid(Printer)
  ' Call gPrtLibros.PrtFlexGrid(Printer)
   'Frm.NewPage
   End If
   
   gPrtLibros.EsContinuacion = True
   Call SetUpPrtGrid2(GridPpmVoluntario, GridTotVolunt)
   
   If GridPpmVoluntario.rows > 1 Then
   'gPrtLibros.InitPag = 2
   
   Pag = gPrtLibros.PrtFlexGrid(Printer)
   
   'Call gPrtLibros.PrtFlexGrid(Printer)
   
    End If
    'gPrtLibros.EsContinuacion = False
    Call SetUpPrtGrid3(GridTotTraspaso)
   gPrtLibros.EsContinuacion = True
   
   EncabezadosCont2(0) = "    "
   EncabezadosCont2(1) = "    "
   EncabezadosCont2(2) = "TOTAL REAJUSTE TRASPASO"
   
   gPrtLibros.EncabezadosCont = EncabezadosCont2
   
   Pag = gPrtLibros.PrtFlexGrid1(Printer)
  ' Call gPrtLibros.PrtFlexGrid1(Printer)
   
    gPrtLibros.EsContinuacion = False
    gPrtLibros.CallEndDoc = True
    Printer.EndDoc
    
    'gPrtLibros.PermitirMasDe1Franja = False
    
   Me.MousePointer = vbDefault

   Set Frm = Nothing
    Call ResetPrtBas(gPrtLibros)
   Me.MousePointer = vbDefault
End Sub

Private Sub Bt_Sum_Click(Index As Integer)
 Dim Frm As FrmSumSimple
If Index = 0 Then

 

   Set Frm = New FrmSumSimple

   Call Frm.FViewSum(GridPpmObligatorio)

   Set Frm = Nothing
   
  Else
    Set Frm = New FrmSumSimple

   Call Frm.FViewSum(GridPpmVoluntario)

   Set Frm = Nothing
   
  End If
End Sub

Private Sub Form_Load()
lOrientacion = ORIENT_VER
SetUpGridAll
LoadAllPpmObli
LoadAllPpmVolun

CalcTotTraspasar
End Sub


Private Sub CalcTotOblig()
   Dim Tot(NCOLS) As Double
   Dim i As Integer, j As Integer
   
   For i = GridPpmObligatorio.FixedRows To GridPpmObligatorio.rows - 1
      If GridPpmObligatorio.TextMatrix(i, C_FECHA_PAGO) = "" Then
         Exit For
      End If
      For j = C_MONTO To C_MONTO_ACTUALIZADO
         If j <> C_FECHA_PAGO Then
            Tot(j) = Tot(j) + vFmt(GridPpmObligatorio.TextMatrix(i, j))
         End If
      Next j
   Next i
   
   GridTotOblig.TextMatrix(0, C_NOMBRE) = "TOTALES"
   GridTotOblig.TextMatrix(0, C_MONTO) = Format(Tot(C_MONTO), NUMFMT)
   GridTotOblig.TextMatrix(0, C_MONTO_ACTUALIZADO) = Format(Tot(C_MONTO_ACTUALIZADO), NUMFMT)
   
   GridTotOblig.TextMatrix(2, C_NOMBRE) = "REAJUSTE"
   GridTotOblig.TextMatrix(2, C_MONTO) = ""
   GridTotOblig.TextMatrix(2, C_MONTO_ACTUALIZADO) = Format(Tot(C_MONTO_ACTUALIZADO) - Tot(C_MONTO), NUMFMT)
   
   ' Tx_ReajusteOblig.Text = Format(Tot(C_MONTO) + Tot(C_MONTO_ACTUALIZADO), NUMFMT)
   
End Sub

Private Sub CalcTotVolunt()
   Dim Tot(NCOLS) As Double
   Dim i As Integer, j As Integer
   
   For i = GridPpmVoluntario.FixedRows To GridPpmVoluntario.rows - 1
      If GridPpmVoluntario.TextMatrix(i, C_FECHA_PAGO) = "" Then
         Exit For
      End If
      For j = C_MONTO To C_MONTO_ACTUALIZADO
         If j <> C_FECHA_PAGO Then
            Tot(j) = Tot(j) + vFmt(GridPpmVoluntario.TextMatrix(i, j))
         End If
      Next j
   Next i
   
   GridTotVolunt.TextMatrix(0, C_NOMBRE) = "TOTALES "
   GridTotVolunt.TextMatrix(0, C_MONTO) = Format(Tot(C_MONTO), NUMFMT)
   GridTotVolunt.TextMatrix(0, C_MONTO_ACTUALIZADO) = Format(Tot(C_MONTO_ACTUALIZADO), NUMFMT)
   
   GridTotVolunt.TextMatrix(2, C_NOMBRE) = "REAJUSTE"
   GridTotVolunt.TextMatrix(2, C_MONTO) = ""
   GridTotVolunt.TextMatrix(2, C_MONTO_ACTUALIZADO) = Format(Tot(C_MONTO_ACTUALIZADO) - Tot(C_MONTO), NUMFMT)
  
  
  'Tx_ReajusteVolunt.Text = Format(Tot(C_MONTO) + Tot(C_MONTO_ACTUALIZADO), NUMFMT)
  
   
End Sub

Private Sub CalcTotTraspasar()
   Dim i As Integer
     
   GridTotTraspaso.ColWidth(C_NOMBRE) = 1200

   
      GridTotTraspaso.FixedAlignment(0) = flexAlignLeftCenter
      GridTotTraspaso.ColAlignment(0) = flexAlignRightCenter


      GridTotTraspaso.FixedAlignment(C_NOMBRE) = flexAlignCenterCenter
   
   GridTotTraspaso.TextMatrix(0, C_NOMBRE) = "TOTAL"
   GridTotTraspaso.TextMatrix(1, 0) = Format(Int(Replace(GridTotVolunt.TextMatrix(2, C_MONTO_ACTUALIZADO), ",", "")) + Int(Replace(GridTotOblig.TextMatrix(2, C_MONTO_ACTUALIZADO), ",", "")), NUMFMT)

   
   'GridTotTraspaso.RowHeight(1) = 0
   
   
   
End Sub


Private Sub SetUpPrtGrid(Grid As Object, GridTot As Object)
   Dim i As Integer
   Dim r As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(11) As String
   Dim Titulos(1) As String
   
   'Printer.Orientation = ORIENT_VER
   Set gPrtLibros.Grid = Grid
   
   'Titulos(0) = Tit
   'Titulos(1) = "Fecha: " & Tx_Desde & " al " & Tx_Hasta
  ' gPrtLibros.Titulos = Titulos
         
   gPrtLibros.GrFontName = Grid.FontName
   gPrtLibros.GrFontSize = Grid.FontSize
  
   For i = 0 To Grid.Cols - 1
      'Grid.FixedAlignment(i) = flexAlignRightCenter
      
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
   For i = 0 To GridTot.Cols - 1
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
   
   For r = 0 To 2 - 1
   i = i + 1
   Total(i) = ""
   
   Next r
   
    For r = 0 To GridTot.Cols - 1
    i = i + 1
      Total(i) = GridTot.TextMatrix(2, r)
   Next r
   
   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
  
   
   'gPrtReportes.ColObligatoria = C_FECHA
   gPrtLibros.NTotLines = 3

End Sub

Private Sub SetUpPrtGrid2(Grid As Object, GridTot As Object)
   Dim i As Integer
   Dim r As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(11) As String
   Dim Titulos(1) As String
   
   'Printer.Orientation = ORIENT_VER
   Set gPrtLibros.Grid = Grid
   
   'Titulos(0) = Tit
   'Titulos(1) = "Fecha: " & Tx_Desde & " al " & Tx_Hasta
   'gPrtLibros.Titulos = Titulos
         
   gPrtLibros.GrFontName = Grid.FontName
   gPrtLibros.GrFontSize = Grid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
                  
   For i = 0 To GridTot.Cols - 1
      Total(i) = GridTot.TextMatrix(0, i)
   Next i
   
    For r = 0 To 2 - 1
   i = i + 1
   Total(i) = ""
   
   Next r
   
    For r = 0 To GridTot.Cols - 1
    i = i + 1
      Total(i) = GridTot.TextMatrix(2, r)
   Next r
      
   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
   'gPrtReportes.ColObligatoria = C_FECHA
   gPrtLibros.NTotLines = 3

End Sub


Private Sub SetUpPrtGrid3(Grid As Object)
   Dim i As Integer
   Dim r As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(11) As String
   Dim Titulos(1) As String
  

   Set gPrtLibros.Grid = Grid
         
   gPrtLibros.GrFontName = Grid.FontName
   gPrtLibros.GrFontSize = Grid.FontSize
   

   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i

      gPrtReportes.ColObligatoria = 0
   gPrtLibros.ColWi = ColWi
   gPrtLibros.NTotLines = 0
 

End Sub

