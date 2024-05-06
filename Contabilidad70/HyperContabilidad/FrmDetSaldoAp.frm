VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmDetSaldoAp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Saldo de Apertura"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Detalle 
      Height          =   6015
      Left            =   60
      TabIndex        =   11
      Top             =   660
      Width           =   8895
      Begin VB.TextBox Tx_TotSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   7320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox Tx_TotDebe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox Tx_TotHaber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox Tx_Cuenta 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   5055
      End
      Begin FlexEdGrid2.FEd2Grid Grid 
         Height          =   4875
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8599
         Cols            =   8
         Rows            =   10
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   8895
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
         Picture         =   "FrmDetSaldoAp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Eliminar detalle seleccionado"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_SelEnt 
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
         Picture         =   "FrmDetSaldoAp.frx":03FC
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Seleccionar Entidad"
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
         Left            =   1560
         Picture         =   "FrmDetSaldoAp.frx":089A
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   1140
         Picture         =   "FrmDetSaldoAp.frx":0D54
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   2040
         Picture         =   "FrmDetSaldoAp.frx":11FB
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   3960
         Picture         =   "FrmDetSaldoAp.frx":1640
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   3120
         Picture         =   "FrmDetSaldoAp.frx":1A69
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   3540
         Picture         =   "FrmDetSaldoAp.frx":1E07
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   2580
         Picture         =   "FrmDetSaldoAp.frx":2168
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Sumar movimientos seleccionados"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Ok 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   6480
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Cancel 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   7680
         TabIndex        =   2
         Top             =   180
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmDetSaldoAp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDENTIDAD = 0
Const C_RUT = 1
Const C_NOMBRE = 2
Const C_DEBE = 3
Const C_HABER = 4
Const C_SALDO = 5
Const C_ID = 6
Const C_UPDATE = 7

Const NCOLS = C_UPDATE

Dim lIdCuenta As Long

Dim lOrientacion As Integer
Dim lRc As Integer
Dim lOper As Integer
Dim lMsgInf As Boolean
Dim lActTotCuenta As Boolean

Public Function FEdit(ByVal IdCuenta As Long, ByVal ActTotCuenta As Boolean) As Integer
   lOper = O_EDIT
   lIdCuenta = IdCuenta
   lActTotCuenta = ActTotCuenta
   
   Me.Show vbModal
   FEdit = lRc

End Function
Public Function FView(ByVal IdCuenta As Long) As Integer
   lOper = O_VIEW
   lIdCuenta = IdCuenta
   
   Me.Show vbModal
   FView = lRc

End Function
Private Sub SetUpGrid()
   Dim Col As Integer
   
   Grid.ColWidth(C_IDENTIDAD) = 0
   Grid.ColWidth(C_RUT) = 1300
   Grid.ColWidth(C_NOMBRE) = 3100
   Grid.ColWidth(C_DEBE) = 1300
   Grid.ColWidth(C_HABER) = 1300
   Grid.ColWidth(C_SALDO) = 1300
   Grid.ColWidth(C_ID) = 0
   Grid.ColWidth(C_UPDATE) = 0
   
   For Col = 0 To Grid.Cols - 1
      Grid.FixedAlignment(Col) = flexAlignCenterCenter
   Next Col
   
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_NOMBRE) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   Grid.ColAlignment(C_SALDO) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_RUT) = "RUT"
   Grid.TextMatrix(0, C_NOMBRE) = "Nombre"
   Grid.TextMatrix(0, C_DEBE) = "Debe"
   Grid.TextMatrix(0, C_HABER) = "Haber"
   Grid.TextMatrix(0, C_SALDO) = "Saldo"
      
   Call FGrLocateCntrl(Grid, Tx_TotDebe, C_DEBE)
   Call FGrLocateCntrl(Grid, Tx_TotHaber, C_HABER)
   Call FGrLocateCntrl(Grid, Tx_TotSaldo, C_SALDO)

   
End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub bt_OK_Click()
   
   If Valida() Then
      Call SaveAll
      lRc = vbOK
      Unload Me
   End If
   
End Sub

Private Sub Form_Activate()

   If Not lMsgInf Then
      MsgBox1 "Si desea ingresar el detalle de cada documento de Compras, Ventas o Retenciones, en vez de un total por RUT, utilice el botón de Ingresar/Editar Documentos, seleccionando el año anterior y el mes correspondiente." & vbNewLine & vbNewLine & "Si ingresa el detalle de los documentos, no ingrese al detalle de saldo de apertura totalizado por RUT.", vbInformation + vbOKOnly
      lMsgInf = True
   End If
   
End Sub

Private Sub Form_Load()

   lMsgInf = False
   
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   Call SetupPriv
   
   Call SetTxRO(Tx_Cuenta, True)
   Call SetTxRO(Tx_TotDebe, True)
   Call SetTxRO(Tx_TotHaber, True)
   Call SetTxRO(Tx_TotSaldo, True)
   
   Call SetUpGrid
   Call LoadAll

End Sub
Private Sub LoadAll()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim i As Integer
   Dim TotDebe As Double
   Dim TotHaber As Double

   Tx_Cuenta = GetDescCuenta(lIdCuenta)
   
   Q1 = "SELECT Id, DetSaldosAp.IdEntidad, Rut, Nombre, NotValidRut, Debe, Haber "
   Q1 = Q1 & " FROM DetSaldosAp INNER JOIN Entidades ON DetSaldosAp.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " AND DetSaldosAp.IdEmpresa = Entidades.IdEmpresa "
   Q1 = Q1 & " WHERE DetSaldosAp.IdCuenta = " & lIdCuenta
   Q1 = Q1 & " AND DetSaldosAp.IdEmpresa = " & gEmpresa.id & " AND DetSaldosAp.Ano = " & gEmpresa.Ano
  
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Do While Rs.EOF = False
   
      Grid.rows = Grid.rows + 1
      Grid.TextMatrix(i, C_ID) = vFld(Rs("Id"))
      Grid.TextMatrix(i, C_IDENTIDAD) = vFld(Rs("IdEntidad"))
      Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("Rut")), vFld(Rs("NotValidRut")) = False)
      Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"))
      Grid.TextMatrix(i, C_DEBE) = Format(vFld(Rs("Debe")), NUMFMT)
      Grid.TextMatrix(i, C_HABER) = Format(vFld(Rs("Haber")), NUMFMT)
      Grid.TextMatrix(i, C_SALDO) = Format(vFld(Rs("Debe")) - vFld(Rs("Haber")), NEGNUMFMT)
      
      TotDebe = TotDebe + vFmt(Grid.TextMatrix(i, C_DEBE))
      TotHaber = TotHaber + vFmt(Grid.TextMatrix(i, C_HABER))
      
      Rs.MoveNext
      i = i + 1
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid)
   
   Grid.rows = Grid.rows + 1
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.RowSel = Grid.Row
   Grid.Col = C_RUT
   Grid.ColSel = Grid.Col
   
   Tx_TotDebe = Format(TotDebe, NUMFMT)
   Tx_TotHaber = Format(TotHaber, NUMFMT)
   Tx_TotSaldo = Format(TotDebe - TotHaber, NEGNUMFMT)
   
   
End Sub
Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Me.Caption)
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
Private Sub Bt_Sum_Click()
   Dim Frm As FrmSumMov
   
   Set Frm = New FrmSumMov
   
   Call Frm.FViewSum(Grid, C_DEBE, C_HABER)
   
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
   Dim Titulos(0) As String
   Dim Encabezados(0) As String
   
   Printer.Orientation = lOrientacion
   
   Set gPrtReportes.Grid = Grid
   
   Titulos(0) = Caption
   
   gPrtReportes.Titulos = Titulos
      
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FlxGrid.FontName
   gPrtReportes.GrFontSize = Grid.FlxGrid.FontSize
   
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
               
   Total(C_NOMBRE) = "Total"
   Total(C_DEBE) = Tx_TotDebe
   Total(C_HABER) = Tx_TotHaber
   Total(C_SALDO) = Tx_TotSaldo
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.Total = Total
   gPrtReportes.ColObligatoria = C_IDENTIDAD
   gPrtReportes.NTotLines = 1
   

End Sub
Private Sub SaveAll()
   Dim i As Integer, j As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim id As Long
   Dim Debe As Double
   Dim Haber As Double
   Dim FldArray(2) As AdvTbAddNew_t
   
   If lIdCuenta <= 0 Then
      Exit Sub
   End If
         
   For i = Grid.FixedRows To Grid.rows - 1
            
      If Grid.TextMatrix(i, C_IDENTIDAD) = "" Then    'ya terminó la lista
         Exit For
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_I Then  'Insert
'         Set Rs = DbMain.OpenRecordset("DetSaldosAp")
'         Rs.AddNew
'
'         id = Rs("Id")
'         Rs("IdCuenta") = lIdCuenta
'         Rs.Update
'         Rs.Close
'         Set Rs = Nothing
         

         FldArray(0).FldName = "IdCuenta"
         FldArray(0).FldValue = lIdCuenta
         FldArray(0).FldIsNum = True
                        
         FldArray(1).FldName = "IdEmpresa"
         FldArray(1).FldValue = gEmpresa.id
         FldArray(1).FldIsNum = True
                     
         FldArray(2).FldName = "Ano"
         FldArray(2).FldValue = gEmpresa.Ano
         FldArray(2).FldIsNum = True
                 
         id = AdvTbAddNewMult(DbMain, "DetSaldosAp", "Id", FldArray)
         
         Grid.TextMatrix(i, C_ID) = id
         Grid.TextMatrix(i, C_UPDATE) = FGR_U       'para que ahora pase por el update
                  
      End If
      
      If Grid.TextMatrix(i, C_UPDATE) = FGR_D Then  'Delete
'         Q1 = "DELETE FROM DetSaldosAp WHERE Id = " & Grid.TextMatrix(i, C_ID)
'         Call ExecSQL(DbMain, Q1)
         Q1 = " WHERE Id = " & Grid.TextMatrix(i, C_ID)
         Call DeleteSQL(DbMain, "DetSaldosAp", Q1)
                  
      ElseIf Grid.TextMatrix(i, C_UPDATE) = FGR_U Then 'Update
         Q1 = "UPDATE DetSaldosAp SET "
         Q1 = Q1 & "  IdEntidad = " & Val(Grid.TextMatrix(i, C_IDENTIDAD))
         'Q1 = Q1 & ", IdCuenta = " & lIdCuenta
         Q1 = Q1 & ", Debe = " & vFmt(Grid.TextMatrix(i, C_DEBE))
         Q1 = Q1 & ", Haber = " & vFmt(Grid.TextMatrix(i, C_HABER))
         Q1 = Q1 & ", Saldo = " & vFmt(Grid.TextMatrix(i, C_SALDO))
         Q1 = Q1 & "  WHERE Id = " & Grid.TextMatrix(i, C_ID)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
         
      End If
                       
   Next i
   
   If lActTotCuenta Then
   
      If MsgBox1("Desea actualizar el Saldo de Apertura de la cuenta:" & vbNewLine & vbNewLine & Tx_Cuenta & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
      
      'actualizamos el saldo de apertura de la cuenta
      If vFmt(Tx_TotSaldo) > 0 Then
         Debe = vFmt(Tx_TotSaldo)
         Haber = 0
      Else
         Debe = 0
         Haber = Abs(vFmt(Tx_TotSaldo))
      End If
      
      Q1 = "UPDATE Cuentas SET Debe =" & Debe & ", Haber =" & Haber & " WHERE IdCuenta = " & lIdCuenta
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
      
   Else
      MsgBox1 "No se actualizará el Saldo de Apertura de la cuenta:" & vbNewLine & vbNewLine & Tx_Cuenta, vbInformation + vbOKOnly
   
   End If
      
End Sub

Private Sub SetupPriv()

   If Not (ChkPriv(PRV_CFG_EMP) And lOper = O_EDIT) Then
      Grid.Locked = True
   End If
   
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Fr_Detalle.Height = Me.Height - Fr_Detalle.Top - 500
   Grid.Height = Fr_Detalle.Height - Tx_TotDebe.Height - Tx_Cuenta.Height - 500
   
   'Grid.Width = Me.Width - 230
   Tx_TotDebe.Top = Grid.Top + Grid.Height + 60
   Tx_TotHaber.Top = Tx_TotDebe.Top
   Tx_TotSaldo.Top = Tx_TotDebe.Top
   Call FGrLocateCntrl(Grid, Tx_TotDebe, C_DEBE)
   Call FGrLocateCntrl(Grid, Tx_TotHaber, C_HABER)
   Call FGrLocateCntrl(Grid, Tx_TotSaldo, C_SALDO)
   
   Call FGrVRows(Grid)

End Sub
Private Sub Grid_AcceptValue(ByVal Row As Integer, ByVal Col As Integer, Value As String, Action As Integer)
   Dim AuxRut As String
   Dim Nombre As String
   Dim NotValidRut As Boolean
   Dim IdEnt As Long
   
   Action = vbOK
   
   Select Case Col
   
      Case C_RUT
      
         If Trim(Value) = "" Then
            Grid.TextMatrix(Row, C_IDENTIDAD) = 0
            Grid.TextMatrix(Row, C_NOMBRE) = ""
            Action = vbOK
                  
         ElseIf Trim(Value) = "0-0" Then
            MsgBox1 "RUT inválido.", vbExclamation
            Grid.TextMatrix(Row, C_RUT) = Value
            Grid.TextMatrix(Row, C_IDENTIDAD) = 0
            Grid.TextMatrix(Row, C_NOMBRE) = ""
            Action = vbRetry
                  
         Else
                     
            IdEnt = GetIdEntidad(Trim(Value), Nombre, NotValidRut)
            
            If IdEnt <= 0 Then
               If MsgBox1("Esta entidad no ha sido ingresada a la lista de entidades predefinidas." & vbNewLine & vbNewLine & "¿Desea agregarla ahora?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
                  AuxRut = FmtCID(vFmtCID(Value))
                  If AuxRut = "0-0" Then
                     AuxRut = Trim(Value)
                  End If
                  If NewEntidad(Row, AuxRut) <> vbOK Then
                     Value = ""
                     Action = vbCancel
                  Else
                     Action = vbOK
                  End If
               Else
                  Value = ""
                  Grid.TextMatrix(Row, C_IDENTIDAD) = 0
                  Grid.TextMatrix(Row, C_NOMBRE) = ""
                  Action = vbCancel
               End If
            Else
               Value = FmtCID(vFmtCID(Value, NotValidRut = False), NotValidRut = False)
               Grid.TextMatrix(Row, C_NOMBRE) = Nombre
               Grid.TextMatrix(Row, C_IDENTIDAD) = IdEnt
               Action = vbOK
            End If
                     
         End If
         
         
      Case C_DEBE
      
         If vFmt(Value) < 0 Then
            MsgBeep vbExclamation
            Action = vbCancel
            
         Else
            Value = Format(vFmt(Value), BL_NUMFMT)
            Grid.TextMatrix(Row, Col) = Value
            
            If vFmt(Grid.TextMatrix(Row, C_HABER)) <> 0 And vFmt(Value) <> 0 Then
               Grid.TextMatrix(Row, C_HABER) = ""
            End If
            
            Grid.TextMatrix(Row, C_SALDO) = Format(vFmt(Grid.TextMatrix(Row, C_DEBE)) - vFmt(Grid.TextMatrix(Row, C_HABER)), NEGNUMFMT)
            
            Call CalcTot
            
         End If
         
      Case C_HABER
   
         If vFmt(Value) < 0 Then
            MsgBeep vbExclamation
            Action = vbCancel
            
         Else
            Value = Format(vFmt(Value), BL_NUMFMT)
            Grid.TextMatrix(Row, Col) = Value
            
            If vFmt(Grid.TextMatrix(Row, C_DEBE)) <> 0 And vFmt(Value) <> 0 Then
               Grid.TextMatrix(Row, C_DEBE) = ""
            End If
            
            Grid.TextMatrix(Row, C_SALDO) = Format(vFmt(Grid.TextMatrix(Row, C_DEBE)) - vFmt(Grid.TextMatrix(Row, C_HABER)), NEGNUMFMT)
            
            Call CalcTot
            
         End If
                  
   End Select
   
   If Action = vbOK Then
      Call FGrModRow(Grid, Row, FGR_U, C_ID, C_UPDATE)
   End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Integer, ByVal Col As Integer, EdType As FEG2_EdType)
   Dim ValPrevLine As Boolean
   Dim Msg As String
   Dim Q1 As String
   Dim Rs As Recordset
         
   'Linea anterior tiene valor o está eliminada?
   ValPrevLine = (Row > Grid.FixedRows And Val(Grid.TextMatrix(Row - 1, C_IDENTIDAD)) > 0 And (Val(Grid.TextMatrix(Row - 1, C_DEBE)) > 0 Or Val(Grid.TextMatrix(Row - 1, C_HABER)) > 0)) Or Grid.RowHeight(Row - 1) = 0
   
   If Not (Row = Grid.FixedRows Or ValPrevLine) Then
      Exit Sub
   End If
   
   'sólo pueden ingresar valores en debe, haber si seleccionó una entidad
   If Col <> C_RUT And Grid.TextMatrix(Row, C_RUT) = "" Then
      Exit Sub
   End If
      
   If Val(Grid.TextMatrix(Row, C_ID)) = 0 Then    'nuevo
      If Row >= Grid.rows - 2 Then
         Grid.rows = Grid.rows + 1
      End If
     
   End If
   
   
   Select Case Col
   
      Case C_RUT
         Grid.TxBox.MaxLength = 13
         EdType = FEG_Edit
                     
      Case C_DEBE
         EdType = FEG_Edit
         
      Case C_HABER
         EdType = FEG_Edit
      
   End Select
End Sub

Private Sub Bt_SelEnt_Click()
   Dim Frm As FrmEntidades
   Dim Entidad As Entidad_t
   Dim Row As Integer
   Dim TipoEnt As Integer
   Dim Col As Integer
   Dim Rc As Integer
      
   Col = Grid.Col
   Row = Grid.Row
   
   TipoEnt = 0
      
   Set Frm = New FrmEntidades
   Rc = Frm.FSelEdit(Entidad, TipoEnt)
   Set Frm = Nothing
   
   If Rc <> vbOK Then
      Exit Sub
   End If
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If (Grid.Col <> C_RUT And Grid.Col <> C_NOMBRE) Or Grid.Locked Then
      Exit Sub
   End If
            
   Grid.TextMatrix(Row, C_IDENTIDAD) = Entidad.id
   Grid.TextMatrix(Row, C_RUT) = FmtCID(Entidad.Rut, Entidad.NotValidRut = False)
   Grid.TextMatrix(Row, C_NOMBRE) = Entidad.Nombre
   
   Call FGrModRow(Grid, Grid.FlxGrid.Row, FGR_U, C_ID, C_UPDATE)
      
End Sub

Private Function NewEntidad(ByVal Row As Integer, ByVal Rut As String) As Integer
   Dim Frm As FrmEntidad
   Dim Entidad As Entidad_t
   Dim i As Integer
   Dim Rc As Integer
 
   Set Frm = New FrmEntidad
   
   MousePointer = vbHourglass
   Entidad.Clasif = 0    'ItemData(Cb_Entidad)
   Entidad.Rut = Rut
   
   Rc = Frm.FNew(Entidad, Rut)
   
   If Rc <> vbCancel Then
            
      Grid.TextMatrix(Row, C_NOMBRE) = Entidad.Nombre
      Grid.TextMatrix(Row, C_IDENTIDAD) = Entidad.id
      Grid.TextMatrix(Row, C_RUT) = Entidad.Rut
         
   Else
      Grid.TextMatrix(Row, C_NOMBRE) = ""
      Grid.TextMatrix(Row, C_IDENTIDAD) = 0
      
   End If
   
   Set Frm = Nothing
   MousePointer = vbDefault
   
   NewEntidad = Rc
End Function
Private Sub CalcTot()
   Dim TotDebe As Double
   Dim TotHaber As Double
   Dim i As Integer
   
   TotDebe = 0
   TotHaber = 0
   
   For i = Grid.FixedRows To Grid.rows - 1
      If Grid.RowHeight(i) > 0 Then     ' no está borrado
         TotDebe = TotDebe + vFmt(Grid.TextMatrix(i, C_DEBE))
         TotHaber = TotHaber + vFmt(Grid.TextMatrix(i, C_HABER))
      End If
   Next i
         
   Tx_TotDebe = Format(TotDebe, BL_NUMFMT)
   Tx_TotHaber = Format(TotHaber, BL_NUMFMT)
   Tx_TotSaldo = Format(TotDebe - TotHaber, NEGNUMFMT)
   
End Sub

Private Function Valida() As Boolean
   Dim i As Integer
   Dim j As Integer
   
   Valida = False
   
   For i = Grid.FixedRows To Grid.rows - 1
   
      If Grid.TextMatrix(i, C_IDENTIDAD) = "" Then
         Exit For
      End If
   
      For j = i + 1 To Grid.rows - 1
      
         If Grid.TextMatrix(j, C_IDENTIDAD) = "" Then
            Exit For
         End If
      
         If vFmt(Grid.TextMatrix(j, C_IDENTIDAD)) = vFmt(Grid.TextMatrix(i, C_IDENTIDAD)) Then
            MsgBox1 "La entidad " & Grid.TextMatrix(j, C_NOMBRE) & " aparece más de una vez en la lista de detalle de saldos de apertura. Sume los totales en una sola línea.", vbExclamation + vbOKOnly
            Exit Function
         End If
         
      Next j
      
   Next i
   
   Valida = True
End Function
Private Sub Bt_Del_Click()
   Dim Row As Integer
   
   Row = Grid.Row
   
   If Grid.TextMatrix(Row, C_RUT) = "" Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea borrar este detalle de saldo?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
      
   Call FGrModRow(Grid, Row, FGR_D, C_ID, C_UPDATE)
      
   Grid.rows = Grid.rows + 1
      
   Call CalcTot
End Sub

