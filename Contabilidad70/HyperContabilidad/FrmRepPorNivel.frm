VERSION 5.00
Object = "{00459760-A356-47A6-9F74-38C489C6D169}#1.1#0"; "FlexEdGrid2.ocx"
Begin VB.Form FrmRepPorNivel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   Icon            =   "FrmRepPorNivel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   1
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   10335
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
         Picture         =   "FrmRepPorNivel.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   9000
         TabIndex        =   11
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton Bt_DetComp 
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
         Picture         =   "FrmRepPorNivel.frx":04C6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Vista previa de la impresión"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Definición del contenido"
      Height          =   855
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   10335
      Begin VB.ComboBox Cb_Nivel 
         Height          =   315
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   300
         Width           =   795
      End
      Begin VB.CommandButton Bt_Fecha 
         Caption         =   "?"
         Height          =   315
         Index           =   1
         Left            =   4260
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   300
         Width           =   215
      End
      Begin VB.CommandButton Bt_Fecha 
         Caption         =   "?"
         Height          =   315
         Index           =   0
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   300
         Width           =   215
      End
      Begin VB.TextBox Tx_Hasta 
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox Tx_Desde 
         Height          =   315
         Left            =   780
         TabIndex        =   3
         Top             =   300
         Width           =   1035
      End
      Begin VB.CommandButton Bt_Buscar 
         Height          =   435
         Left            =   9000
         Picture         =   "FrmRepPorNivel.frx":096D
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Buscar una cuenta"
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel de detalle:"
         Height          =   195
         Index           =   0
         Left            =   4980
         TabIndex        =   14
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Width           =   555
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Index           =   7
         Left            =   2700
         TabIndex        =   7
         Top             =   360
         Width           =   465
      End
   End
   Begin FlexEdGrid2.FEd2Grid Grid 
      Height          =   5715
      Left            =   60
      TabIndex        =   0
      Top             =   1620
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10081
      Cols            =   6
      Rows            =   20
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
End
Attribute VB_Name = "FrmRepPorNivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDCUENTA = 0
Const C_NIVEL = 1
Const C_CODCUENTA = 2
Const C_CUENTA = 3
Const C_DEBE = 4
Const C_HABER = 5


Private Sub Bt_Buscar_Click()
   Call LoadAll
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim i As Integer

   Call SetUpGrid
   
   Call SetTxDate(Tx_Desde, DateSerial(gEmpresa.Ano, 1, 1))
   If gEmpresa.Ano < Year(Now) Then
      Call SetTxDate(Tx_Hasta, DateSerial(gEmpresa.Ano, 12, 31))
   Else
      Call SetTxDate(Tx_Hasta, Now)
   End If
   
   For i = 1 To MAX_NIVELES
      Cb_Nivel.AddItem i
   Next i
   Cb_Nivel.ListIndex = 0
   
   Call LoadAll

End Sub
Private Sub SetUpGrid()
   Dim i As Integer
   
   Call FGrVRows(Grid)
   
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_NIVEL) = 0   'solo para debugging
   Grid.ColWidth(C_CODCUENTA) = 2000
   Grid.ColWidth(C_CUENTA) = 5000
   Grid.ColWidth(C_DEBE) = 1500
   Grid.ColWidth(C_HABER) = 1500
      
   Grid.ColAlignment(C_IDCUENTA) = flexAlignRightCenter
   Grid.ColAlignment(C_CODCUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
    
   
   Grid.TextMatrix(0, C_IDCUENTA) = ""
   Grid.TextMatrix(0, C_CODCUENTA) = "Cuenta"
   Grid.TextMatrix(0, C_CUENTA) = "Descripción"
   Grid.TextMatrix(0, C_DEBE) = "Debe"
   Grid.TextMatrix(0, C_HABER) = "Haber"
   
   For i = 0 To Grid.Cols - 1
      Grid.FixedAlignment(i) = flexAlignCenterCenter
   Next i

End Sub

Private Sub LoadAll()
   Dim Q1 As String
   Dim Nivel As Integer
   Dim Rs As Recordset
   Dim Total(MAX_NIVELES) As RepNiv_t
   Dim CurNiv As Integer
   Dim CurCta As String
   Dim i As Integer, j As Integer
   Dim WhereFecha As String
   Dim JoinComp As String
   
   Grid.FlxGrid.Redraw = False
   
   Nivel = Val(Cb_Nivel)
   WhereFecha = " (Comprobante.Fecha BETWEEN " & GetTxDate(Tx_Desde) & " AND " & GetTxDate(Tx_Hasta) & ")"
   JoinComp = " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "


   'lista de cuentas de menor nivel
   Q1 = "SELECT Cuentas.idCuenta, Codigo, Nivel, Descripcion, 0 as Debe, 0 As Haber"
   Q1 = Q1 & " FROM Cuentas"
   Q1 = Q1 & " WHERE Nivel <= " & Nivel

   Q1 = Q1 & " UNION"

   'lista de cuentas con nivel igual
   Q1 = Q1 & " SELECT Cuentas.idCuenta, Cuentas.Codigo, Cuentas.Nivel, Cuentas.Descripcion, MovComprobante.Debe, MovComprobante.Haber"
   Q1 = Q1 & " FROM (Cuentas INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta)" & JoinComp
   Q1 = Q1 & " WHERE Cuentas.Nivel <= " & Nivel & " AND " & WhereFecha

   Q1 = Q1 & " UNION"

   'suma de cuentas en que este nivel es el padre
   Q1 = Q1 & " SELECT Cuentas_1.idCuenta, Cuentas_1.Codigo, Cuentas_1.Nivel, Cuentas_1.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber"
   Q1 = Q1 & " FROM ((Cuentas INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta) INNER JOIN Cuentas AS Cuentas_1 ON Cuentas.idPadre = Cuentas_1.idCuenta)" & JoinComp
   Q1 = Q1 & " WHERE Cuentas_1.Nivel = " & Nivel & " AND " & WhereFecha
   Q1 = Q1 & " GROUP BY Cuentas_1.idCuenta, Cuentas_1.Codigo, Cuentas_1.Nivel, Cuentas_1.Descripcion"
   
   Q1 = Q1 & " UNION"
   
   'suma de cuentas en que este nivel es el abuelo
   Q1 = Q1 & " SELECT Cuentas_2.idCuenta, Cuentas_2.Codigo, Cuentas_2.Nivel, Cuentas_2.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber"
   Q1 = Q1 & " FROM (((Cuentas INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta) INNER JOIN Cuentas AS Cuentas_1 ON Cuentas.idPadre = Cuentas_1.idCuenta) INNER JOIN Cuentas AS Cuentas_2 ON Cuentas_1.idPadre = Cuentas_2.idCuenta)" & JoinComp
   Q1 = Q1 & " Where Cuentas_2.Nivel = " & Nivel & " AND " & WhereFecha
   Q1 = Q1 & " GROUP BY Cuentas_2.idCuenta, Cuentas_2.Codigo, Cuentas_2.Nivel, Cuentas_2.Descripcion"
   
   Q1 = Q1 & " UNION"
   
   'suma de cuentas en que este nivel es el bis-abuelo
   Q1 = Q1 & " SELECT Cuentas_3.idCuenta, Cuentas_3.Codigo, Cuentas_3.Nivel, Cuentas_3.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber"
   Q1 = Q1 & " FROM ((((Cuentas INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta) INNER JOIN Cuentas AS Cuentas_1 ON Cuentas.idPadre = Cuentas_1.idCuenta) INNER JOIN Cuentas AS Cuentas_2 ON Cuentas_1.idPadre = Cuentas_2.idCuenta)INNER JOIN Cuentas AS Cuentas_3 ON Cuentas_2.idPadre = Cuentas_3.idCuenta)" & JoinComp
   Q1 = Q1 & " Where Cuentas_3.Nivel = " & Nivel & " AND " & WhereFecha
   Q1 = Q1 & " GROUP BY Cuentas_3.idCuenta, Cuentas_3.Codigo, Cuentas_3.Nivel, Cuentas_3.Descripcion"
   Q1 = Q1 & " UNION"
   
   'suma de cuentas en que este nivel es el tatara-abuelo
   Q1 = Q1 & " SELECT Cuentas_4.idCuenta, Cuentas_4.Codigo, Cuentas_4.Nivel, Cuentas_4.Descripcion, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber"
   Q1 = Q1 & " FROM (((((Cuentas INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta) INNER JOIN Cuentas AS Cuentas_1 ON Cuentas.idPadre = Cuentas_1.idCuenta) INNER JOIN Cuentas AS Cuentas_2 ON Cuentas_1.idPadre = Cuentas_2.idCuenta)INNER JOIN Cuentas AS Cuentas_3 ON Cuentas_2.idPadre = Cuentas_3.idCuenta)INNER JOIN Cuentas AS Cuentas_4 ON Cuentas_3.idPadre = Cuentas_4.idCuenta)" & JoinComp
   Q1 = Q1 & " Where Cuentas_4.Nivel = " & Nivel & " AND " & WhereFecha
   Q1 = Q1 & " GROUP BY Cuentas_4.idCuenta, Cuentas_4.Codigo, Cuentas_4.Nivel, Cuentas_4.Descripcion"
      
   Q1 = Q1 & " ORDER BY Codigo"

   Set Rs = OpenRs(DbMain, Q1)
   
   For j = 0 To MAX_NIVELES
      Total(j).Debe = 0
      Total(j).Haber = 0
      Total(j).Linea = 0
   Next j
   
   i = Grid.FixedRows - 1
   Grid.rows = Grid.FixedRows
   
   CurNiv = 0
   CurCta = ""
   
   Do While Rs.EOF = False
   
      If vFld(Rs("Nivel")) < CurNiv Then    'disminuye el nivel
         For j = CurNiv - 1 To vFld(Rs("Nivel")) Step -1
            Grid.TextMatrix(Total(j).Linea, C_DEBE) = Format(Total(j).Debe, NUMFMT)
            Grid.TextMatrix(Total(j).Linea, C_HABER) = Format(Total(j).Haber, NUMFMT)
            Total(j).Debe = 0
            Total(j).Haber = 0
            Total(j).Linea = 0
         Next j
      End If

      If CurCta <> vFld(Rs("Codigo")) Then
      
         If CurCta <> "" Then
            'ponemos totales de cuenta actual
            Grid.TextMatrix(Total(CurNiv).Linea, C_DEBE) = Format(Total(CurNiv).Debe, NUMFMT)
            Grid.TextMatrix(Total(CurNiv).Linea, C_HABER) = Format(Total(CurNiv).Haber, NUMFMT)
         End If
      
         'actualizamos el nivel
         CurNiv = vFld(Rs("Nivel"))
         
         'agregamos la nueva cuenta
         i = i + 1
         Grid.rows = i + 1
         CurCta = vFld(Rs("Codigo"))
  
         Grid.TextMatrix(i, C_IDCUENTA) = vFld(Rs("idCuenta"))
         Grid.TextMatrix(i, C_NIVEL) = CurNiv
         Grid.TextMatrix(i, C_CODCUENTA) = Format(vFld(Rs("Codigo")), gFmtCodigoCta)
         Grid.TextMatrix(i, C_CUENTA) = String(REP_INDENT * (CurNiv - 1), " ") & FCase(vFld(Rs("Descripcion"), True))
         
         'cambiamos el color de la fila para desacar los niveles (lento)
         'Call FGrForeColor(Grid, i, -1, gRepNivColor(CurNiv))
         
         Total(CurNiv).Debe = 0
         Total(CurNiv).Haber = 0
         Total(CurNiv).Linea = i
         
      End If
   
      'sumamos los totales al nivel actual y a los niveles anteriores
      For j = CurNiv To 1 Step -1
         Total(j).Debe = Total(j).Debe + vFld(Rs("Debe"))
         Total(j).Haber = Total(j).Haber + vFld(Rs("Haber"))
      Next j
            
      Rs.MoveNext
   Loop
      
   'ponemos el total de la última línea
   If CurCta <> "" Then
      'ponemos totales de cuenta actual
      Grid.TextMatrix(Total(CurNiv).Linea, C_DEBE) = Format(Total(CurNiv).Debe, NUMFMT)
      Grid.TextMatrix(Total(CurNiv).Linea, C_HABER) = Format(Total(CurNiv).Haber, NUMFMT)
   End If
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid)
      
   Grid.TopRow = Grid.FixedRows
   Grid.Row = Grid.FixedRows
   Grid.Col = C_CODCUENTA
   Grid.RowSel = Grid.Row
   Grid.ColSel = Grid.Col
   
   Grid.FlxGrid.Redraw = True
    
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
   
End Sub

