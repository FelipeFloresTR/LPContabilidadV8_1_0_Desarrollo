VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPrtEmpresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Empresas"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   Icon            =   "FrmPrtEmpresas.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   11205
      Begin VB.CommandButton bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   330
         Left            =   9720
         TabIndex        =   5
         Top             =   180
         Width           =   1320
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
         Picture         =   "FrmPrtEmpresas.frx":000C
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
         Left            =   90
         Picture         =   "FrmPrtEmpresas.frx":04C6
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
         Left            =   990
         Picture         =   "FrmPrtEmpresas.frx":096D
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6090
      Left            =   45
      TabIndex        =   0
      Top             =   765
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   10742
      _Version        =   393216
      Rows            =   30
      Cols            =   11
      FixedCols       =   0
   End
End
Attribute VB_Name = "FrmPrtEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const C_RUT = 0
Const C_RSOCIAL = 1
Const C_ESTADO = 2         'no activa
Const C_CALLE = 3
Const C_TELEFONOS = 4
Const C_FAX = 5
Const C_CIUDAD = 6
Const C_CODACTECO = 7
Const C_GIRO = 8
Const C_EMAIL = 9
Const C_WEB = 10
Const C_CONTADOR = 11

Const NCOLS = C_CONTADOR

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Caption)
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
   
   Printer.Orientation = PrtOrient
   MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()
   Call SetUpGrid
   Call LoadAll
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Row As Integer
   Dim IdEmpresa As Long
   
#If DATACON = 1 Then
   Dim PathEmp As String
   Dim BDatos As String

   Q1 = "SELECT Max(Ano) as Anos, Rut, Estado, NombreCorto "
   Q1 = Q1 & " FROM Empresas "
   Q1 = Q1 & " LEFT JOIN EmpresasAno ON EmpresasAno.idEmpresa = Empresas.idEmpresa"
   Q1 = Q1 & " GROUP BY Rut, Estado, NombreCorto "
   Q1 = Q1 & " ORDER BY Rut "
   Set Rs = OpenRs(DbMain, Q1)
     
   Grid.Rows = Grid.FixedRows
   Row = Grid.Rows
   
   Do While Rs.EOF = False
      PathEmp = gDbPath & "\Empresas\" & vFld(Rs("Anos"))
      BDatos = PathEmp & "\" & vFld(Rs("Rut")) & ".mdb"
      
      Grid.Rows = Grid.Rows + 1
      
      If vFld(Rs("NombreCorto")) = "Mersud" Then
         MsgBeep vbExclamation
      End If
         
      If ExistFile(BDatos) Then
         Call FillEmpresa(Row, vFld(Rs("Rut")), BDatos)
      
      Else
         Grid.TextMatrix(Row, C_RUT) = FmtCID(vFld(Rs("Rut")))
         Grid.TextMatrix(Row, C_RSOCIAL) = vFld(Rs("NombreCorto"))
      
      End If
      
      Grid.TextMatrix(Row, C_ESTADO) = IIf(vFld(Rs("Estado")) = 0, "Si", "No")
      
      Row = Row + 1
      Rs.MoveNext
            
   Loop
   
#Else

   Q1 = "SELECT Empresas.IdEmpresa, Empresas.Rut, Empresas.Estado, Empresas.NombreCorto, "
   Q1 = Q1 & " Ano, RazonSocial, ApPaterno, ApMaterno, Nombre, Calle, Numero,"
   Q1 = Q1 & " Dpto, Telefonos, Fax, Ciudad, Giro, email, Web, Contador, CodActEconom "
   Q1 = Q1 & " FROM Empresas "
   Q1 = Q1 & " LEFT JOIN Empresa ON Empresa.id = Empresas.idEmpresa"
'   Q1 = Q1 & " ORDER BY Empresas.IdEmpresa, Ano desc "
   Q1 = Q1 & " ORDER BY Empresas.Rut, Ano desc "
   Set Rs = OpenRs(DbMain, Q1)

   Grid.Rows = Grid.FixedRows
   Row = Grid.Rows
   IdEmpresa = 0
   
   Do While Rs.EOF = False
      
      If IdEmpresa <> vFld(Rs("IdEmpresa")) Then
         IdEmpresa = vFld(Rs("IdEmpresa"))
         Grid.Rows = Grid.Rows + 1
         
         Grid.TextMatrix(Row, C_RUT) = FmtCID(vFld(Rs("Rut")))
         Grid.TextMatrix(Row, C_ESTADO) = IIf(vFld(Rs("Estado")) = 0, "Si", "No")
         
         If vFld(Rs("Razonsocial")) <> "" Then
            If Trim(vFld(Rs("Nombre"))) <> "" Then
               Grid.TextMatrix(Row, C_RSOCIAL) = vFld(Rs("Razonsocial")) & " " & vFld(Rs("ApMaterno")) & " " & vFld(Rs("Nombre"))
            Else
               Grid.TextMatrix(Row, C_RSOCIAL) = vFld(Rs("Razonsocial"))
            End If
            
            Grid.TextMatrix(Row, C_CALLE) = vFld(Rs("Calle")) & " " & vFld(Rs("Numero")) & " " & vFld(Rs("Dpto"))
            Grid.TextMatrix(Row, C_TELEFONOS) = vFld(Rs("Telefonos"))
            Grid.TextMatrix(Row, C_FAX) = vFld(Rs("Fax"))
            Grid.TextMatrix(Row, C_CIUDAD) = vFld(Rs("Ciudad"))
            Grid.TextMatrix(Row, C_CODACTECO) = vFld(Rs("CodActEconom"))
            Grid.TextMatrix(Row, C_GIRO) = vFld(Rs("Giro"))
            Grid.TextMatrix(Row, C_EMAIL) = vFld(Rs("email"))
            Grid.TextMatrix(Row, C_WEB) = vFld(Rs("WEB"))
            Grid.TextMatrix(Row, C_CONTADOR) = vFld(Rs("Contador"))
         
         Else
            Grid.TextMatrix(Row, C_RSOCIAL) = vFld(Rs("NombreCorto"))
         End If
         Row = Row + 1
      End If
      
      Rs.MoveNext
   Loop
        
#End If
   Call CloseRs(Rs)
   Call FGrVRows(Grid)
   
End Sub
Private Sub SetUpGrid()
   Dim Col As Integer
   
   Grid.Cols = NCOLS + 1
   
   Call FGrSetup(Grid, True)
   
   Grid.ColWidth(C_RUT) = 1100
   Grid.ColWidth(C_RSOCIAL) = 2000
   Grid.ColWidth(C_ESTADO) = 600
   Grid.ColWidth(C_CALLE) = 1500
   Grid.ColWidth(C_TELEFONOS) = 1500
   Grid.ColWidth(C_FAX) = 1500
   Grid.ColWidth(C_CIUDAD) = 1500
   Grid.ColWidth(C_CODACTECO) = 1200
   Grid.ColWidth(C_GIRO) = 1700
   Grid.ColWidth(C_EMAIL) = 1500
   Grid.ColWidth(C_WEB) = 0
   Grid.ColWidth(C_CONTADOR) = 1500
   
   
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_ESTADO) = flexAlignCenterCenter
   
   Grid.TextMatrix(0, C_RUT) = "Rut"
   Grid.TextMatrix(0, C_RSOCIAL) = "Razón Social"
   Grid.TextMatrix(0, C_ESTADO) = "Activa"
   Grid.TextMatrix(0, C_CALLE) = "Dirección"
   Grid.TextMatrix(0, C_TELEFONOS) = "Telefonos"
   Grid.TextMatrix(0, C_FAX) = "Fax"
   Grid.TextMatrix(0, C_CIUDAD) = "Ciudad"
   Grid.TextMatrix(0, C_GIRO) = "Giro"
   Grid.TextMatrix(0, C_EMAIL) = "Email"
   Grid.TextMatrix(0, C_WEB) = ""
   Grid.TextMatrix(0, C_CONTADOR) = "Contador"
   Grid.TextMatrix(0, C_CODACTECO) = "Cód. act. ecón."
   
End Sub


Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(0) As String
   Dim Encabezados(0) As String
   
   Me.MousePointer = vbHourglass
   
   Set gPrtReportes.Grid = Grid
      
   Titulos(0) = "LISTADO DE EMPRESAS"
   gPrtReportes.Titulos = Titulos
   
   gPrtReportes.Encabezados = Encabezados
   
   gPrtReportes.GrFontName = Grid.FontName
   gPrtReportes.GrFontSize = Grid.FontSize

   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
   
   ColWi(C_TELEFONOS) = ColWi(C_TELEFONOS) - 200
   ColWi(C_FAX) = ColWi(C_FAX) - 200
   
   gPrtReportes.ColWi = ColWi
   gPrtReportes.ColObligatoria = C_RUT
   gPrtReportes.NTotLines = 0
   
End Sub
Private Sub FillEmpresa(Row As Integer, Rut As String, BDatos As String)
#If DATACON = 1 Then
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "PWD=" & PASSW_PREFIX & Rut & ";"
   If Not LinkMdbTable(DbMain, BDatos, "Empresa", "Empresas" & Rut, , , Q1) Then
      MsgBox1 "Archivo " & BDatos & vbCrLf & vbCrLf & " no encontrado o no se pudo acceder.", vbExclamation
      Exit Sub
   End If

   Q1 = "SELECT Rut, RazonSocial, ApPaterno, ApMaterno, Nombre, Calle, Numero,"
   Q1 = Q1 & " Dpto, Telefonos, Fax, Ciudad, Giro, email, Web, Contador, CodActEconom "
   Q1 = Q1 & " FROM Empresas" & Rut
   If gAppCode.Demo Then
      Q1 = Q1 & " WHERE RUT IN ('1','2','3')"
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Grid.TextMatrix(Row, C_RUT) = FmtCID(Rut)
      If Trim(vFld(Rs("Nombre"))) <> "" Then
         Grid.TextMatrix(Row, C_RSOCIAL) = vFld(Rs("Razonsocial"), True) & " " & vFld(Rs("ApMaterno"), True) & " " & vFld(Rs("Nombre"), True)
      Else
         Grid.TextMatrix(Row, C_RSOCIAL) = vFld(Rs("Razonsocial"), True)
      End If
      
      Grid.TextMatrix(Row, C_CALLE) = vFld(Rs("Calle"), True) & " " & vFld(Rs("Numero"), True) & " " & vFld(Rs("Dpto"), True)
      Grid.TextMatrix(Row, C_TELEFONOS) = vFld(Rs("Telefonos"), True)
      Grid.TextMatrix(Row, C_FAX) = vFld(Rs("Fax"), True)
      Grid.TextMatrix(Row, C_CIUDAD) = vFld(Rs("Ciudad"), True)
      Grid.TextMatrix(Row, C_CODACTECO) = vFld(Rs("CodActEconom"), True)
      Grid.TextMatrix(Row, C_GIRO) = vFld(Rs("Giro"), True)
      Grid.TextMatrix(Row, C_EMAIL) = vFld(Rs("email"), True)
      Grid.TextMatrix(Row, C_WEB) = vFld(Rs("WEB"), True)
      Grid.TextMatrix(Row, C_CONTADOR) = vFld(Rs("Contador"), True)
         
   End If
   Call CloseRs(Rs)
   Call UnLinkTable(DbMain, "Empresas" & Rut)
   
#End If
End Sub

Private Sub Form_Resize()

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   Grid.Height = Me.Height - Grid.Top - 500
   Grid.Width = Me.Width - 230
   
   Call FGrVRows(Grid)

End Sub
