VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmResRet3Porc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Retención 3% Préstamo Solidario"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5940
      TabIndex        =   1
      Top             =   5580
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5175
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9128
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmResRet3Porc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_RUT = 0
Const C_TIPODOC = 1
Const C_VALOR = 2
Const C_COLOBLIGATORIA = 3

Const NCOLS = C_COLOBLIGATORIA

Dim lEnImpresion As Boolean

Dim lMes As Integer
Dim lAno As Integer
Dim lPrtObj As Object

Public Function FView(ByVal Mes As Integer, ByVal Ano As Integer)

   lMes = Mes
   lAno = Ano
   
   Me.Show vbModal
End Function
Public Function FPrtRes(PrtObj As Object, ByVal Mes As Integer, ByVal Ano As Integer)
   
   lEnImpresion = True
   Set lPrtObj = PrtObj
   lMes = Mes
   lAno = Ano
   
   Load Me
      
   lEnImpresion = False
   
   DoEvents
   
   Unload Me
   
End Function



Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Call SetUpGrid
   
   Call LoadAll
   
   If lEnImpresion Then
      Call SetUpPrtGrid
      Call gPrtLibros.PrtFlexGrid(lPrtObj)
   End If
   
End Sub

Private Function SetUpGrid()

   Grid.Cols = NCOLS + 1
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_RUT) = 1800
   Grid.ColWidth(C_TIPODOC) = 3000
   Grid.ColWidth(C_VALOR) = 2000
   Grid.ColWidth(C_COLOBLIGATORIA) = 0
 
   Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_TIPODOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
   
   Grid.TextMatrix(0, C_RUT) = "RUT"
   Grid.TextMatrix(0, C_TIPODOC) = "Tipo Doc"
   Grid.TextMatrix(0, C_VALOR) = "Valor"
   
   Call FGrVRows(Grid)
   
End Function
Private Sub SetUpPrtGrid()
   Dim i As Integer
   Dim ColWi(NCOLS) As Integer
   Dim Total(NCOLS) As String
   Dim Titulos(0) As String
   Dim Encabezados(0) As String
      
   Set gPrtLibros.Grid = Grid
   
   Titulos(0) = Caption
   
   gPrtLibros.Titulos = Titulos
      
   gPrtLibros.Encabezados = Encabezados
      
   For i = 0 To Grid.Cols - 1
      ColWi(i) = Grid.ColWidth(i)
   Next i
               
   
   gPrtLibros.ColWi = ColWi
   gPrtLibros.Total = Total
   gPrtLibros.ColObligatoria = C_COLOBLIGATORIA
   gPrtLibros.NTotLines = 0
   

End Sub

Private Function LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Total As Double
   Dim FirstDay As Long, LastDay As Long
   
   If lAno = 0 Then
      Exit Function
   End If
   
   If lMes = 0 Then
      FirstDay = DateSerial(lAno, 1, 1)
      LastDay = DateSerial(lAno, 12, 31)
   Else
      Call FirstLastMonthDay(DateSerial(lAno, lMes, 1), FirstDay, LastDay)
   End If
   
   'registros con retención
   Q1 = "SELECT Rut, TipoDoc, ValRet3Porc, Exento, Afecto "
   Q1 = Q1 & " FROM Documento INNER JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " WHERE Documento.IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND TipoLib = " & LIB_RETEN
   Q1 = Q1 & " AND ValRet3Porc > 0 AND Afecto > 0 "
   Q1 = Q1 & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " ORDER BY IdDoc "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Grid.Redraw = False
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Total = 0
   
   Do While Not Rs.EOF
   
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("RUT")))
      Grid.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(LIB_RETEN, vFld(Rs("TipoDoc")))
      Grid.TextMatrix(i, C_VALOR) = Format(vFld(Rs("ValRet3Porc")), NUMFMT)
      Grid.TextMatrix(i, C_COLOBLIGATORIA) = "."    'sólo para la impresión
      Total = Total + vFld(Rs("ValRet3Porc"))
      
      i = i + 1
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Grid.rows = Grid.rows + 1
   
   Grid.TextMatrix(i, C_TIPODOC) = "Total"
   Grid.TextMatrix(i, C_VALOR) = Format(Total, NUMFMT)
   Call FGrSetRowStyle(Grid, i, "B")
   Grid.TextMatrix(i, C_COLOBLIGATORIA) = "."    'sólo para la impresión
   
   i = i + 1
   Grid.rows = Grid.rows + 1
   Grid.TextMatrix(i, C_COLOBLIGATORIA) = "."    'sólo para la impresión
   
   'registros sin retención
   Q1 = "SELECT Rut, TipoDoc, ValRet3Porc, Exento, Afecto "
   Q1 = Q1 & " FROM Documento INNER JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " WHERE Documento.IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano & " AND TipoLib = " & LIB_RETEN
   Q1 = Q1 & " AND ValRet3Porc > 0 AND Afecto = 0 "
   Q1 = Q1 & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " ORDER BY IdDoc "
   
   Set Rs = OpenRs(DbMain, Q1)
      
   Total = 0
   
   i = i + 1
   
   Do While Not Rs.EOF
   
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_RUT) = FmtCID(vFld(Rs("RUT")))
      Grid.TextMatrix(i, C_TIPODOC) = "Boleta Sin Retención"
      Grid.TextMatrix(i, C_VALOR) = Format(vFld(Rs("ValRet3Porc")), NUMFMT)
      Grid.TextMatrix(i, C_COLOBLIGATORIA) = "."    's´lo para la impresión
      
      Total = Total + vFld(Rs("ValRet3Porc"))
      
      i = i + 1
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   If Total > 0 Then
   
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_TIPODOC) = "Total"
      Grid.TextMatrix(i, C_VALOR) = Format(Total, NUMFMT)
      Call FGrSetRowStyle(Grid, i, "B")
      Grid.TextMatrix(i, C_COLOBLIGATORIA) = "."    'sólo para la impresión
      
   End If
   
   Call FGrVRows(Grid)
   
   Grid.Redraw = True
 
End Function
 
