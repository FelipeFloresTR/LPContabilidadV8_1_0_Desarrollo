VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmHelpCred33bis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda para Crédito Activo Fijo (Art. 33 Bis Ley de renta)"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   8
      Cols            =   3
   End
End
Attribute VB_Name = "FrmHelpCred33bis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_FECHA = 0
Const C_VENTAS = 1
Const C_TASA = 2
Const C_TOPE = 3

Const NCOLS = C_TOPE

Private Sub SetUpGrid()
   Dim i As Integer

   Grid.Cols = NCOLS + 1
   Grid.rows = 8
   
   Call FGrSetup(Grid)
   
   For i = 0 To Grid.Cols - 1
      Grid.ColWidth(i) = 3000
      Grid.ColAlignment(i) = flexAlignLeftCenter
   Next i
   Grid.ColWidth(C_VENTAS) = 3500
   Grid.ColWidth(C_TASA) = 3170
   Grid.ColWidth(C_TOPE) = 700
   
   For i = 0 To Grid.rows - 1
      Grid.RowHeight(i) = Grid.RowHeight(i) * 2
   Next i
   
   Grid.TextMatrix(0, C_FECHA) = "Fecha de Adquisición"
   Grid.TextMatrix(0, C_VENTAS) = "Promedio de Ventas"
   Grid.TextMatrix(0, C_TASA) = "Tasa de Crédito"
   Grid.TextMatrix(0, C_TOPE) = "Tope"
   
   
End Sub

Private Sub FillGrid()
   Dim i As Integer
   
   i = 1
   Grid.TextMatrix(i, C_FECHA) = "Hasta 30/09/2014"
   Grid.TextMatrix(i, C_VENTAS) = "No importa"
   Grid.TextMatrix(i, C_TASA) = "4%"
   
   i = i + 1
   Grid.TextMatrix(i, C_FECHA) = "Desde 01/10/2014 hasta 30/09/2015"
   Grid.TextMatrix(i, C_VENTAS) = "Menor o igual a 25.000 UF"
   Grid.TextMatrix(i, C_TASA) = "8%"
   
   i = i + 1
   Grid.TextMatrix(i, C_FECHA) = "Desde 01/10/2014 hasta 30/09/2015"
   Grid.TextMatrix(i, C_VENTAS) = "Superior a 25.000 y menor o igual a 100.000 UF"
   Grid.TextMatrix(i, C_TASA) = "8% * (100.000 - Ventas Anuales) / 75.000"
   Grid.TextMatrix(i, C_TOPE) = "Mín. 4%"
   
   i = i + 1
   Grid.TextMatrix(i, C_FECHA) = "Desde 01/10/2014 hasta 30/09/2015"
   Grid.TextMatrix(i, C_VENTAS) = "Superior a 100.000 UF"
   Grid.TextMatrix(i, C_TASA) = "4%"
   
   i = i + 1
   Grid.TextMatrix(i, C_FECHA) = "A contar del 01/10/2015"
   Grid.TextMatrix(i, C_VENTAS) = "Menor o igual a 25.000 UF"
   Grid.TextMatrix(i, C_TASA) = "6%"
   
   i = i + 1
   Grid.TextMatrix(i, C_FECHA) = "A contar del 01/10/2015"
   Grid.TextMatrix(i, C_VENTAS) = "Superior a 25.000 y menor o igual a 100.000 UF"
   Grid.TextMatrix(i, C_TASA) = "6% * (100.000 - Ventas Anuales) / 75.000"
   Grid.TextMatrix(i, C_TOPE) = "Mín. 4%"
   
   i = i + 1
   Grid.TextMatrix(i, C_FECHA) = "A contar del 01/10/2015"
   Grid.TextMatrix(i, C_VENTAS) = "Superior a 100.000 UF"
   Grid.TextMatrix(i, C_TASA) = "4%"
  
End Sub

Private Sub Form_Load()
   Call SetUpGrid
      Call FillGrid
End Sub
