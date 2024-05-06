VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmEjemploImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejemplo Formato"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   13935
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13815
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
         Left            =   120
         Picture         =   "FrmEjemploImport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   12420
         TabIndex        =   2
         Top             =   180
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5715
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   10081
      _Version        =   393216
      BackColorBkg    =   16777215
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmEjemploImport.frx":0445
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   6720
      Width           =   13575
   End
End
Attribute VB_Name = "FrmEjemploImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CC_TIPO = 0
Const CC_FECHA = 1
Const CC_TOTALCOMP = 2
Const CC_ESTADO = 3
Const CC_GLOSACOMP = 4
Const CC_CUENTA = 5
Const CC_DEBE = 6
Const CC_HABER = 7
Const CC_DESCRIP = 8

Const CC_NCOLS = CC_DESCRIP


Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
   Call FGr2Clip(Grid, Me.Caption)

End Sub

Public Sub FViewComprobantes()

   Call FillComprobantes
   Me.Show vbModal


End Sub

Private Sub FillComprobantes()
   Dim Row As Integer
   
   Grid.Cols = CC_NCOLS + 1
   Grid.rows = 1
   
   
   Call FGrSetup(Grid, True)
   
   Grid.ColWidth(CC_TIPO) = 1000
   Grid.ColWidth(CC_FECHA) = 1000
   Grid.ColWidth(CC_TOTALCOMP) = 1200
   Grid.ColWidth(CC_ESTADO) = 1000
   Grid.ColWidth(CC_GLOSACOMP) = 2800
   Grid.ColWidth(CC_CUENTA) = 1200
   Grid.ColWidth(CC_DEBE) = 1200
   Grid.ColWidth(CC_HABER) = 1200
   Grid.ColWidth(CC_DESCRIP) = 2700
   
   Row = 0
   
   Grid.TextMatrix(Row, CC_TIPO) = "Tipo"
   Grid.TextMatrix(Row, CC_FECHA) = "Fecha"
   Grid.TextMatrix(Row, CC_TOTALCOMP) = "Total Comp"
   Grid.TextMatrix(Row, CC_ESTADO) = "Estado"
   Grid.TextMatrix(Row, CC_GLOSACOMP) = "Glosa Comp"
   Grid.TextMatrix(Row, CC_CUENTA) = "Cód. Cuenta"
   Grid.TextMatrix(Row, CC_DEBE) = "Debe"
   Grid.TextMatrix(Row, CC_HABER) = "Haber"
   Grid.TextMatrix(Row, CC_DESCRIP) = "Descrip"
   
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   Grid.TextMatrix(Row, CC_TIPO) = "Ingreso"
   Grid.TextMatrix(Row, CC_FECHA) = "22/07/20XX"
   Grid.TextMatrix(Row, CC_TOTALCOMP) = "300000"
   Grid.TextMatrix(Row, CC_ESTADO) = "Aprobado"
   Grid.TextMatrix(Row, CC_GLOSACOMP) = "Por Asiento de Apertura"
   Grid.TextMatrix(Row, CC_CUENTA) = "1-01-01-01"
   Grid.TextMatrix(Row, CC_DEBE) = "300000"
   Grid.TextMatrix(Row, CC_HABER) = "0"
   Grid.TextMatrix(Row, CC_DESCRIP) = "Por Asiento de Apertura"
   
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   Grid.TextMatrix(Row, CC_TIPO) = ""
   Grid.TextMatrix(Row, CC_FECHA) = ""
   Grid.TextMatrix(Row, CC_TOTALCOMP) = ""
   Grid.TextMatrix(Row, CC_ESTADO) = ""
   Grid.TextMatrix(Row, CC_GLOSACOMP) = ""
   Grid.TextMatrix(Row, CC_CUENTA) = "2-03-01-01"
   Grid.TextMatrix(Row, CC_DEBE) = "0"
   Grid.TextMatrix(Row, CC_HABER) = "300000"
   Grid.TextMatrix(Row, CC_DESCRIP) = ""
   
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   Grid.TextMatrix(Row, CC_TIPO) = "Egreso"
   Grid.TextMatrix(Row, CC_FECHA) = "25/07/20XX"
   Grid.TextMatrix(Row, CC_TOTALCOMP) = "200000"
   Grid.TextMatrix(Row, CC_ESTADO) = "Pendiente"
   Grid.TextMatrix(Row, CC_GLOSACOMP) = "Por Depósito en Cuenta Corriente"
   Grid.TextMatrix(Row, CC_CUENTA) = "1-01-01-04"
   Grid.TextMatrix(Row, CC_DEBE) = "200000"
   Grid.TextMatrix(Row, CC_HABER) = "0"
   Grid.TextMatrix(Row, CC_DESCRIP) = "Depósito en Banco"
   
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   Grid.TextMatrix(Row, CC_TIPO) = ""
   Grid.TextMatrix(Row, CC_FECHA) = ""
   Grid.TextMatrix(Row, CC_TOTALCOMP) = ""
   Grid.TextMatrix(Row, CC_ESTADO) = ""
   Grid.TextMatrix(Row, CC_GLOSACOMP) = ""
   Grid.TextMatrix(Row, CC_CUENTA) = "1-01-01-01"
   Grid.TextMatrix(Row, CC_DEBE) = "0"
   Grid.TextMatrix(Row, CC_HABER) = "200000"
   Grid.TextMatrix(Row, CC_DESCRIP) = ""
   
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   Grid.TextMatrix(Row, CC_TIPO) = "Traspaso"
   Grid.TextMatrix(Row, CC_FECHA) = "28/07/20XX"
   Grid.TextMatrix(Row, CC_TOTALCOMP) = "119000"
   Grid.TextMatrix(Row, CC_ESTADO) = "Pendiente"
   Grid.TextMatrix(Row, CC_GLOSACOMP) = "Centralización del Libro de Compras"
   Grid.TextMatrix(Row, CC_CUENTA) = "1-01-08-01"
   Grid.TextMatrix(Row, CC_DEBE) = "100000"
   Grid.TextMatrix(Row, CC_HABER) = "0"
   Grid.TextMatrix(Row, CC_DESCRIP) = "Por Libro de Compras mes de Julio 20XX"
   
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   Grid.TextMatrix(Row, CC_TIPO) = ""
   Grid.TextMatrix(Row, CC_FECHA) = ""
   Grid.TextMatrix(Row, CC_TOTALCOMP) = ""
   Grid.TextMatrix(Row, CC_ESTADO) = ""
   Grid.TextMatrix(Row, CC_GLOSACOMP) = ""
   Grid.TextMatrix(Row, CC_CUENTA) = "1-01-09-02"
   Grid.TextMatrix(Row, CC_DEBE) = "19000"
   Grid.TextMatrix(Row, CC_HABER) = "0"
   Grid.TextMatrix(Row, CC_DESCRIP) = ""
   
   Grid.rows = Grid.rows + 1
   Row = Row + 1
   Grid.TextMatrix(Row, CC_TIPO) = ""
   Grid.TextMatrix(Row, CC_FECHA) = ""
   Grid.TextMatrix(Row, CC_TOTALCOMP) = ""
   Grid.TextMatrix(Row, CC_ESTADO) = ""
   Grid.TextMatrix(Row, CC_GLOSACOMP) = ""
   Grid.TextMatrix(Row, CC_CUENTA) = "2-01-06-01"
   Grid.TextMatrix(Row, CC_DEBE) = "0"
   Grid.TextMatrix(Row, CC_HABER) = "119000"
   Grid.TextMatrix(Row, CC_DESCRIP) = ""
   
   Call FGrVRows(Grid, 1)

End Sub

