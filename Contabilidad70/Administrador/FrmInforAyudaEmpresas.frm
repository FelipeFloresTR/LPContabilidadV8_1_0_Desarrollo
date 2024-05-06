VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmInforAyudaEmpresas 
   Caption         =   "Informacion de Ayuda Importador de Empresas"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   360
      TabIndex        =   5
      Top             =   5280
      Width           =   10335
      Begin VB.Label Label1 
         Caption         =   "* Indica los campos que deben tener un valor válido (distinto de blanco)"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   1080
         TabIndex        =   10
         Top             =   1620
         Width           =   8955
      End
      Begin VB.Label Lb_NotaImp 
         Caption         =   $"FrmInforAyudaEmpresas.frx":0000
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   8955
      End
      Begin VB.Label Label3 
         Caption         =   "NOTAS:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "El formato del archivo es posicional, por lo que se deben incluir  TODOS los campos, aunque vayan en blanco. "
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   7
         Top             =   1200
         Width           =   8955
      End
      Begin VB.Label Lb_OtrosDocs 
         Caption         =   "Las Empresas importadas quedarán en estado ACTIVO"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   2040
         Width           =   8955
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   10275
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
         Picture         =   "FrmInforAyudaEmpresas.frx":0118
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Copiar Excel"
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   8760
         TabIndex        =   1
         Top             =   180
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4095
      Left            =   360
      TabIndex        =   3
      Top             =   1140
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7223
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label Label2 
      Caption         =   "Columnas o campos del archivo:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   2475
   End
End
Attribute VB_Name = "FrmInforAyudaEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CAMPO = 0
Const C_FORMATO = 1

'Dim lFmtArray() As gImpEmpresas
Dim lFmtCaption As String

Dim lTit As String

Public Function FView(ByVal Tit As String)

   lTit = Tit
   Me.Show vbModal
   
End Function

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub


Private Sub Bt_Close_Click()
Unload Me
End Sub

Private Sub Bt_CopyExcel_Click()
Call FGr2ClipTrasp(Grid, Me.Caption)
End Sub

Private Sub Form_Load()

   Call SetUpGrid
   Call FillComprobantes
   
   Call LoadGrid
   
End Sub

Private Sub SetUpGrid()

   Call FGrSetup(Grid)

   Grid.ColWidth(C_CAMPO) = 2400
   Grid.ColWidth(C_FORMATO) = 8400
   
   Grid.ColAlignment(C_CAMPO) = flexAlignLeftCenter
   Grid.ColAlignment(C_FORMATO) = flexAlignLeftCenter
   
   
   Grid.TextMatrix(0, C_CAMPO) = "Campo de Información"
   Grid.TextMatrix(0, C_FORMATO) = "Formato"
   
End Sub

Private Sub LoadGrid()
   Dim i As Integer
   Dim j As Integer

   Grid.rows = Grid.FixedRows
   i = Grid.rows - 1
   
   For j = 0 To UBound(gImpEmpresas)
      Grid.rows = Grid.rows + 1
      i = i + 1
      Grid.TextMatrix(i, C_CAMPO) = gImpEmpresas(j).Campo
      Grid.TextMatrix(i, C_FORMATO) = gImpEmpresas(j).Formato
   Next j
   
   Call FGrVRows(Grid)
End Sub

Private Sub FillComprobantes()
   Dim i As Integer
   
   ' ReDim gImpEmpresas(0)

   i = 0
   
   ReDim gImpEmpresas(i)
   gImpEmpresas(i).Campo = "Rut*"
   gImpEmpresas(i).Formato = "Si es RUT: Con o sin punto y digito verificador Ejemplo: 11.111.111-1 / 11111111-1"
   
   i = i + 1
   ReDim Preserve gImpEmpresas(i)
   gImpEmpresas(i).Campo = "Nombre Corto o Razon Social*"
   gImpEmpresas(i).Formato = "Nombre Corto de la entidad y sin blancos, largo 15. Campo no se puede repetir con otra Entidad"

End Sub

