VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmConfigCompActFijo 
   Caption         =   "Datos de Comprobante Activo Fijo IFRS"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Bt_Comprobante 
      Caption         =   "Crear Comprobante"
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   3495
      Left            =   240
      TabIndex        =   16
      Top             =   3960
      Width           =   7695
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   2895
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
      End
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Bt_Aceptar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   360
      Width           =   1575
   End
   Begin VB.Frame Frm 
      Caption         =   "Datos Comprobante"
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.CheckBox Ck_ValorLibroDespRevalo 
         Caption         =   "Valor Libro después de Revalorización"
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CheckBox Ck_ValorLibroAntRevalor 
         Caption         =   "Valor Libro antes de Revalorización"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CheckBox Ck_ValorDepreciar 
         Caption         =   "Valor a Depreciar"
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox Ck_ValorResidual 
         Caption         =   "Valor Residual"
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox Ck_ValorLibro 
         Caption         =   "Valor Libro"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox Ck_ValorBien 
         Caption         =   "Valor del Bien"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox Ck_ValorInicial 
         Caption         =   "Valor Inicial"
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox Ck_ValorRazonable 
         Caption         =   "Valor Razonable"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox Ck_Haber 
         Caption         =   "Haber"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox Ck_Debe 
         Caption         =   "Debe"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Cuentas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         Picture         =   "FrmConfigCompActFijo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Plan de Cuentas"
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox Tx_Cuenta 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   3
         Top             =   420
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmConfigCompActFijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const C_IDCUENTA = 0
Const C_CUENTA = 1
Const C_DESCRCUENTA = 2
Const C_DEBE = 3
Const C_HABER = 4

Const NCOLS = C_HABER

Dim vIdCuenta As String
Dim vCuenta As String
Dim vCodCuenta As String
Dim vDescCuenta As String
Dim vDebe As Double
Dim vHaber As Double
Dim vValorRazonable As String
Dim vValorInicial As String
Dim vValorDelBien As String
Dim vValorLibro As String
Dim vValorResidual As String
Dim vValorDeprecial As String
Dim vValorLibroAntRevalor As String
Dim vValorLibroDespRevalo As String

Dim i As Integer


Private Sub Bt_Aceptar_Click()

vDebe = 0
vHaber = 0

If Tx_Cuenta = "" Then
 MsgBox1 "Debe seleccionar cuenta.", vbExclamation + vbOKOnly
 
 Exit Sub
End If

If Ck_Debe = 0 And Ck_Haber = 0 Then
 MsgBox1 "Seleccionar Debe o Haber.", vbExclamation + vbOKOnly
 
 Exit Sub
End If


i = i + 1

If Ck_Debe Then

    If Ck_ValorRazonable Then
       vDebe = vValorRazonable
       Ck_ValorRazonable = 0
       Ck_ValorRazonable.Enabled = False
    End If
    If Ck_ValorInicial Then
       vDebe = vDebe + vValorInicial
       Ck_ValorInicial = 0
       Ck_ValorInicial.Enabled = False
    End If
    If Ck_ValorBien Then
       vDebe = vDebe + vValorDelBien
       Ck_ValorBien = 0
       Ck_ValorBien.Enabled = False
    End If
    If Ck_ValorLibro Then
       vDebe = vDebe + vValorLibro
       Ck_ValorLibro = 0
       Ck_ValorLibro.Enabled = False
    End If
    If Ck_ValorResidual Then
       vDebe = vDebe + vValorResidual
       Ck_ValorResidual = 0
       Ck_ValorResidual.Enabled = False
    End If
    If Ck_ValorDepreciar Then
       vDebe = vDebe + vValorDeprecial
       Ck_ValorDepreciar = 0
       Ck_ValorDepreciar.Enabled = False
    End If
    If Ck_ValorLibroAntRevalor Then
       vDebe = vDebe + vValorLibroAntRevalor
       Ck_ValorLibroAntRevalor = 0
       Ck_ValorLibroAntRevalor.Enabled = False
    End If
    If Ck_ValorLibroDespRevalo Then
       vDebe = vDebe + vValorLibroDespRevalo
       Ck_ValorLibroDespRevalo = 0
       Ck_ValorLibroDespRevalo.Enabled = False
    End If

      
    Grid.TextMatrix(i, C_DEBE) = Format(vDebe, NUMFMT)
    
      
ElseIf Ck_Haber Then

    If Ck_ValorRazonable Then
      vHaber = vValorRazonable
      Ck_ValorRazonable = 0
      Ck_ValorRazonable.Enabled = False
    End If
    If Ck_ValorInicial Then
      vHaber = vHaber + vValorInicial
      Ck_ValorInicial = 0
      Ck_ValorInicial.Enabled = False
    End If
    If Ck_ValorBien Then
      vHaber = vHaber + vValorDelBien
      Ck_ValorBien = 0
      Ck_ValorBien.Enabled = False
    End If
    If Ck_ValorLibro Then
       vHaber = vHaber + vValorLibro
       Ck_ValorLibro = 0
       Ck_ValorLibro.Enabled = False
    End If
    If Ck_ValorResidual Then
       vHaber = vHaber + vValorResidual
       Ck_ValorResidual = 0
       Ck_ValorResidual.Enabled = False
    End If
    If Ck_ValorDepreciar Then
       vHaber = vHaber + vValorDeprecial
       Ck_ValorDepreciar = 0
       Ck_ValorDepreciar.Enabled = False
    End If
    If Ck_ValorLibroAntRevalor Then
       vHaber = vHaber + vValorLibroAntRevalor
       Ck_ValorLibroAntRevalor = 0
       Ck_ValorLibroAntRevalor.Enabled = False
    End If
    If Ck_ValorLibroDespRevalo Then
       vHaber = vHaber + vValorLibroDespRevalo
       Ck_ValorLibroDespRevalo = 0
       Ck_ValorLibroDespRevalo.Enabled = False
    End If

   
    Grid.TextMatrix(i, C_HABER) = Format(vHaber, NUMFMT)
    
End If
      
 
    Grid.TextMatrix(i, C_CUENTA) = vCodCuenta
    Grid.TextMatrix(i, C_IDCUENTA) = vIdCuenta
    Grid.TextMatrix(i, C_DESCRCUENTA) = vDescCuenta

   Grid.rows = Grid.rows + 1
   
   Tx_Cuenta = ""
   Ck_Debe = 0
   Ck_Haber = 0

End Sub

Private Sub Bt_Cancelar_Click()
Unload Me
End Sub

Private Sub Bt_Comprobante_Click()
Dim FrmComp As FrmComprobante

Set FrmComp = New FrmComprobante
  
For i = Grid.FixedRows To Grid.rows - 2
     
      Call FrmComp.FNewCompActivo(i, Val(Replace(Replace(Grid.TextMatrix(i, C_DEBE), ".", ""), ",", "")), Val(Replace(Replace(Grid.TextMatrix(i, C_HABER), ".", ""), ",", "")), Grid.TextMatrix(i, C_IDCUENTA), Grid.TextMatrix(i, C_CUENTA), Grid.TextMatrix(i, C_DESCRCUENTA))
Next i

FrmComp.Show vbModal
      Set FrmComp = Nothing


End Sub

Private Sub Bt_Cuentas_Click()
Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   Dim ClasCta As Integer
   
   Set Frm = New FrmPlanCuentas

   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre, False) = vbOK Then
      ClasCta = GetClasCuenta(IdCuenta)
            
      vCodCuenta = Codigo
      Tx_Cuenta = Descrip
      vIdCuenta = IdCuenta
      vDescCuenta = Descrip
      'lIdCtaPatrimonio = IdCuenta
   End If
   
   Set Frm = Nothing


End Sub

Public Sub FSelect(v_IdCuenta As String, v_Cuenta As String, v_DescCuenta As String, v_Debe As Boolean, v_Haber As Boolean, v_ValorRazonable As Boolean, v_ValorInicial As Boolean, v_ValorDelBien As Boolean, v_ValorLibro As Boolean, v_ValorResidual As Boolean, v_ValorDeprecial As Boolean, v_ValorLibroAntRevalor As Boolean, v_ValorLibroDespRevalo As Boolean)
  
    v_Cuenta = vCuenta
    v_IdCuenta = vIdCuenta
    v_DescCuenta = vDescCuenta
    v_Debe = vDebe
    v_Haber = vHaber
    v_ValorRazonable = vValorRazonable
    v_ValorInicial = vValorInicial
    v_ValorDelBien = vValorDelBien
    v_ValorLibro = vValorLibro
    v_ValorResidual = vValorResidual
    v_ValorDeprecial = vValorDeprecial
    v_ValorLibroAntRevalor = vValorLibroAntRevalor
    v_ValorLibroDespRevalo = v_ValorLibroDespRevalo
       
End Sub

Public Sub FSelect2(ByVal v_IdCuenta As String, ByVal v_Cuenta As String, ByVal v_DescCuenta As String, ByVal v_ValorRazonable As String, ByVal v_ValorInicial As String, ByVal v_ValorDelBien As String, ByVal v_ValorLibro As String, ByVal v_ValorResidual As String, ByVal v_ValorDeprecial As String, ByVal v_ValorLibroAntRevalor As String, ByVal v_ValorLibroDespRevalo As String)
  
vValorRazonable = v_ValorRazonable
vValorInicial = v_ValorInicial
vValorDelBien = v_ValorDelBien
vValorLibro = v_ValorLibro
vValorResidual = v_ValorResidual
vValorDeprecial = v_ValorDeprecial
vValorLibroAntRevalor = v_ValorLibroAntRevalor
vValorLibroDespRevalo = v_ValorLibroDespRevalo
       
       Me.Show vbModal
End Sub


Public Sub FView()
   
  Me.Show vbModal
   
End Sub

Private Sub Ck_Debe_Click()
If Ck_Debe = 1 Then
    Ck_Haber = 0
End If
End Sub

Private Sub Ck_Haber_Click()
If Ck_Haber = 1 Then
    Ck_Debe = 0
End If
End Sub


Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
    
   Grid.ColWidth(C_DEBE) = 1300
   Grid.ColWidth(C_HABER) = 1300
   Grid.ColWidth(C_CUENTA) = 1300
   Grid.ColWidth(C_IDCUENTA) = 0
   Grid.ColWidth(C_DESCRCUENTA) = 1300
   
      
   Grid.ColAlignment(C_DEBE) = flexAlignRightCenter
   Grid.ColAlignment(C_HABER) = flexAlignRightCenter
   Grid.ColAlignment(C_CUENTA) = flexAlignLeftCenter
   Grid.ColAlignment(C_IDCUENTA) = flexAlignRightCenter
   Grid.ColAlignment(C_DESCRCUENTA) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_DEBE) = "Valor Debe"
   Grid.TextMatrix(0, C_HABER) = "Valor Haber"
   Grid.TextMatrix(0, C_CUENTA) = "Cod.Cuenta"
   Grid.TextMatrix(0, C_IDCUENTA) = "Id Cuenta"
   Grid.TextMatrix(0, C_DESCRCUENTA) = "Descr. Cuenta"
   
   'GridTot.Cols = Grid.Cols

   Call FGrSetup(Grid)

   'Call FGrVRows(Grid)

End Sub

Private Sub Form_Load()
SetUpGrid
End Sub


