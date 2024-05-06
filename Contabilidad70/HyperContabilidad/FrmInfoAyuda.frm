VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmInfoAyuda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   1440
      TabIndex        =   7
      Top             =   8160
      Width           =   7095
      Begin VB.Label Lbl_Nota 
         Caption         =   "Nota"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Captura Libro Ventas..."
      Height          =   555
      Left            =   8880
      TabIndex        =   6
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   8880
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4755
      Left            =   1440
      TabIndex        =   1
      Top             =   780
      Width           =   7180
      _ExtentX        =   12674
      _ExtentY        =   8387
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   420
      Picture         =   "FrmInfoAyuda.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   480
      Width           =   675
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1845
      Left            =   1440
      TabIndex        =   4
      Top             =   6240
      Width           =   7180
      _ExtentX        =   12674
      _ExtentY        =   3254
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estos documentos deben ser capturados según formato que tiene el sistema en el Libro de Ventas"
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   5940
      Width           =   7180
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "El sistema LP Contabilidad efectuará captura del Registro de Venta de los siguientes documentos: "
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   7180
   End
End
Attribute VB_Name = "FrmInfoAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CODIGO = 0
Const C_NOMBRE = 1

Const NCOLS = C_NOMBRE + 1

Dim lTit As String

Public Function FView(ByVal Tit As String)

   lTit = Tit
   Me.Show vbModal
   
End Function

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Dim Frm As FrmFmtImpEnt
   
   Set Frm = New FrmFmtImpEnt
   Call Frm.FViewLibVentas
   
   Set Frm = Nothing
End Sub

Private Sub Form_Load()

   Me.Caption = lTit

   Call LoadInfo
   
End Sub

Private Function LoadInfo()

   Grid.Cols = NCOLS
   Grid.rows = Grid.FixedRows
   
   Call FGrSetup(Grid)
   
   Grid.ColWidth(C_NOMBRE) = 5800
   Grid.TextMatrix(0, C_CODIGO) = "Código"
   Grid.TextMatrix(0, C_NOMBRE) = "Nombre Documento"
   
   Call Grid.AddItem("29" & vbTab & "Factura Inicio")
   Call Grid.AddItem("30" & vbTab & "Factura  ")
   Call Grid.AddItem("32" & vbTab & "Factura de Ventas y Servicios no Afectos o Exentos")
   Call Grid.AddItem("33" & vbTab & "Factura Electrónica")
   Call Grid.AddItem("34" & vbTab & "Factura de Ventas y Servicios no Afectos o Exentos Electrónica")
   Call Grid.AddItem("35" & vbTab & "Boleta de Ventas y Servicios (Afecta)")
   Call Grid.AddItem("38" & vbTab & "Boleta de Ventas y Servicios no afectos o exentos de IVA")
   Call Grid.AddItem("39" & vbTab & "Boleta de Ventas y Servicios (Afecta) Electrónica")
   Call Grid.AddItem("41" & vbTab & "Boleta de Ventas y Servicios no afectos o exentos de IVA Electrónica")
   Call Grid.AddItem("45" & vbTab & "Factura de Compra")
   Call Grid.AddItem("46" & vbTab & "Factura de Compra Electrónica")
   Call Grid.AddItem("55" & vbTab & "Nota de Debito")
   Call Grid.AddItem("56" & vbTab & "Nota de Debito Electrónica")
   Call Grid.AddItem("60" & vbTab & "Nota de Crédito")
   Call Grid.AddItem("61" & vbTab & "Nota de Crédito Electrónica")
   Call Grid.AddItem("101" & vbTab & "Factura de Exportación")
   Call Grid.AddItem("104" & vbTab & "Nota debito Exportación")
   Call Grid.AddItem("106" & vbTab & "Nota crédito Exportación")
   Call Grid.AddItem("110" & vbTab & "Factura de Exportación Electrónica")
   Call Grid.AddItem("111" & vbTab & "Nota debito Exportación Electrónica")
   Call Grid.AddItem("112" & vbTab & "Nota crédito Exportación Electrónica")
   
   Grid2.Cols = NCOLS
   Grid2.rows = Grid2.FixedRows
   
   Call FGrSetup(Grid2)
   
   Grid2.ColWidth(C_NOMBRE) = 6100
   Grid2.TextMatrix(0, C_CODIGO) = "Código"
   Grid2.TextMatrix(0, C_NOMBRE) = "Nombre Documento"
   
   Call Grid2.AddItem("LFV" & vbTab & "Liquidación Factura")
   Call Grid2.AddItem("DVB" & vbTab & "Devolución Venta con Boleta")
   Call Grid2.AddItem("VEM" & vbTab & "Venta Menor")
   Call Grid2.AddItem("VSD" & vbTab & "Venta sin documentos ")
   Call Grid2.AddItem("MRG" & vbTab & "Máquina Registradora")
   Call Grid2.AddItem("VPE" & vbTab & "Vale Pago electrónico")
      
   Me.Lbl_Nota.Caption = "Nota: El sistema LP Conta captura solo hasta 3500 Registros ya sea para Compras o Ventas " & vbNewLine & "           Si desea capturar una cantidad mayor deber utilizar la Versión LP Conta SQL"
      
End Function
