VERSION 5.00
Begin VB.Form FrmPrtFoliacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Foliar Hojas para Timbraje"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   Icon            =   "FrmPrtFoliacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Orientación Papel"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Index           =   2
      Left            =   1365
      TabIndex        =   24
      Top             =   2820
      Width           =   7035
      Begin VB.OptionButton Op_Orientacion 
         Caption         =   "Horizontal"
         Height          =   255
         Index           =   2
         Left            =   4590
         TabIndex        =   4
         Top             =   465
         Width           =   1095
      End
      Begin VB.OptionButton Op_Orientacion 
         Caption         =   "Vertical"
         Height          =   255
         Index           =   1
         Left            =   2070
         TabIndex        =   3
         Top             =   420
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   1
         Left            =   4005
         Picture         =   "FrmPrtFoliacion.frx":000C
         Top             =   315
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   0
         Left            =   1485
         Picture         =   "FrmPrtFoliacion.frx":0316
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.CommandButton Bt_Print 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   8535
      Picture         =   "FrmPrtFoliacion.frx":0620
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1935
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   8535
      TabIndex        =   7
      Top             =   435
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imprimir"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Index           =   1
      Left            =   1365
      TabIndex        =   15
      Top             =   1860
      Width           =   7035
      Begin VB.CheckBox Ch_Prt 
         Caption         =   "Actualizar último impreso"
         Height          =   375
         Left            =   4860
         TabIndex        =   2
         Top             =   315
         Value           =   1  'Checked
         Width           =   2040
      End
      Begin VB.TextBox Tx_Hasta 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3645
         TabIndex        =   1
         Top             =   315
         Width           =   1095
      End
      Begin VB.TextBox Tx_Desde 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         TabIndex        =   0
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folio hasta:"
         Height          =   195
         Index           =   4
         Left            =   2745
         TabIndex        =   17
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folio desde:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Folios"
      ForeColor       =   &H00FF0000&
      Height          =   1395
      Index           =   0
      Left            =   1365
      TabIndex        =   8
      Top             =   360
      Width           =   7035
      Begin VB.CommandButton Bt_Mod 
         Caption         =   "Modificar..."
         Height          =   735
         Left            =   5760
         Picture         =   "FrmPrtFoliacion.frx":0C4F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Tx_UltUsado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Tx_FUltUsado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Tx_FUltTimbrado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Tx_FUltImpreso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Tx_UltTimbrado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Tx_UltImpreso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha último usado:"
         Height          =   195
         Index           =   7
         Left            =   2760
         TabIndex        =   20
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha último timbrado:"
         Height          =   195
         Index           =   6
         Left            =   2760
         TabIndex        =   19
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha último impreso:"
         Height          =   195
         Index           =   5
         Left            =   2760
         TabIndex        =   18
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Último usado:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Último timbrado:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Último impreso:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   360
      Picture         =   "FrmPrtFoliacion.frx":1222
      Top             =   360
      Width           =   750
   End
End
Attribute VB_Name = "FrmPrtFoliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lIdxOrientacion As Integer
Dim lRc As Integer


Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_Mod_Click()
   Dim Frm As FrmFoliacion
   
   Set Frm = New FrmFoliacion
   If Frm.FEdit = vbOK Then
      Call LoadAll
   End If
   
   Set Frm = Nothing
   
End Sub

Private Sub bt_OK_Click()
   lRc = vbOK
   Unload Me
End Sub

Private Sub Bt_Print_Click()
   Dim Q1 As String
   Dim F1 As Long
   
   If Valida() = False Then
      Exit Sub
   End If
   
   'Imprimir, luego guardar fecha y nº último folio impreso
   Call PrtHojasFoliadas
   
   If Ch_Prt Then
      If gFoliacion.Estado = EF_NOEXISTE Then
         Q1 = "INSERT INTO Timbraje (idEmpresa) VALUES (" & gEmpresa.id & ")"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      F1 = CLng(Int(Now))
      
      Q1 = "UPDATE Timbraje SET UltImpreso=" & vFmt(Tx_Hasta)
      Q1 = Q1 & ", FUltImpreso=" & F1
      Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id
      Call ExecSQL(DbMain, Q1)
      
      gFoliacion.UltImpreso = vFmt(Tx_Hasta)
      gFoliacion.FUltImpreso = F1
      gFoliacion.Estado = EF_EXISTE
            
   End If
   
   Unload Me
   
End Sub

Private Function Valida() As Boolean
   Valida = False
   
   If vFmt(Tx_Desde) <= vFmt(Tx_UltTimbrado) Or vFmt(Tx_Hasta) <= vFmt(Tx_UltTimbrado) Then
      If MsgBox1("¡ATENCION! Usted está indicando un número de folio menor o igual a un folio ya timbrado. ¿Está seguro de continuar?", vbQuestion Or vbYesNo Or vbDefaultButton2) <> vbYes Then
         Exit Function
      End If
   End If
   
   If vFmt(Tx_Desde) <= vFmt(Tx_UltUsado) Or vFmt(Tx_Hasta) <= vFmt(Tx_UltUsado) Then
      MsgBox1 "No puede volver a imprimir folio ya usados.", vbExclamation
      Exit Function
   End If
   
   If vFmt(Tx_Hasta) = 0 Then
      MsgBox1 "Debe ingresar el folio hasta el cual desea imprimir.", vbExclamation
      Tx_Hasta.SetFocus
      Exit Function
   End If
   
   If vFmt(Tx_Desde) > vFmt(Tx_Hasta) Then
      MsgBox1 "Folio inicio no puede ser mayor al folio de término.", vbExclamation
      Tx_Hasta.SetFocus
      Exit Function
   End If
   
   Valida = True
   
End Function
Private Sub PrtHojasFoliadas()
   Dim LeftX As Integer
   Dim RightX As Long
   Dim OldFName As String
   Dim OldFBold As Integer
   Dim OldFSize As Single
   Dim PrtOrient As Integer
   Dim nhojas As Long
   Dim TLeft As Integer
   Dim Width As Long
   Dim nFolio As String
   Dim nDesde As Long
   Dim nHasta As Long
   Dim TabX As Integer
      
   If SelPrinter() Then
      Exit Sub
   End If
   
   If gUseCourier = False Then
      'veamos si hay problema con el font
      On Error Resume Next
      Printer.FontName = "Arial"
      Printer.FontSize = 9
      Printer.FontBold = False
      
      If Err Then
         MsgBox "Error " & Err & ", " & Error, vbExclamation
      End If
   End If
   
   TLeft = 10
   TabX = 28
   OldFName = Printer.FontName
   OldFBold = Printer.FontBold
   OldFSize = Printer.FontSize
   PrtOrient = Printer.Orientation
   Printer.Orientation = lIdxOrientacion
   
   'ENCABEZADO EMPRESA
   Printer.FontSize = 9
   Printer.FontName = "Arial"
   Printer.FontBold = False
   
   nDesde = vFmt(Tx_Desde)
   nHasta = vFmt(Tx_Hasta)
   
   For nhojas = nDesde To nHasta
      Width = Printer.Width - 1000
      nFolio = "Folio: " & Right("00000000" & nhojas, 8)
      
      Printer.Print
      'Printer.Print
      Printer.Print Tab(TLeft); "Razón Social: "; Tab(TabX); gEmpresa.RazonSocial;
      Printer.FontSize = 11
      Printer.CurrentX = Width - Printer.TextWidth(nFolio)
      Printer.Print nFolio
      Printer.FontSize = 9
      If gEmpresa.RutDisp = "" Then
         Printer.Print Tab(TLeft); "RUT:"; Tab(TabX); FmtCID(gEmpresa.Rut)
      Else
         Printer.Print Tab(TLeft); "RUT:"; Tab(TabX); FmtCID(gEmpresa.RutDisp)
      End If
      Printer.Print Tab(TLeft); "Dirección:"; Tab(TabX); gEmpresa.Direccion & ", " & IIf(gEmpresa.Ciudad <> "", gEmpresa.Ciudad, gEmpresa.Comuna)
      Printer.Print Tab(TLeft); "Giro:"; Tab(TabX); gEmpresa.Giro
      If gEmpresa.RepLegal1 <> "" Then
         Printer.Print Tab(TLeft); "Rep. Legal:"; Tab(TabX); gEmpresa.RepLegal1
         Printer.Print Tab(TLeft); "RUT Rep. Legal:"; Tab(TabX); FmtCID(gEmpresa.RutRepLegal1)
      End If
      If gEmpresa.RepConjunta And gEmpresa.RepLegal2 <> "" Then
         Printer.Print Tab(TLeft); "Rep. Legal:"; Tab(TabX); gEmpresa.RepLegal2
         Printer.Print Tab(TLeft); "RUT Rep. Legal:"; Tab(TabX); FmtCID(gEmpresa.RutRepLegal2)
      End If
      
      If nhojas < nHasta Then
         Printer.NewPage
      End If
      
   Next nhojas
   
   Printer.EndDoc
   
   Printer.FontName = OldFName
   Printer.FontBold = OldFBold
   Printer.FontSize = OldFSize
   Printer.Orientation = PrtOrient
   
End Sub

Private Sub Form_Load()

   Op_Orientacion(ORIENT_VER) = True
   
   Call LoadAll
   
   'Call EnableForm(Me, gEmpresa.FCierre = 0)
   
   Call SetupPriv
   
End Sub

Private Sub Op_Orientacion_Click(Index As Integer)
   lIdxOrientacion = Index
   
End Sub
Private Sub LoadAll()
   
   If gFoliacion.Estado = EF_EXISTE Then
   
      Tx_UltImpreso = IIf(gFoliacion.UltImpreso <> 0, gFoliacion.UltImpreso, "")
      Tx_FUltImpreso = IIf(gFoliacion.FUltImpreso <> 0, Format(gFoliacion.FUltImpreso, DATEFMT), "")
      
      Tx_UltTimbrado = IIf(gFoliacion.UltTimbrado <> 0, gFoliacion.UltTimbrado, "")
      Tx_FUltTimbrado = IIf(gFoliacion.FUltTimbrado <> 0, Format(gFoliacion.FUltTimbrado, DATEFMT), "")
      
      Tx_UltUsado = IIf(gFoliacion.UltUsado <> 0, gFoliacion.UltUsado, "")
      Tx_FUltUsado = IIf(gFoliacion.FUltUsado <> 0, Format(gFoliacion.FUltUsado, DATEFMT), "")
   
   End If
   
   Tx_Desde = vFmt(Tx_UltImpreso) + 1
   Tx_Hasta = Tx_Desde
   
End Sub

Private Sub SetupPriv()

    If Not ChkPriv(PRV_ADM_TIMB) Then
        Call EnableForm(Me, False)
    End If
End Sub
