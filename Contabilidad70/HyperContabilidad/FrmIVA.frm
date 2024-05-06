VERSION 5.00
Begin VB.Form FrmIVA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Impuestos"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   Icon            =   "FrmIVA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Afecto Libro Compra y Ventas"
      Height          =   1335
      Left            =   1440
      TabIndex        =   25
      Top             =   6600
      Width           =   3975
      Begin VB.CheckBox Ch_AfectoCero 
         Caption         =   "Afecto Igual 0"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   2475
      End
      Begin VB.Label Label2 
         Caption         =   "¿Permite en Compras y Ventas Facturas Afectas con monto neto igual a Cero?"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.CommandButton Bt_Ayuda 
      Caption         =   "Tasas 33 Bis"
      Height          =   375
      Left            =   6060
      TabIndex        =   21
      ToolTipText     =   "Ayuda para ingresar la Tasa de Crédito Activo Fijo"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton bt_Ok 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   6060
      TabIndex        =   20
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6060
      TabIndex        =   19
      Top             =   900
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Crédito Activo Fijo (Art. 33 Bis Ley de renta)"
      Height          =   4035
      Left            =   1440
      TabIndex        =   9
      Top             =   2520
      Width           =   3975
      Begin VB.Frame Frame4 
         Caption         =   "Hasta 30 Sept. 2014"
         Height          =   915
         Left            =   300
         TabIndex        =   16
         Top             =   360
         Width           =   3315
         Begin VB.TextBox Tx_CredArt33 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   2
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   6
            Left            =   2880
            TabIndex        =   18
            Top             =   420
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Crédito Activo Fijo:"
            Height          =   195
            Index           =   5
            Left            =   420
            TabIndex        =   17
            Top             =   420
            Width           =   1320
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "A contar del 1 Oct. 2015"
         Height          =   975
         Index           =   2
         Left            =   300
         TabIndex        =   13
         Top             =   2640
         Width           =   3315
         Begin VB.TextBox Tx_CredArt33 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   4
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   10
            Left            =   2880
            TabIndex        =   15
            Top             =   420
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Crédito Activo Fijo:"
            Height          =   195
            Index           =   9
            Left            =   420
            TabIndex        =   14
            Top             =   420
            Width           =   1320
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "1 Oct. 2014 a 30 Sept. 2015"
         Height          =   915
         Index           =   1
         Left            =   300
         TabIndex        =   10
         Top             =   1500
         Width           =   3315
         Begin VB.TextBox Tx_CredArt33 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   3
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Crédito Activo Fijo:"
            Height          =   195
            Index           =   8
            Left            =   420
            TabIndex        =   12
            Top             =   420
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   7
            Left            =   2880
            TabIndex        =   11
            Top             =   420
            Width           =   120
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   1440
      TabIndex        =   6
      Top             =   480
      Width           =   3975
      Begin VB.TextBox Tx_ImpPrimCategoria 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         MaxLength       =   4
         TabIndex        =   1
         Top             =   960
         Width           =   795
      End
      Begin VB.TextBox Tx_IVA 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2580
         MaxLength       =   4
         TabIndex        =   0
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "vigente hasta el 31.12.2019"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   24
         Top             =   1260
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   2
         Left            =   3420
         TabIndex        =   23
         Top             =   1020
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje IDPC 14 TER A:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   22
         Top             =   1020
         Width           =   1980
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje IVA:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   4
         Left            =   3420
         TabIndex        =   7
         Top             =   540
         Width           =   120
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   420
      Picture         =   "FrmIVA.frx":000C
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   600
      Width           =   675
   End
End
Attribute VB_Name = "FrmIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_CRED33OLD = 1
Const C_CRED33_OCT2014 = 2
Const C_CRED33_OCT2015 = 3


Private Sub Bt_Ayuda_Click()
   Dim Frm As FrmHelpCred33bis
   
   Set Frm = New FrmHelpCred33bis
   Frm.Show vbModal
   Set Frm = Nothing
   

End Sub

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_OK_Click()
   Dim Q1 As String
   Dim IVA As Single, Imp1Cat As Single
   Dim CredArt33 As Single
   Dim i As Integer
   
   Dim Rs As Recordset
   
   If vFmt(Tx_IVA) <= 0 Then
      MsgBox1 "El porcentaje de IVA debe ser mayor que cero.", vbExclamation
      Exit Sub
   End If
   
   If gEmpresa.Franq14Ter Then
      If vFmt(Tx_ImpPrimCategoria) <= 0 Then
         MsgBox1 "El porcentaje de Impuesto de Primera Categoría debe ser mayor que cero.", vbExclamation
         Exit Sub
      End If
   End If
   
   If vFmt(Tx_IVA) >= 100 Then
      MsgBox1 "El porcentaje de IVA debe ser inferior al 100%.", vbExclamation
      Exit Sub
   End If
   
   If gEmpresa.Franq14Ter Then
      If vFmt(Tx_ImpPrimCategoria) >= 100 Then
         MsgBox1 "El porcentaje de Impuesto de Primera Categoría debe ser inferior al 100%.", vbExclamation
         Exit Sub
      End If
   End If
   
   For i = 1 To C_CRED33_OCT2015
      If vFmt(Tx_CredArt33(i)) < 0 Then
         MsgBox1 "El porcentaje de crédito para Activo Fijo debe ser mayor o igual a cero.", vbExclamation
         Exit Sub
      End If
      
      If vFmt(Tx_CredArt33(i)) >= 100 Then
         MsgBox1 "El porcentaje de crédito para Activo Fijo debe ser inferior al 100%.", vbExclamation
         Exit Sub
      End If
   Next i
   
   IVA = vFmt(Tx_IVA) / 100
   If gIVA <> IVA Then
      If MsgBox1("¿Está seguro de cambiar el valor del IVA?", vbQuestion Or vbYesNo Or vbDefaultButton2) <> vbYes Then
         Exit Sub
      End If
   End If
   
   IVA = vFmt(Tx_IVA) / 100
   Q1 = "UPDATE ParamEmpresa SET Valor='" & str(IVA) & "' WHERE Tipo='VALORIVA'"
   Q1 = Q1 & " AND IdEmpresa = 0 AND Ano = 0"
   Call ExecSQL(DbMain, Q1)
   
   gIVA = IVA
   
   If gEmpresa.Franq14Ter Then
      Imp1Cat = vFmt(Tx_ImpPrimCategoria) / 100
      Q1 = "UPDATE ParamEmpresa SET Valor='" & str(Imp1Cat) & "' WHERE Tipo='IMP1CAT'"
'      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Q1 = Q1 & " AND IdEmpresa = 0 AND Ano = 0"
      Call ExecSQL(DbMain, Q1)
      
      gImpPrimCategoria = Imp1Cat
   End If
   
   CredArt33 = vFmt(Tx_CredArt33(C_CRED33OLD)) / 100
   Q1 = "UPDATE ParamEmpresa SET Valor='" & str(CredArt33) & "' WHERE Tipo='CREDART33'"
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " AND IdEmpresa = 0 AND Ano = 0"
   Call ExecSQL(DbMain, Q1)
   
   gCredArt33 = CredArt33
   
   CredArt33 = vFmt(Tx_CredArt33(C_CRED33_OCT2014)) / 100
   Q1 = "UPDATE ParamEmpresa SET Valor='" & str(CredArt33) & "' WHERE Tipo='CREDART334'"
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " AND IdEmpresa = 0 AND Ano = 0"
   Call ExecSQL(DbMain, Q1)
   
   gCredArt33_2014 = CredArt33
  
   CredArt33 = vFmt(Tx_CredArt33(C_CRED33_OCT2015)) / 100
   Q1 = "UPDATE ParamEmpresa SET Valor='" & str(CredArt33) & "' WHERE Tipo='CREDART335'"
'   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'   Q1 = Q1 & " AND IdEmpresa = 0 AND Ano = 0"
   Call ExecSQL(DbMain, Q1)
   
   gCredArt33_2015 = CredArt33
   
   
   Q1 = "SELECT Valor FROM ParamEmpresa "
   Q1 = Q1 & " WHERE Tipo='AFECTOCERO' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = True Then
   
    Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES( 'AFECTOCERO', 0, '" & Ch_AfectoCero.Value & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")"
    Call ExecSQL(DbMain, Q1)
    
    gAfectoCero = Ch_AfectoCero.Value
     
   Else
   
   Q1 = "UPDATE ParamEmpresa SET Valor='" & Ch_AfectoCero.Value & "' WHERE Tipo='AFECTOCERO'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   gAfectoCero = Ch_AfectoCero.Value
   End If
   
   
   Unload Me
   
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   Dim Rs As Recordset
   
   Tx_IVA = Format(gIVA * 100, NUMFMT)
   Tx_ImpPrimCategoria = Format(gImpPrimCategoria * 100, NUMFMT)
   Tx_CredArt33(C_CRED33OLD) = Format(gCredArt33 * 100, NUMFMT)
   Tx_CredArt33(C_CRED33_OCT2014) = Format(gCredArt33_2014 * 100, NUMFMT)
   Tx_CredArt33(C_CRED33_OCT2015) = Format(gCredArt33_2015 * 100, DBLFMT2)
   
   If Not gEmpresa.Franq14Ter Or gEmpresa.Ano >= 2020 Then
      Call SetRO(Tx_ImpPrimCategoria, True)
      Tx_ImpPrimCategoria = ""
   End If
   
   Q1 = "SELECT Valor FROM ParamEmpresa "
   Q1 = Q1 & " WHERE Tipo='AFECTOCERO' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Ch_AfectoCero.Value = vFld(Rs("Valor"))
   End If
   
   Call CloseRs(Rs)
   
   
End Sub

Private Sub Tx_CredArt33_GotFocus(Index As Integer)
   If Trim(Tx_CredArt33(Index)) <> "" Then
      Exit Sub
   End If
   Tx_CredArt33(Index) = vFmt(Tx_CredArt33(Index))

End Sub

Private Sub Tx_CredArt33_KeyPress(Index As Integer, KeyAscii As Integer)
   Call KeyDec(KeyAscii)

End Sub

Private Sub Tx_CredArt33_LostFocus(Index As Integer)
   If Trim(Tx_CredArt33(Index)) <> "" Then
      Exit Sub
   End If
   Tx_CredArt33(Index) = Format(Tx_CredArt33(Index), NUMFMT)

End Sub

Private Sub Tx_IVA_GotFocus()
   If Trim(Tx_IVA) <> "" Then
      Exit Sub
   End If
   Tx_IVA = vFmt(Tx_IVA)
End Sub

Private Sub Tx_IVA_KeyPress(KeyAscii As Integer)
   Call KeyDec(KeyAscii)
End Sub

Private Sub Tx_IVA_LostFocus()
   If Trim(Tx_IVA) <> "" Then
      Exit Sub
   End If
   Tx_IVA = Format(Tx_IVA, NUMFMT)
   
End Sub

Private Sub Tx_ImpPrimCategoria_GotFocus()
   If Trim(Tx_ImpPrimCategoria) <> "" Then
      Exit Sub
   End If
   Tx_ImpPrimCategoria = vFmt(Tx_ImpPrimCategoria)
End Sub

Private Sub Tx_ImpPrimCategoria_KeyPress(KeyAscii As Integer)
   Call KeyDec(KeyAscii)
End Sub

Private Sub Tx_ImpPrimCategoria_LostFocus()
   If Trim(Tx_ImpPrimCategoria) <> "" Then
      Exit Sub
   End If
   Tx_ImpPrimCategoria = Format(Tx_ImpPrimCategoria, NUMFMT)
   
End Sub

