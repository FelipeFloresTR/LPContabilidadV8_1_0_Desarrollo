VERSION 5.00
Begin VB.Form FrmParamRaz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros Razones Financieras"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_Denominador 
      Caption         =   "Cuentas (Rubros) Denominador"
      Height          =   2355
      Left            =   1320
      TabIndex        =   13
      Top             =   5400
      Width           =   6615
      Begin VB.CommandButton Bt_AddDenomResta 
         Caption         =   "&Agregar (-)"
         Height          =   585
         Left            =   5280
         Picture         =   "FrmParamRaz.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Agregar cuenta que resta a la lista"
         Top             =   900
         Width           =   1095
      End
      Begin VB.CommandButton Bt_AddDenomSuma 
         Caption         =   "&Agregar (+)"
         Height          =   585
         Left            =   5280
         Picture         =   "FrmParamRaz.frx":040E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Agregar cuenta que suma a la lista"
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton Bt_DelDenom 
         Caption         =   "&Eliminar"
         Height          =   585
         Left            =   5280
         Picture         =   "FrmParamRaz.frx":081C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Eliminar cuenta seleccionada"
         Top             =   1500
         Width           =   1095
      End
      Begin VB.ListBox Ls_CtaDenominador 
         Height          =   1815
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   4815
      End
   End
   Begin VB.Frame Fr_Numerador 
      Caption         =   "Cuentas (Rubros) Numerador"
      Height          =   2355
      Left            =   1320
      TabIndex        =   12
      Top             =   2820
      Width           =   6615
      Begin VB.CommandButton Bt_AddNumResta 
         Caption         =   "&Agregar (-)"
         Height          =   585
         Left            =   5280
         Picture         =   "FrmParamRaz.frx":0B26
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Agregar cuenta que resta a la lista"
         Top             =   900
         Width           =   1095
      End
      Begin VB.CommandButton Bt_AddNumSuma 
         Caption         =   "&Agregar (+)"
         Height          =   585
         Left            =   5280
         Picture         =   "FrmParamRaz.frx":0F34
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Agregar cuenta que suma a la lista"
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton Bt_DelNum 
         Caption         =   "&Eliminar"
         Height          =   585
         Left            =   5280
         Picture         =   "FrmParamRaz.frx":1342
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Eliminar cuenta seleccionada"
         Top             =   1500
         Width           =   1095
      End
      Begin VB.ListBox Ls_CtaNumerador 
         Height          =   1815
         Left            =   240
         TabIndex        =   0
         Top             =   300
         Width           =   4815
      End
   End
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   8100
      TabIndex        =   2
      Top             =   600
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Razón Financiera"
      Height          =   2235
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   6615
      Begin VB.TextBox Tx_CantDias 
         Height          =   315
         Left            =   5340
         TabIndex        =   22
         ToolTipText     =   "El resultado de la razón financiera se multiplica por esta cantidad de días"
         Top             =   1020
         Width           =   1035
      End
      Begin VB.TextBox Tx_Tipo 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   660
         Width           =   4635
      End
      Begin VB.TextBox Tx_Denominador 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1740
         Width           =   4635
      End
      Begin VB.TextBox Tx_Nombre 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   300
         Width           =   4635
      End
      Begin VB.TextBox Tx_Unidad 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox Tx_Numerador 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1380
         Width           =   4635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad de días:"
         Height          =   195
         Index           =   5
         Left            =   4020
         TabIndex        =   23
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   15
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Texto Denominador:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   11
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unidad del resultado:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Texto Numerador:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   1440
         Width           =   1275
      End
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   480
      Picture         =   "FrmParamRaz.frx":164C
      Top             =   540
      Width           =   585
   End
End
Attribute VB_Name = "FrmParamRaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lRc As Integer
Dim lOper As Integer
Dim lIdRazon As Long

Private Sub Bt_AddDenomSuma_Click()

   Call AddCta(CTA_DENOMINADOR, "+")

End Sub
Private Sub Bt_AddDenomResta_Click()

   Call AddCta(CTA_DENOMINADOR, "-")

End Sub

Private Sub Bt_AddNumSuma_Click()

   Call AddCta(CTA_NUMERADOR, "+")

End Sub
Private Sub Bt_AddNumResta_Click()

   Call AddCta(CTA_NUMERADOR, "-")

End Sub

Private Sub AddCta(ByVal NumDenom As Integer, ByVal Operador As String)
   Dim Frm As FrmPlanCuentas
   Dim IdCuenta As Long
   Dim Codigo As String
   Dim Descrip As String
   Dim Nombre As String
   Dim Rs As Recordset
   Dim Q1 As String
   
   Set Frm = New FrmPlanCuentas
   
   If Frm.FSelect(IdCuenta, Codigo, Descrip, Nombre, False) = vbOK Then
      
      Q1 = "SELECT CodCuenta FROM CuentasRazon WHERE IdRazon=" & lIdRazon
      Q1 = Q1 & " AND NumDenom = " & NumDenom
      Q1 = Q1 & " AND CodCuenta = '" & Codigo & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         MsgBox1 "Esta cuenta ya está en la lista.", vbExclamation + vbOKOnly
         Call CloseRs(Rs)
         Exit Sub
      End If
      
      Call CloseRs(Rs)
      
      Q1 = "INSERT INTO CuentasRazon "
      Q1 = Q1 & " (IdRazon, NumDenom, CodCuenta, Operador, IdEmpresa) "
      Q1 = Q1 & " VALUES(" & lIdRazon & "," & NumDenom & ",'" & Codigo & "','" & Trim(Operador) & "'," & gEmpresa.id & ")"
      Call ExecSQL(DbMain, Q1)
      
      Call LoadCuentas(NumDenom)

   End If

End Sub
Private Sub DelCta(ByVal NumDenom As Integer, ByVal CodCuenta As String)
   Dim Q1 As String
   
'   Q1 = "DELETE * FROM CuentasRazon "
'   Q1 = Q1 & " WHERE IdRazon = " & lIdRazon
'   Q1 = Q1 & " AND NumDenom = " & NumDenom
'   Q1 = Q1 & " AND CodCuenta = '" & VFmtCodigoCta(CodCuenta) & "'"
'
'   Call ExecSQL(DbMain, Q1)
   
   Q1 = " WHERE IdRazon = " & lIdRazon
   Q1 = Q1 & " AND NumDenom = " & NumDenom
   Q1 = Q1 & " AND CodCuenta = '" & VFmtCodigoCta(CodCuenta) & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
   Call DeleteSQL(DbMain, "CuentasRazon", Q1)
   
   
   Call LoadCuentas(NumDenom)
   
End Sub


Private Sub bt_Cerrar_Click()

   If SaveParam Then
    
      lRc = vbOK
      Unload Me
      
   End If

End Sub

Private Sub Bt_DelDenom_Click()
   Dim Idx As Integer
   Dim CodCuenta As String

   If Ls_CtaDenominador.ListIndex < 0 Then
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea eliminar esta cuenta de la lista de cuentas del Denominador?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
   End If
   
   Idx = InStr(3, Ls_CtaDenominador, " ")     'saltamos el operador y el blanco antes del código
   
   If Idx > 0 Then
      CodCuenta = Mid(Ls_CtaDenominador, 3, Idx - 3)
   
      Call DelCta(CTA_DENOMINADOR, CodCuenta)
   End If
   
End Sub

Private Sub Bt_DelNum_Click()
   Dim Idx As Integer
   Dim CodCuenta As String
   
   If Ls_CtaNumerador.ListIndex < 0 Then
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea eliminar esta cuenta de la lista de cuentas del Numerador?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
   End If
   
   Idx = InStr(3, Ls_CtaNumerador, " ")     'saltamos el operador y el blanco antes del código
   
   If Idx > 0 Then
      CodCuenta = Mid(Ls_CtaNumerador, 3, Idx - 3)
   
      Call DelCta(CTA_NUMERADOR, CodCuenta)
   End If
   
End Sub
Private Sub Form_Load()

   Call SetTxRO(Tx_CantDias, True)  'debe estar antes del LoadRazon
   
   Call LoadRazon
   
   If Not Tx_CantDias.Locked Then
      Bt_Cerrar.Caption = "Guardar"
   End If
   
   If Not ChkPriv(PRV_CFG_EMP) Or lOper = O_VIEW Then
      Bt_AddNumSuma.Enabled = False
      Bt_AddNumResta.Enabled = False
      Bt_AddDenomSuma.Enabled = False
      Bt_AddDenomResta.Enabled = False
      Bt_DelNum.Enabled = False
      Bt_DelDenom.Enabled = False
   End If

End Sub

Private Sub LoadRazon()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
      
   If lIdRazon <= 0 Then
      Exit Sub
   End If
   
   'fill detalle
   Q1 = "SELECT Nombre, Tipo, UnidadRes, TxtNumerador, TxtDenominador FROM RazonesFin"
   Q1 = Q1 & " WHERE IdRazon = " & lIdRazon
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      
      Tx_Nombre = vFld(Rs("Nombre"), True)
      
      For i = 0 To UBound(gTipoRazFin)
         If gTipoRazFin(i).id = vFld(Rs("Tipo")) Then
            Tx_Tipo = gTipoRazFin(i).Nombre
            Exit For
         End If
      Next i
      
      Tx_Unidad = vFld(Rs("UnidadRes"), True)
      If Trim(LCase(Tx_Unidad)) = "dias" Or Trim(LCase(Tx_Unidad)) = "días" Then
         Tx_CantDias = 365
         Call SetTxRO(Tx_CantDias, False)
      Else
         Tx_CantDias = ""
         Call SetTxRO(Tx_CantDias, True)
      End If

      Tx_Numerador = vFld(Rs("TxtNumerador"), True)
      Tx_Denominador = vFld(Rs("TxtDenominador"), True)
   End If
   
   Call CloseRs(Rs)
   
   If Tx_CantDias.Locked = False Then
      Q1 = "SELECT CantDias FROM ParamRazon "
      Q1 = Q1 & " WHERE IdRazon = " & lIdRazon
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = False Then
         Tx_CantDias = IIf(vFld(Rs("CantDias")) = 0, 365, Format(vFld(Rs("CantDias")), NUMFMT))
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   If Tx_Denominador = "" Then
'      Q1 = "DELETE * FROM CuentasRazon "
'      Q1 = Q1 & " WHERE IdRazon = " & lIdRazon
'      Q1 = Q1 & " AND NumDenom = " & CTA_DENOMINADOR
'
'      Call ExecSQL(DbMain, Q1)
      
      Q1 = " WHERE IdRazon = " & lIdRazon
      Q1 = Q1 & " AND NumDenom = " & CTA_DENOMINADOR
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      
      Call DeleteSQL(DbMain, "CuentasRazon", Q1)

      Fr_Denominador.Enabled = False
   End If
   
   Call LoadCuentas(CTA_NUMERADOR)
   Call LoadCuentas(CTA_DENOMINADOR)

End Sub
Private Sub LoadCuentas(ByVal NumDenom As Integer)
   Dim Rs As Recordset
   Dim Q1 As String
   
   If NumDenom = CTA_NUMERADOR Then
      Ls_CtaNumerador.Clear
   Else
      Ls_CtaDenominador.Clear
   End If
   
   Q1 = "SELECT CuentasRazon.CodCuenta, Descripcion, CuentasRazon.Operador, Cuentas.IdCuenta "
   Q1 = Q1 & " FROM CuentasRazon INNER JOIN Cuentas ON CuentasRazon.CodCuenta = Cuentas.Codigo "
   Q1 = Q1 & " AND CuentasRazon.IdEmpresa = Cuentas.IdEmpresa "
   Q1 = Q1 & " WHERE IdRazon = " & lIdRazon
   Q1 = Q1 & " AND NumDenom = " & NumDenom
   Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
   
      If NumDenom = CTA_NUMERADOR Then
         Ls_CtaNumerador.AddItem vFld(Rs("Operador"), True) & " " & Format(vFld(Rs("CodCuenta")), gFmtCodigoCta) & " " & vFld(Rs("Descripcion"), True)
         Ls_CtaNumerador.ItemData(Ls_CtaNumerador.NewIndex) = vFld(Rs("IdCuenta"))
      
      Else
         Ls_CtaDenominador.AddItem vFld(Rs("Operador"), True) & " " & Format(vFld(Rs("CodCuenta")), gFmtCodigoCta) & " " & vFld(Rs("Descripcion"), True)
         Ls_CtaDenominador.ItemData(Ls_CtaDenominador.NewIndex) = vFld(Rs("IdCuenta"))
      
      End If
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)

End Sub
Public Function FEdit(ByVal IdRazon As Long) As Integer
   lOper = O_EDIT
   lIdRazon = IdRazon
   
   Me.Show vbModal
   
   FEdit = lRc

End Function
Public Function FView(ByVal IdRazon As Long) As Integer
   lOper = O_VIEW
   lIdRazon = IdRazon
   
   Me.Show vbModal
   
   FView = lRc

End Function


Private Sub Tx_CantDias_KeyPress(KeyAscii As Integer)
   
   Call KeyNumPos(KeyAscii)
   
End Sub

Private Function SaveParam() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   If Tx_CantDias.Locked Then
      SaveParam = True
      Exit Function
   End If
   
   SaveParam = False
   
   If vFmt(Tx_CantDias) <= 0 Then
      MsgBox1 "Cantidad de días inválido.", vbExclamation
      Exit Function
   End If
   
   
   Q1 = "SELECT * FROM ParamRazon WHERE IdRazon = " & lIdRazon
   '3024907
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id '& " AND Ano = " & gEmpresa.Ano
   'Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   '3024907
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      
      Q1 = "UPDATE ParamRazon SET CantDias = " & vFmt(Tx_CantDias) & " WHERE IdRazon = " & lIdRazon
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      Call ExecSQL(DbMain, Q1)
      
   Else
   
      Q1 = "INSERT INTO ParamRazon (IdRazon, CantDias, IdEmpresa) VALUES(" & lIdRazon & "," & vFmt(Tx_CantDias) & "," & gEmpresa.id & ")"
      Call ExecSQL(DbMain, Q1)
   
   End If
       
   Call CloseRs(Rs)
   
   SaveParam = True
     
End Function
