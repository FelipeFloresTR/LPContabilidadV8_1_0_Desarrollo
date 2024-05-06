VERSION 5.00
Begin VB.Form FrmFoliacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información de Folios"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "FrmFoliacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_Clear 
      Cancel          =   -1  'True
      Caption         =   "Limpiar datos"
      Height          =   915
      Left            =   8460
      Picture         =   "FrmFoliacion.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1500
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Ok 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   8460
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8460
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información Folios"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Index           =   0
      Left            =   1140
      TabIndex        =   12
      Top             =   480
      Width           =   6975
      Begin VB.CommandButton Bt_SelFecha 
         Height          =   315
         Index           =   0
         Left            =   6180
         Picture         =   "FrmFoliacion.frx":0477
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   420
         Width           =   255
      End
      Begin VB.TextBox Tx_UltImpreso 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1785
         TabIndex        =   0
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox Tx_UltTimbrado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Tx_FUltImpreso 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5100
         TabIndex        =   1
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox Tx_FUltTimbrado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5100
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Tx_FUltUsado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5100
         TabIndex        =   7
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox Tx_UltUsado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CommandButton Bt_SelFecha 
         Height          =   315
         Index           =   2
         Left            =   6180
         Picture         =   "FrmFoliacion.frx":04EC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1260
         Width           =   255
      End
      Begin VB.CommandButton Bt_SelFecha 
         Height          =   315
         Index           =   1
         Left            =   6180
         Picture         =   "FrmFoliacion.frx":0561
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Último impreso:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   18
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Último timbrado:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   17
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Último usado:"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   16
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha último impreso:"
         Height          =   195
         Index           =   5
         Left            =   3420
         TabIndex        =   15
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha último timbrado:"
         Height          =   195
         Index           =   6
         Left            =   3420
         TabIndex        =   14
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha último usado:"
         Height          =   195
         Index           =   7
         Left            =   3420
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   300
      Picture         =   "FrmFoliacion.frx":05D6
      Top             =   540
      Width           =   555
   End
End
Attribute VB_Name = "FrmFoliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lRc As Integer
Dim lClear As Boolean

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_Clear_Click()
   If MsgBox1("¿Está seguro que desea dejar en cero todos los folios?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
   Tx_UltImpreso = ""
   Tx_FUltImpreso = ""
   Tx_UltTimbrado = ""
   Tx_FUltTimbrado = ""
   Tx_UltUsado = ""
   Tx_FUltUsado = ""
   
   lClear = True
   
End Sub

Private Sub Bt_OK_Click()
   Dim Q1 As String
   Dim Rs As Recordset
   
   If Valida() = False Then
      Exit Sub
   End If
   
   If gFoliacion.Estado = EF_NOEXISTE Then
      Q1 = "INSERT INTO Timbraje (idEmpresa) VALUES (" & gEmpresa.id & ")"
      Call ExecSQL(DbMain, Q1)
   End If
   
   Q1 = "UPDATE Timbraje SET "
   Q1 = Q1 & " UltImpreso = " & vFmt(Tx_UltImpreso)
   Q1 = Q1 & ", FUltImpreso=" & GetTxDate(Tx_FUltImpreso)
   Q1 = Q1 & ", UltTimbrado = " & vFmt(Tx_UltTimbrado)
   Q1 = Q1 & ", FUltTimbrado=" & GetTxDate(Tx_FUltTimbrado)
   Q1 = Q1 & ", UltUsado = " & vFmt(Tx_UltUsado)
   Q1 = Q1 & ", FUltUsado=" & GetTxDate(Tx_FUltUsado)
   
   Q1 = Q1 & " WHERE idEmpresa=" & gEmpresa.id
   Call ExecSQL(DbMain, Q1)
   
   gFoliacion.UltImpreso = vFmt(Tx_UltImpreso)
   gFoliacion.FUltImpreso = GetTxDate(Tx_FUltImpreso)

   gFoliacion.UltTimbrado = vFmt(Tx_UltTimbrado)
   gFoliacion.FUltTimbrado = GetTxDate(Tx_FUltTimbrado)
   
   gFoliacion.UltUsado = vFmt(Tx_UltUsado)
   gFoliacion.FUltUsado = GetTxDate(Tx_FUltUsado)
   
   gFoliacion.Estado = EF_EXISTE
   
   lRc = vbOK
   
   Unload Me
   
End Sub

Private Sub Bt_SelFecha_Click(Index As Integer)
    Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   If Index = 0 Then
      Call Frm.TxSelDate(Tx_FUltImpreso)
   ElseIf Index = 1 Then
      Call Frm.TxSelDate(Tx_FUltTimbrado)
   Else
      Call Frm.TxSelDate(Tx_FUltUsado)
   End If
   
   Set Frm = Nothing
   
End Sub

Private Sub Form_Load()

   Call LoadAll
   
   Call EnableForm(Me, gEmpresa.FCierre = 0)
   
'   Call SettxRO(Tx_UltImpreso, True)
'   Call SettxRO(Tx_FUltImpreso, True)
   
   Call SetupPriv
   
End Sub
Private Sub LoadAll()

   'actualizamos variable global
   Call Foliacion
   
   Tx_UltImpreso = IIf(gFoliacion.UltImpreso <> 0, gFoliacion.UltImpreso, "")
   Tx_FUltImpreso = IIf(gFoliacion.FUltImpreso <> 0, Format(gFoliacion.FUltImpreso, DATEFMT), "")
   
   Tx_UltTimbrado = IIf(gFoliacion.UltTimbrado <> 0, gFoliacion.UltTimbrado, "")
   Tx_FUltTimbrado = IIf(gFoliacion.FUltTimbrado <> 0, Format(gFoliacion.FUltTimbrado, DATEFMT), "")
   
   Tx_UltUsado = IIf(gFoliacion.UltUsado <> 0, gFoliacion.UltUsado, "")
   Tx_FUltUsado = IIf(gFoliacion.FUltUsado <> 0, Format(gFoliacion.FUltUsado, DATEFMT), "")
   
End Sub
Private Function Valida() As Boolean
   Valida = False
   
   If lClear Then
      Valida = True
      Exit Function
   End If

   '*** ULTIMO FOLIO IMPRESO
   If vFmt(Tx_UltImpreso) <> 0 And GetTxDate(Tx_FUltImpreso) = 0 Then
      MsgBox1 "No ha ingresado la fecha para el último folio impreso.", vbExclamation
      Tx_UltTimbrado.SetFocus
      Exit Function
   End If
   
   '*** ULTIMO FOLIO TIMBRADO
   
   If vFmt(Tx_UltImpreso) = 0 And vFmt(Tx_UltTimbrado) <> 0 Then
      MsgBox1 "No existen folios impresos para poder agregarlos como timbrados, debe imprimir o actualizar folios impresos.", vbExclamation
      Tx_UltTimbrado.SetFocus
      Exit Function
   End If
   
   If vFmt(Tx_UltImpreso) < vFmt(Tx_UltTimbrado) Then
      MsgBox1 "El último folio timbrado no puede ser mayor al último folio impreso.", vbExclamation
      Tx_UltTimbrado.SetFocus
      Exit Function
   End If
   
'   If vFmt(Tx_UltTimbrado) < gFoliacion.UltTimbrado Then
'      MsgBox1 "El último folio timbrado fue " & gFoliacion.UltTimbrado & ", usted está ingresando uno menor a éste.", vbExclamation
'      Tx_UltTimbrado.SetFocus
'      Exit Function
'   End If
   
   If vFmt(Tx_UltTimbrado) <> 0 And GetTxDate(Tx_FUltTimbrado) = 0 Then
      MsgBox1 "No ha ingresado la fecha para el último folio timbrado.", vbExclamation
      Tx_UltTimbrado.SetFocus
      Exit Function
   End If
   
'   If GetTxDate(Tx_FUltTimbrado) < gFoliacion.FUltTimbrado Then
'      MsgBox1 "Para Folio último timbraje ha ingresado una fecha menor a la que ya existía. " & Format(gFoliacion.FUltTimbrado, DATEFMT), vbExclamation
'      Tx_UltTimbrado.SetFocus
'      Exit Function
'   End If
   '*********
   
   '*****ULTIMO FOLIO USADO
   
   If vFmt(Tx_UltTimbrado) = 0 And vFmt(Tx_UltUsado) <> 0 Then
      MsgBox1 "No ha ingresado folios timbrados para poder agregarlos como usados, debe actualizar los folios timbrados ", vbExclamation
      Tx_UltUsado.SetFocus
      Exit Function
   End If
   
   If vFmt(Tx_UltTimbrado) < vFmt(Tx_UltUsado) Then
      MsgBox1 "Ha ingresado último folio usado mayor al último folio timbrado " & gFoliacion.UltTimbrado, vbExclamation
      Tx_UltUsado.SetFocus
      Exit Function
   End If
   
'   If vFmt(Tx_UltUsado) < gFoliacion.UltUsado Then
'      MsgBox1 "El último folio usado fue " & gFoliacion.UltUsado & ", usted está ingresando uno menor a éste", vbExclamation
'      Tx_UltUsado.SetFocus
'      Exit Function
'   End If
   
   If vFmt(Tx_UltUsado) <> 0 And GetTxDate(Tx_FUltUsado) = 0 Then
      MsgBox1 "No ha ingresado la fecha para el folio último usado.", vbExclamation
      Tx_FUltUsado.SetFocus
      Exit Function
   End If
   
'   If GetTxDate(Tx_FUltUsado) < gFoliacion.FUltUsado Then
'      MsgBox1 "Ha ingresado una fecha menor a la que ya existía para el último folio usado" & Format(gFoliacion.FUltUsado, DATEFMT), vbExclamation
'      Tx_FUltUsado.SetFocus
'      Exit Function
'   End If
   
   Valida = True
   
End Function

Private Sub Tx_FUltImpreso_GotFocus()
   Call DtGotFocus(Tx_FUltImpreso)
End Sub

Private Sub Tx_FUltImpreso_LostFocus()
   
   If Trim$(Tx_FUltImpreso) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FUltImpreso)
   
End Sub

Private Sub Tx_FUltImpreso_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Tx_FUltUsado_GotFocus()
   Call DtGotFocus(Tx_FUltUsado)
End Sub

Private Sub Tx_FUltUsado_LostFocus()
   
   If Trim$(Tx_FUltUsado) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FUltUsado)
   
End Sub

Private Sub Tx_FUltUsado_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub

Private Sub Tx_FUltTimbrado_GotFocus()
   Call DtGotFocus(Tx_FUltTimbrado)
   
End Sub

Private Sub Tx_FUltTimbrado_LostFocus()
   
   If Trim$(Tx_FUltTimbrado) = "" Then
      Exit Sub
   End If
   
   Call DtLostFocus(Tx_FUltTimbrado)
      
End Sub

Private Sub Tx_FUltTimbrado_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub

Public Function FEdit()

   Me.Show vbModal
   
   FEdit = lRc
End Function

Private Function SetupPriv()
   
   If Not ChkPriv(PRV_ADM_TIMB) Then
      Call EnableForm(Me, False)
   End If
   
End Function
