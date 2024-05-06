VERSION 5.00
Begin VB.Form FrmRazones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Definición Razones Financieras"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "FrmRazones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Razones Financieras"
      Height          =   5115
      Left            =   180
      TabIndex        =   19
      Top             =   120
      Width           =   7935
      Begin VB.ComboBox Cb_TipoLst 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   4635
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   600
         Picture         =   "FrmRazones.frx":000C
         ScaleHeight     =   675
         ScaleWidth      =   585
         TabIndex        =   21
         Top             =   900
         Width           =   585
      End
      Begin VB.CommandButton Bt_Del 
         Caption         =   "&Eliminar"
         Height          =   765
         Left            =   6600
         Picture         =   "FrmRazones.frx":0632
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Eliminar razón financiera seleccionada"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton Bt_DefCuentas 
         Caption         =   "Configurar"
         Height          =   765
         Left            =   6600
         Picture         =   "FrmRazones.frx":093C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Definir cuentas que intervienen en el cálculo de la razón financiera seleccionada"
         Top             =   2940
         Width           =   1095
      End
      Begin VB.CommandButton Bt_Cerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   6600
         TabIndex        =   13
         Top             =   360
         Width           =   1155
      End
      Begin VB.ListBox Ls_Razon 
         Height          =   3570
         Left            =   1740
         TabIndex        =   1
         Top             =   900
         Width           =   4635
      End
      Begin VB.Label Label3 
         Caption         =   "* Razón financiera predefinida, no modificable por el usuario"
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Top             =   4620
         Width           =   4455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   660
         TabIndex        =   20
         Top             =   420
         Width           =   360
      End
   End
   Begin VB.Frame Fr_DetRazon 
      Caption         =   "Detalle Razón Financiera Seleccionada"
      Height          =   3195
      Left            =   180
      TabIndex        =   14
      Top             =   5340
      Width           =   7935
      Begin VB.TextBox Tx_Glosa 
         Height          =   675
         Left            =   1740
         MaxLength       =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Definición de la parte inferior de la razón financiera"
         Top             =   2220
         Width           =   4635
      End
      Begin VB.CommandButton Bt_Clear 
         Caption         =   "Limpiar datos"
         Height          =   765
         Left            =   6600
         Picture         =   "FrmRazones.frx":0ED7
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpiar los datos de los textos"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ComboBox Cb_Tipo 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   660
         Width           =   4635
      End
      Begin VB.CommandButton Bt_Actualizar 
         Caption         =   "A&ctualizar"
         Height          =   765
         Left            =   6600
         Picture         =   "FrmRazones.frx":1388
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Actualizar la información razón financiera seleccionada"
         Top             =   1140
         Width           =   1095
      End
      Begin VB.CommandButton Bt_New 
         Caption         =   "&Agregar"
         Height          =   765
         Left            =   6600
         Picture         =   "FrmRazones.frx":195B
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Agregar nueva razón financiera"
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox Tx_Denominador 
         Height          =   315
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   8
         ToolTipText     =   "Definición de la parte inferior de la razón financiera"
         Top             =   1860
         Width           =   4635
      End
      Begin VB.TextBox Tx_Numerador 
         Height          =   315
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   7
         ToolTipText     =   "Definición de la parte superior de la razón financiera"
         Top             =   1500
         Width           =   4635
      End
      Begin VB.TextBox Tx_Unidad 
         Height          =   315
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   4
         Top             =   300
         Width           =   4635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Glosa:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   24
         Top             =   2280
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   22
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Texto Denominador:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Texto Numerador:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unidad del resultado:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmRazones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const O_DEFINIR = 100
Const O_CONFIGPARAM = 101

Dim lOper As Integer

Public Function FDefinir()

   lOper = O_DEFINIR
   Me.Show vbModal
   
End Function

Public Function FConfigParam()

   lOper = O_CONFIGPARAM
   Me.Show vbModal
   
End Function

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub


Private Sub Bt_Clear_Click()
   Call ClearRazon
End Sub

#If Admin <> 1 Then

Private Sub Bt_DefCuentas_Click()
   Dim Frm As FrmParamRaz
   
   If Ls_Razon.ListIndex < 0 Then
      Exit Sub
   End If
   
   Set Frm = New FrmParamRaz
   Frm.FEdit (ItemData(Ls_Razon))
   
   Set Frm = Nothing
   
End Sub

#End If

Private Sub Bt_Del_Click()
   Dim Q1 As String
   
   If ItemData(Ls_Razon) <= 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If MsgBox1("¿Está seguro que desea eliminar la Razón Financiera """ & Ls_Razon.List(Ls_Razon.ListIndex) & """?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
'   Q1 = "DELETE * FROM RazonesFin WHERE IdRazon = " & ItemData(Ls_Razon)
'   Call ExecSQL(DbMain, Q1)
   Q1 = " WHERE IdRazon = " & ItemData(Ls_Razon)
   Call DeleteSQL(DbMain, "RazonesFin", Q1)
   
   Ls_Razon.RemoveItem (Ls_Razon.ListIndex)
   
   If Ls_Razon.ListCount > 0 Then
      Ls_Razon.ListIndex = 0
   Else
      Call ClearRazon
   End If
      
      
End Sub

Private Sub Bt_Actualizar_Click()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdRazon As Long
   
   IdRazon = ItemData(Ls_Razon)
   
   If IdRazon <= 0 Then
      MsgBeep vbExclamation
      Exit Sub
   End If
   
   If Valida() Then
   
      'vemos si ya existe otra razón financiera con este nombre
      
      Q1 = "SELECT IdRazon FROM RazonesFin WHERE Nombre = '" & ParaSQL(Tx_Nombre) & "'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
      
         If IdRazon <> vFld(Rs("IdRazon")) Then
            MsgBox1 "Ya existe una razón financiera con este nombre.", vbExclamation
            Call CloseRs(Rs)
            Exit Sub
         End If
         
      End If
      
      Call CloseRs(Rs)
   
      Call SaveRazon(IdRazon)
      Call LoadLst(IdRazon)
         
   End If

End Sub

Private Sub SaveRazon(ByVal IdRazon As Long)
   Dim Q1 As String
   
   'actualizamos
   Q1 = "UPDATE RazonesFin SET "
   Q1 = Q1 & "  Nombre = '" & ParaSQL(Tx_Nombre) & "'"
   Q1 = Q1 & ", Tipo = " & ItemData(Cb_Tipo)
   Q1 = Q1 & ", UnidadRes = '" & ParaSQL(Tx_Unidad) & "'"
   Q1 = Q1 & ", TxtNumerador = '" & ParaSQL(Tx_Numerador) & "'"
   Q1 = Q1 & ", TxtDenominador = '" & ParaSQL(Tx_Denominador) & "'"
   Q1 = Q1 & ", Glosa = '" & ParaSQL(Tx_Glosa) & "'"
   Q1 = Q1 & ", Operador = '/'"
   Q1 = Q1 & " WHERE IdRazon = " & IdRazon
   
   Call ExecSQL(DbMain, Q1)
   
End Sub


Private Sub Bt_New_Click()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdRazon As Long
   
   If Valida() Then
   
      'vemos si ya existe una razón socal con este nombre
      
      Q1 = "SELECT IdRazon FROM RazonesFin WHERE Nombre = '" & ParaSQL(Tx_Nombre) & "'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         MsgBox1 "Ya existe una razón financiera con este nombre.", vbExclamation
         Call CloseRs(Rs)
         Exit Sub
      End If
      
      Call CloseRs(Rs)
      
      'insertamos
'      Set Rs = DbMain.OpenRecordset("RazonesFin")
'      Rs.AddNew
'
'      IdRazon = Rs("IdRazon")
'
'      Rs.Update
'      Rs.Close
'      Set Rs = Nothing
      
      IdRazon = AdvTbAddNew(DbMain, "RazonesFin", "IdRazon", "Nombre", ParaSQL(Tx_Nombre))
            
      'ahora el resto de los datos
      Call SaveRazon(IdRazon)
      Call LoadLst(IdRazon)
      
   End If

End Sub


Private Sub Cb_TipoLst_Click()
   Call LoadLst(0)
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   Call AddItem(Cb_TipoLst, "(todas)", 0)
   
   For i = 0 To UBound(gTipoRazFin)
   
      If gTipoRazFin(i).id = 0 Or gTipoRazFin(i).Nombre = "" Then
         Exit For
      End If
      
      Call AddItem(Cb_TipoLst, gTipoRazFin(i).Nombre, gTipoRazFin(i).id)
      Call AddItem(Cb_Tipo, gTipoRazFin(i).Nombre, gTipoRazFin(i).id)
      
   Next i
   
   If Cb_TipoLst.ListCount > 1 Then
      Cb_TipoLst.ListIndex = 1
   ElseIf Cb_TipoLst.ListCount > 0 Then
      Cb_TipoLst.ListIndex = 0
   End If
   If Cb_Tipo.ListCount > 0 Then
      Cb_Tipo.ListIndex = 0
   End If
   
   Call LoadLst(0)
   
   If lOper = O_DEFINIR Then
      Bt_DefCuentas.Visible = False
      Me.Caption = "Definir Razones Financieras"
   Else
      Bt_DefCuentas.Visible = True
      Bt_Actualizar.Visible = False
      Bt_New.Visible = False
      Bt_Clear.Visible = False
      Me.Caption = "Configurar Razones Financieras"
      Fr_DetRazon.Enabled = False
      Call SetTxRO(Tx_Nombre, True)
      Cb_Tipo.Locked = True
      Call SetTxRO(Tx_Unidad, True)
      Call SetTxRO(Tx_Numerador, True)
      Call SetTxRO(Tx_Denominador, True)
   End If
   
   If Not ChkPriv(PRV_CFG_EMP) Then
      Bt_New.Enabled = False
      Bt_Actualizar.Enabled = False
      Bt_Del.Enabled = False
      Bt_Clear.Enabled = False
   End If

End Sub
Private Sub LoadLst(ByVal IdRazon As Long)
   Dim Q1 As String
   Dim i As Integer
   Dim Rs As Recordset
   
   Ls_Razon.Clear
   
   Q1 = "SELECT Nombre, IdRazon, RazonFija FROM RazonesFin "
   If ItemData(Cb_TipoLst) > 0 Then
      Q1 = Q1 & " WHERE Tipo = " & ItemData(Cb_TipoLst)
   End If
   Q1 = Q1 & " ORDER BY Nombre "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      If vFld(Rs("RazonFija")) <> 0 Then
         Call CbAddItem(Ls_Razon, "* " & vFld(Rs("Nombre")), vFld(Rs("IdRazon")), False)
      Else
         Call CbAddItem(Ls_Razon, "  " & vFld(Rs("Nombre")), vFld(Rs("IdRazon")), False)
      End If
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   If Ls_Razon.ListCount = 0 Then
      Call ClearRazon
      
   Else

      If IdRazon > 0 Then
      
         For i = 0 To Ls_Razon.ListCount - 1
            If Ls_Razon.ItemData(i) = IdRazon Then
               Ls_Razon.ListIndex = i
               Exit For
            End If
         Next i
      
      Else
         Ls_Razon.ListIndex = 0
         
      End If
      
   End If
   
End Sub
Private Sub Ls_Razon_Click()
   Dim id As Long
   Dim Q1 As String
   Dim Rs As Recordset
   
   id = ItemData(Ls_Razon)
   
   If id <= 0 Then
      Exit Sub
   End If
   
   'fill detalle
   Q1 = "SELECT Nombre, Tipo, UnidadRes, TxtNumerador, TxtDenominador, RazonFija, Glosa FROM RazonesFin"
   Q1 = Q1 & " WHERE IdRazon = " & id
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Tx_Nombre = vFld(Rs("Nombre"), True)
      Call SelItem(Cb_Tipo, vFld(Rs("Tipo")))
      
      Tx_Unidad = vFld(Rs("UnidadRes"), True)
      Tx_Numerador = vFld(Rs("TxtNumerador"), True)
      Tx_Denominador = vFld(Rs("TxtDenominador"), True)
      Tx_Glosa = vFld(Rs("Glosa"), True)
      
      If vFld(Rs("RazonFija")) <> 0 Then
         Bt_Actualizar.Enabled = False
         Bt_Del.Enabled = False
      Else
         Bt_Actualizar.Enabled = True
         Bt_Del.Enabled = True
      End If
      
   End If
   
   Call CloseRs(Rs)
         
End Sub

Private Function Valida() As Boolean
   
   Valida = False
   
   If Trim(Tx_Nombre) = "" Then
      MsgBox1 "Debe ingresar un nombre para la razón financiera.", vbExclamation
      Exit Function
   End If
   
   If ItemData(Cb_Tipo) <= 0 Then
      MsgBox1 "Debe seleccionar un tipo de razón financiera.", vbExclamation
      Exit Function
   End If
      
   
   If Trim(Tx_Unidad) = "" Then
      MsgBox1 "Debe ingresar una unidad para expresar el resultado de la razón financiera.", vbExclamation
      Exit Function
   End If
   
   If Trim(Tx_Numerador) = "" Then
      MsgBox1 "Debe ingresar un texto para representar el numerador de la razón financiera.", vbExclamation
      Exit Function
   End If
   
   If Trim(Tx_Denominador) = "" Then
      If MsgBox1("Falta ingresar un texto para representar el denominador de la razón financiera." & vbCrLf & "Al momento de realizar el cálculo, el denominador se reemplazará por el valor 1." & vbCrLf & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   
   Valida = True

End Function

Private Sub ClearRazon()

   Tx_Nombre = ""
   If Cb_Tipo.ListCount > 0 Then
      Cb_Tipo.ListIndex = 0
   End If
   Tx_Unidad = ""
   Tx_Numerador = ""
   Tx_Denominador = ""
   
End Sub

