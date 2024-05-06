VERSION 5.00
Begin VB.Form FrmConverMoneda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Convertir Moneda"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "FrmConverMoneda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_Equivalencia 
      Caption         =   "Equi&valencias"
      Height          =   795
      Left            =   5700
      Picture         =   "FrmConverMoneda.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   1500
      TabIndex        =   15
      Top             =   300
      Width           =   3795
      Begin VB.CommandButton Bt_SelFecha 
         Height          =   315
         Left            =   1980
         Picture         =   "FrmConverMoneda.frx":064E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.CommandButton Bt_Convertir 
      Caption         =   "Convertir"
      Height          =   795
      Left            =   5700
      Picture         =   "FrmConverMoneda.frx":06C3
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda Destino"
      ForeColor       =   &H00FF0000&
      Height          =   2115
      Index           =   1
      Left            =   1500
      TabIndex        =   12
      Top             =   3540
      Width           =   3795
      Begin VB.TextBox Tx_FechaEquivDest 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1500
         Width           =   1035
      End
      Begin VB.TextBox Tx_EquivDest 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CommandButton Bt_Copy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         Picture         =   "FrmConverMoneda.frx":0C4B
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Copiar valor"
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox tx_Resultado 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   900
         Width           =   1335
      End
      Begin VB.ComboBox Cb_MonedaDest 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Equiv.:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   22
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Valor:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   5700
      TabIndex        =   7
      Top             =   420
      Width           =   1215
   End
   Begin VB.CommandButton Bt_Cerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   5700
      TabIndex        =   8
      Top             =   780
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda Origen"
      ForeColor       =   &H00FF0000&
      Height          =   2115
      Index           =   0
      Left            =   1500
      TabIndex        =   9
      Top             =   1140
      Width           =   3795
      Begin VB.TextBox Tx_FechaEquivOri 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1500
         Width           =   1035
      End
      Begin VB.TextBox Tx_EquivOri 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox Tx_Valor 
         Height          =   315
         Left            =   1020
         TabIndex        =   3
         Top             =   900
         Width           =   1335
      End
      Begin VB.ComboBox Cb_MonedaOri 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Equiv.:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Valor:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   540
         Width           =   675
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   360
      Picture         =   "FrmConverMoneda.frx":102B
      Top             =   420
      Width           =   780
   End
End
Attribute VB_Name = "FrmConverMoneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cbMonedaOri As ClsCombo
Dim cbMonedaDest As ClsCombo

Dim lResultado As Double
Dim lRc As Integer
Dim lValor As Double

Dim lFormatOri As String
Dim lOper As Integer

Const C_CARACTMON = 2

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub
Private Sub bt_Convertir_Click()
   Dim Fecha As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Wh As String
   Dim ValPesosOri As Double, ValPesosDest As Double
   Dim Valor As Double, Resultado As Double
   Dim i As Integer
   Dim Idx As Integer
   Dim OrderBy As String
   Dim FechaOri As Long, FechaDest As Long
   Dim CaractOri As Integer, CaractDest As Integer
   Dim FechaEquivOri As Long, FechaEquivDest As Long
      
   If cbMonedaOri.ListIndex < 0 Then
      MsgBox1 "Seleccione moneda de origen.", vbExclamation
      Cb_MonedaOri.SetFocus
      Exit Sub
   End If
      
   If cbMonedaDest.ListIndex < 0 Then
      MsgBox1 "Seleccione moneda de destino.", vbExclamation
      Cb_MonedaDest.SetFocus
      Exit Sub
   End If
      
   For i = 0 To UBound(gMonedas)
      If gMonedas(i).id = cbMonedaDest.ItemData Then
         Idx = i
         Exit For
      End If
   Next i

   If cbMonedaOri.ItemData = cbMonedaDest.ItemData Then
      tx_Resultado = Format(vFmt(Tx_Valor), gMonedas(Idx).FormatVenta)
      Exit Sub
   End If
   
   If Trim(Tx_Fecha) = "" And (Val(cbMonedaOri.Matrix(C_CARACTMON)) = MON_VDIA Or Val(cbMonedaOri.Matrix(C_CARACTMON)) = MON_VMES) Then
      MsgBox1 "Para este tipo de moneda debe ingresar una fecha.", vbExclamation
      Tx_Fecha.SetFocus
      Exit Sub
   End If
   
   If Trim(Tx_Valor) = "" Then
      MsgBox1 "Debe ingresar el valor a convertir.", vbExclamation
      Tx_Valor.SetFocus
      Exit Sub
   End If
   
   Valor = vFmt(Tx_Valor)
   Fecha = GetTxDate(Tx_Fecha)
   CaractOri = Val(cbMonedaOri.Matrix(C_CARACTMON))
   CaractDest = Val(cbMonedaDest.Matrix(C_CARACTMON))
   ValPesosOri = 0
   ValPesosDest = 0
   FechaEquivOri = 0
   FechaEquivDest = 0
      
   If CaractOri = MON_NACION Then
      
      FechaOri = Fecha
      ValPesosOri = 1
      Tx_FechaEquivOri = Format(FechaOri, DATEFMT)
      Tx_EquivOri = Format(ValPesosOri, NUMFMT)
            
   Else
   
      If CaractOri = MON_VMES Then
         FechaOri = DateSerial(Year(Fecha), Month(Fecha), 1)
         
      ElseIf CaractOri = MON_VDIA Then
         FechaOri = Fecha
         
      End If
      
      Q1 = "SELECT Fecha, Valor FROM Equivalencia WHERE idMoneda=" & cbMonedaOri.ItemData & " AND Valor > 0 AND Fecha <= " & FechaOri & " ORDER BY Fecha desc"
      
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         ValPesosOri = vFld(Rs("Valor"))
         FechaEquivOri = vFld(Rs("Fecha"))
      End If
      
      Call CloseRs(Rs)
      
      If ValPesosOri = 0 Or FechaEquivOri < FechaOri Then
         MsgBox1 "No ha ingresado equivalencia de la moneda de origen para la fecha especificada. Se utilizará la equivalencia más reciente.", vbExclamation
      End If
      
      Tx_EquivOri = Format(ValPesosOri, NUMFMT)
      Tx_FechaEquivOri = Format(FechaEquivOri, DATEFMT)
               
   End If
   
   If CaractDest = MON_NACION Then
      
      FechaDest = Fecha
      ValPesosDest = 1
      Tx_FechaEquivDest = Format(FechaDest, DATEFMT)
      Tx_EquivDest = Format(ValPesosDest, NUMFMT)
      
      
   Else
   
      If CaractDest = MON_VMES Then
         FechaDest = DateSerial(Year(Fecha), Month(Fecha), 1)
         
      ElseIf CaractDest = MON_VDIA Then
         FechaDest = Fecha
         
      End If
      
      Q1 = "SELECT Fecha, Valor FROM Equivalencia WHERE idMoneda=" & cbMonedaDest.ItemData & " AND Valor > 0 AND Fecha <= " & FechaDest & " ORDER BY Fecha desc"
      
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         ValPesosDest = vFld(Rs("Valor"))
         FechaEquivDest = vFld(Rs("Fecha"))
      End If
      
      Call CloseRs(Rs)
      
      If ValPesosDest = 0 Or FechaEquivDest < FechaDest Then
         MsgBox1 "No ha ingresado equivalencia de la moneda de Destgen para la fecha especificada. Se utilizará la equivalencia más reciente.", vbExclamation
      End If
      
      Tx_EquivDest = Format(ValPesosDest, NUMFMT)
      Tx_FechaEquivDest = Format(FechaEquivDest, DATEFMT)
               
   End If
   
   If ValPesosDest <> 0 Then
      Resultado = Valor * ValPesosOri / ValPesosDest
   End If
   
   tx_Resultado = Format(Resultado, gMonedas(Idx).FormatVenta)
        
End Sub
Private Sub bt_ConvertirOld_Click()
   Dim Fecha As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Wh As String
   Dim ValPesosOri As Double
   Dim i As Integer
   Dim Idx As Integer
   Dim OrderBy As String
   Dim FechaOri As Long, FechaDest As Long
      
   If cbMonedaOri.ListIndex < 0 Then
      MsgBox1 "Seleccione moneda de origen.", vbExclamation
      Cb_MonedaOri.SetFocus
      Exit Sub
   End If
      
   If cbMonedaDest.ListIndex < 0 Then
      MsgBox1 "Seleccione moneda de destino.", vbExclamation
      Cb_MonedaDest.SetFocus
      Exit Sub
   End If
      
   For i = 0 To UBound(gMonedas)
      If gMonedas(i).id = cbMonedaDest.ItemData Then
         Idx = i
         Exit For
      End If
   Next i

   If cbMonedaOri.ItemData = cbMonedaDest.ItemData Then
      tx_Resultado = Format(vFmt(Tx_Valor), gMonedas(Idx).FormatVenta)
      Exit Sub
   End If
   
   If Trim(Tx_Fecha) = "" And (Val(cbMonedaOri.Matrix(C_CARACTMON)) = MON_VDIA Or Val(cbMonedaOri.Matrix(C_CARACTMON)) = MON_VMES) Then
      MsgBox1 "Para este tipo de moneda debe ingresar una fecha.", vbExclamation
      Tx_Fecha.SetFocus
      Exit Sub
   End If
   
   If Trim(Tx_Valor) = "" Then
      MsgBox1 "Debe ingresar el valor a convertir.", vbExclamation
      Tx_Valor.SetFocus
      Exit Sub
   End If
   
   Fecha = GetTxDate(Tx_Fecha)
   OrderBy = ""
   
   If Val(cbMonedaOri.Matrix(C_CARACTMON)) = MON_VMES Or Val(cbMonedaOri.Matrix(C_CARACTMON)) = MON_VDIA Then
      If cbMonedaOri.Matrix(C_CARACTMON) = MON_VMES Then
         Fecha = DateSerial(Year(Fecha), Month(Fecha), 1)
      End If
      Wh = " AND Fecha=" & Fecha
   Else   'dólar
      Wh = " AND Fecha= " & Fecha
      OrderBy = " ORDER BY Fecha desc"
   End If
   If Val(cbMonedaOri.Matrix(C_CARACTMON)) = MON_VMES Or Val(cbMonedaOri.Matrix(C_CARACTMON)) = MON_VDIA Then
      If cbMonedaOri.Matrix(C_CARACTMON) = MON_VMES Then
         Fecha = DateSerial(Year(Fecha), Month(Fecha), 1)
      End If
      Wh = " AND Fecha=" & Fecha
   Else   'dólar
      Wh = " AND Fecha= " & Fecha
      OrderBy = " ORDER BY Fecha desc"
   End If
   
   
   If Val(cbMonedaOri.Matrix(C_CARACTMON)) = MON_NACION Then   'es moneda nacional (Pesos)
   
      If cbMonedaDest.Matrix(C_CARACTMON) = MON_NACION Then    'de pesos a pesos
         tx_Resultado = Format(vFmt(Tx_Valor), gMonedas(Idx).FormatVenta)
      
      Else
      
         Q1 = "SELECT Valor FROM Equivalencia WHERE idMoneda=" & cbMonedaDest.ItemData & Wh & " ORDER BY Fecha desc"
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then
            If vFld(Rs("Valor")) > 0 Then
               tx_Resultado = Format(vFmt(Tx_Valor) / vFld(Rs("Valor")), gMonedas(Idx).FormatVenta)
            End If
            
         ElseIf OrderBy <> "" Then
            
            Call CloseRs(Rs)
            
            Q1 = "SELECT Valor FROM Equivalencia WHERE idMoneda=" & cbMonedaDest.ItemData & " AND Valor > 0 " & OrderBy
            Set Rs = OpenRs(DbMain, Q1)
            
            If Rs.EOF = False Then
               MsgBox1 "No ha ingresado equivalencia de la moneda de destino para la fecha especificada. Se utilizará la equivalencia más reciente.", vbExclamation
               If vFld(Rs("Valor")) > 0 Then
                  tx_Resultado = Format(vFmt(Tx_Valor) / vFld(Rs("Valor")), gMonedas(Idx).FormatVenta)
               End If
            Else
               MsgBox1 "No ha ingresado equivalencia para la moneda de destino.", vbExclamation
               tx_Resultado = ""
            End If
         
         Else
            MsgBox1 "No ha ingresado equivalencia para la moneda de destino.", vbExclamation
            tx_Resultado = ""
         End If
         
         Call CloseRs(Rs)
      End If
      
   ElseIf Val(cbMonedaDest.Matrix(C_CARACTMON)) = MON_NACION Then
  
      Q1 = "SELECT Valor FROM Equivalencia WHERE idMoneda=" & cbMonedaOri.ItemData & Wh
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         tx_Resultado = Format(vFmt(Tx_Valor) * vFld(Rs("Valor")), gMonedas(Idx).FormatVenta)
      Else
         MsgBox1 "No ha ingresado equivalencia para la moneda de origen.", vbExclamation
         tx_Resultado = ""
      End If
      
   Else
   
      Q1 = "SELECT Valor FROM Equivalencia WHERE idMoneda=" & cbMonedaOri.ItemData & Wh
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         ValPesosOri = vFmt(Tx_Valor) * vFld(Rs("Valor"))
         
         Call CloseRs(Rs)
  
         Q1 = "SELECT Valor FROM Equivalencia WHERE idMoneda=" & cbMonedaDest.ItemData & Wh
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then
            tx_Resultado = Format(ValPesosOri / vFld(Rs("Valor")), gMonedas(Idx).FormatVenta)
         Else
            MsgBox1 "No ha ingresado equivalencia para la moneda de destino.", vbExclamation
            tx_Resultado = ""
         End If
   
      End If
   End If
   
   Call CloseRs(Rs)
        
End Sub

Private Sub bt_Copy_Click()
   Clipboard.Clear
   Clipboard.SetText vFmt(tx_Resultado)

End Sub

Private Sub bt_Equivalencia_Click()
   Dim Frm As FrmEquivalencias
   
   Set Frm = New FrmEquivalencias
   Call Frm.FEdit(0)
   Set Frm = Nothing
End Sub

Private Sub bt_OK_Click()
   lRc = vbOK
   lResultado = vFmt(tx_Resultado)
   
   Call SetIniString(gIniFile, "Monedas", "MonedaOri", CbItemData(Cb_MonedaOri))
   Call SetIniString(gIniFile, "Monedas", "MonedaDest", CbItemData(Cb_MonedaDest))

   Unload Me
   
End Sub

Private Sub Bt_SelFecha_Click()
   Dim Fecha As Long
   Dim Frm As FrmCalendar
   
   Set Frm = New FrmCalendar
  
   Call Frm.TxSelDate(Tx_Fecha)
   
   Set Frm = Nothing
End Sub


Private Sub Cb_MonedaOri_Click()
   Dim i As Integer
   
   For i = 0 To UBound(gMonedas)
      If gMonedas(i).id = cbMonedaOri.ItemData Then
         lFormatOri = gMonedas(i).FormatVenta
         Exit For
      End If
   Next i
      
End Sub

Private Sub Form_Load()
   Dim Q1 As String
   Dim IdMonedaOri As Integer, IdMonedaDest As Integer
   
   lRc = vbCancel
   
   Set cbMonedaOri = New ClsCombo
   Call cbMonedaOri.SetControl(Cb_MonedaOri)
   
   Set cbMonedaDest = New ClsCombo
   Call cbMonedaDest.SetControl(Cb_MonedaDest)
   
   Q1 = "SELECT Descrip, idMoneda, Caracteristica FROM Monedas"
   Call cbMonedaOri.FillCombo(DbMain, Q1, -1)
   Call cbMonedaDest.FillCombo(DbMain, Q1, -1)
   cbMonedaDest.SelItem (MON_NACION)
    
   IdMonedaOri = Val(GetIniString(gIniFile, "Monedas", "MonedaOri", "-1"))
   IdMonedaDest = Val(GetIniString(gIniFile, "Monedas", "MonedaDest", "-1"))
 
   If IdMonedaOri >= 0 Then
      Call CbSelItem(Cb_MonedaOri, IdMonedaOri)
   End If
    
   If IdMonedaDest >= 0 Then
      Call CbSelItem(Cb_MonedaDest, IdMonedaDest)
   End If
    
    
   Call SetTxDate(Tx_Fecha, Now)
   
   If lValor <> 0 Then
      Tx_Valor = Format(lValor, lFormatOri)
   End If
   
   If lOper = O_VIEW Then
      Bt_Ok.Visible = False
      Bt_Cerrar.Top = Bt_Ok.Top
   End If
End Sub
Private Sub Tx_Fecha_GotFocus()
   Call DtGotFocus(Tx_Fecha)
End Sub

Private Sub Tx_Fecha_LostFocus()
   Call DtLostFocus(Tx_Fecha)
End Sub

Public Function FSelect(Valor As Double) As Double
   lValor = Valor
   
   Me.Show vbModal
   
   FSelect = lRc
   Valor = lResultado
   
End Function

Public Sub FView(Valor As Double)
   lValor = Valor
   
   lOper = O_VIEW
   
   Me.Show vbModal
      
End Sub
Private Sub Tx_Valor_GotFocus()
   If Trim(Tx_Valor) <> "" Then
      Tx_Valor = vFmt(Tx_Valor)
   End If
End Sub

Private Sub Tx_Valor_KeyPress(KeyAscii As Integer)
   Call KeyDec(KeyAscii)
End Sub

Private Sub Tx_Valor_LostFocus()
   Tx_Valor = Format(Tx_Valor, lFormatOri)
   
End Sub
