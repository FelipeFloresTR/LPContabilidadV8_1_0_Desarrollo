VERSION 5.00
Begin VB.Form FrmPrtCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir Cheque"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "FrmPrtCheque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   360
      Picture         =   "FrmPrtCheque.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   23
      Top             =   540
      Width           =   585
   End
   Begin VB.ComboBox Cb_Tipo 
      Height          =   315
      Left            =   2580
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Cheque"
      Height          =   3075
      Left            =   1380
      TabIndex        =   13
      Top             =   1020
      Width           =   5895
      Begin VB.TextBox Tx_NumEgreso 
         Height          =   315
         Left            =   1140
         TabIndex        =   10
         Top             =   2460
         Width           =   1395
      End
      Begin VB.TextBox Tx_Ciudad 
         Height          =   315
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   2235
      End
      Begin VB.CheckBox Ch_PrtComp 
         Caption         =   "Imprimir Comprobante"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   2100
         Width           =   2295
      End
      Begin VB.TextBox Tx_Valor 
         Height          =   315
         Left            =   1140
         TabIndex        =   8
         Top             =   2040
         Width           =   1395
      End
      Begin VB.TextBox Tx_Banco 
         Height          =   315
         Left            =   3360
         TabIndex        =   7
         Top             =   1620
         Width           =   2355
      End
      Begin VB.TextBox Tx_NumCheque 
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Top             =   1620
         Width           =   1395
      End
      Begin VB.TextBox Tx_Ref 
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox Tx_Nombre 
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   780
         Width           =   4575
      End
      Begin VB.TextBox Tx_Fecha 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Top             =   360
         Width           =   1155
      End
      Begin VB.CommandButton Bt_SelFecha 
         Height          =   315
         Left            =   2280
         Picture         =   "FrmPrtCheque.frx":0433
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Egreso:"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   21
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad:"
         Height          =   195
         Index           =   6
         Left            =   2880
         TabIndex        =   20
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   19
         Top             =   2100
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Index           =   4
         Left            =   2760
         TabIndex        =   18
         Top             =   1680
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Cheque:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ref.:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   1260
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7500
      TabIndex        =   12
      Top             =   720
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Prt 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   7500
      TabIndex        =   11
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo papel:"
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   480
      Width           =   1035
   End
End
Attribute VB_Name = "FrmPrtCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEditable As Boolean
Dim lNumCheque As Long
Dim lPrtGrid As ClsPrtFlxGrid
Dim lFecha As Long
Dim lNombre As String
Dim lRef As String
Dim lBanco As String
Dim lValor As Double
Dim lNumEgreso As Long
Dim lRelBancoNumCheque As Boolean
Dim lSelPrt As Boolean

Dim lInLoad As Boolean


'Imprime y retorna la fecha del cheque que seleccionó el usuario

'Editable: indica si se permite que el usuario edite los campos
'NroCheque: recibe N° del cheque, suponiendo que se lleva la cuenta de en que cheque vamos y retorna el N° del cheque, por si el usuario lo modificó
'Grid: grilla comprobante, si viene entonces se muestra checkbox de impresión de comprobante
'RelBancoNumCheque: si cambia el nombre del banco, dejar NumCheque en blanco
'Retorna: Num Cheque
Public Function FPrint(Editable As Boolean, PrtGrid As ClsPrtFlxGrid, Optional ByVal NumCheque As Long = 0, Optional ByVal Fecha As Long = 0, Optional ByVal Nombre As String = "", Optional ByVal Ref As String = "", Optional ByVal Banco As String = "", Optional ByVal Valor As Double = 0, Optional ByVal NumEgreso As Long = 0, Optional ByVal RelBancoNumCheque As Boolean = False, Optional ByVal SelPrt As Boolean = True) As Long
   
   lEditable = Editable
   lNumCheque = NumCheque
   Set lPrtGrid = PrtGrid
   lFecha = Fecha
   lNombre = Nombre
   lRef = Ref
   lBanco = Banco
   lValor = Valor
   lNumEgreso = NumEgreso
   lRelBancoNumCheque = RelBancoNumCheque
   lSelPrt = SelPrt

   Me.Show vbModal
   
   FPrint = lNumCheque
End Function

Private Sub Bt_Cancel_Click()
   lNumCheque = 0
   Unload Me
End Sub

Private Sub bt_Prt_Click()
   Dim Titulos(0) As String

   If Not Valida() Then
      Exit Sub
   End If

   If Ch_PrtComp <> 0 Then
      Set gPrtCheques.PrtGrid = lPrtGrid
   Else
      Set gPrtCheques.PrtGrid = Nothing
   End If
   
   gPrtCheques.TipoPapel = CbItemData(Cb_Tipo)
   
   gPrtCheques.NumCheque = vFmt(Tx_NumCheque)
   lNumCheque = vFmt(Tx_NumCheque)
   gPrtCheques.Fecha = GetTxDate(Tx_Fecha)
   gPrtCheques.NominativoA = Trim(Tx_Nombre)
   gPrtCheques.Ref = Trim(Tx_Ref)
   gPrtCheques.Banco = Trim(Tx_Banco)
   gPrtCheques.Valor = vFmt(tx_Valor)
   gPrtCheques.Lugar = Trim(Tx_Ciudad)
   gPrtCheques.NumEgreso = Trim(Tx_NumEgreso)
   
   If lSelPrt Then
      If Not PrepararPrt(FrmMain.Cm_PrtDlg) Then
         Exit Sub
      End If
   End If
   
   Call gPrtCheques.PrintCheque(Printer)
   
   Call SetIniString(gIniFile, "Cheques", "Ciudad", Tx_Ciudad)
   Call SetIniString(gIniFile, "Cheques", "Banco", Tx_Banco)
   Call SetIniString(gIniFile, "Cheques", "NumEgreso", Tx_NumEgreso)
   Call SetIniString(gIniFile, "Cheques", "TipoPapel", CbItemData(Cb_Tipo))

   Unload Me
   
End Sub


Private Sub Form_Load()
   Dim Tipo As Long
   
   lInLoad = True

   Call FrmEnable
   
   Call CbAddItem(Cb_Tipo, "Hoja carta", CHEQUE_CARTA)
   Call CbAddItem(Cb_Tipo, "Papel continuo", CHEQUE_CONTINUO)
   Cb_Tipo.ListIndex = 0
   
   Tipo = Val(GetIniString(gIniFile, "Cheques", "TipoPapel", CHEQUE_CARTA))
   Call CbSelItem(Cb_Tipo, Tipo)
   
   Tx_NumCheque = IIf(lNumCheque > 0, lNumCheque, " ")
   If lFecha > 0 Then
      Call SetTxDate(Tx_Fecha, lFecha)
   End If
   Tx_Nombre = lNombre
   Tx_Ref = lRef
   
   Tx_Banco = lBanco
   If Tx_Banco = "" Then
      Tx_Banco = GetIniString(gIniFile, "Cheques", "Banco", "")
   End If
   
   tx_Valor = Format(lValor, BL_NUMFMT)
   
   Tx_NumEgreso = IIf(lNumEgreso > 0, lNumEgreso, "")
   
   If Tx_NumEgreso = "" Then
      Tx_NumEgreso = GetIniString(gIniFile, "Cheques", "NumEgreso", "")
      If Tx_NumEgreso <> "" Then
         Tx_NumEgreso = Val(Tx_NumEgreso) + 1
      End If
   End If
   
   Tx_Ciudad = GetIniString(gIniFile, "Cheques", "Ciudad", "")
   
   If lPrtGrid Is Nothing Then
      Ch_PrtComp.Enabled = False
   Else
      Ch_PrtComp = 1
   End If
   
   lInLoad = False

End Sub

Private Sub FrmEnable()

   Call SetRO(Tx_Fecha, Not lEditable)
   Bt_SelFecha.Enabled = lEditable
   Call SetRO(Tx_Nombre, Not lEditable)
   'Call SetRO(Tx_Ref, Not lEditable)
   Call SetRO(Tx_NumCheque, Not lEditable)
   Call SetRO(Tx_Banco, Not lEditable)
   Call SetRO(tx_Valor, Not lEditable)
   Call SetRO(Tx_NumEgreso, Not lEditable)
   
   
End Sub

Private Function Valida() As Boolean

   Valida = False
   
   If GetTxDate(Tx_Fecha) = 0 Then
      MsgBox1 "Falta ingresar la fecha de emisión del cheque.", vbExclamation
      Tx_Fecha.SetFocus
      Exit Function
   End If
   
   If Trim(Tx_Nombre) = "" Then
      MsgBox1 "Falta ingresar a nombre de quién debe ser emitido el cheque.", vbExclamation
      Tx_Nombre.SetFocus
      Exit Function
   End If
   
   If vFmt(Tx_NumCheque) = 0 Then
      MsgBox1 "Falta ingresar el número o serie del cheque.", vbExclamation
      Tx_NumCheque.SetFocus
      Exit Function
   End If
   
   If Trim(Tx_Banco) = "" Then
      MsgBox1 "Falta ingresar el nombre del banco.", vbExclamation
      Tx_Banco.SetFocus
      Exit Function
   End If
   
   If vFmt(tx_Valor) = 0 Then
      MsgBox1 "Falta ingresar el valor del cheque.", vbExclamation
      tx_Valor.SetFocus
      Exit Function
   End If
   
   Valida = True
End Function

Private Sub Tx_Banco_Change()
   If lInLoad Then
      Exit Sub
   End If
   
   If lRelBancoNumCheque Then
      Tx_NumCheque = ""
   End If
End Sub

Private Sub Tx_Valor_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub
