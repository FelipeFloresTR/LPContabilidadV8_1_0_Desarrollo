VERSION 5.00
Begin VB.Form FrmMantMoneda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Moneda"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "FrmMantMoneda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Tx_DecVenta 
      Height          =   315
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   15
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame Fr_Caract 
      Caption         =   "Característica Moneda"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   1500
      TabIndex        =   10
      Top             =   2400
      Width           =   5595
      Begin VB.OptionButton Op_Moneda 
         Caption         =   "Valor Diario"
         Height          =   255
         Index           =   3
         Left            =   3540
         TabIndex        =   14
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Op_Moneda 
         Caption         =   "Valor Mensual"
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Op_Moneda 
         Caption         =   "Valor Unico"
         Height          =   255
         Index           =   1
         Left            =   3540
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Op_Moneda 
         Caption         =   "Moneda Nacional"
         Height          =   375
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7320
      TabIndex        =   9
      Top             =   780
      Width           =   1035
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   7320
      TabIndex        =   8
      Top             =   420
      Width           =   1035
   End
   Begin VB.TextBox Tx_DecInf 
      Height          =   315
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Tx_Simbolo 
      Height          =   315
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Tx_Descrip 
      Height          =   315
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   3
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox Tx_Id 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   360
      Picture         =   "FrmMantMoneda.frx":000C
      Top             =   480
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Formato de salida"
      Height          =   195
      Index           =   6
      Left            =   4080
      TabIndex        =   18
      Top             =   1980
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingreso equivalencia moneda nacional"
      Height          =   195
      Index           =   5
      Left            =   4080
      TabIndex        =   17
      Top             =   1620
      Width           =   2730
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Decimales salida:"
      Height          =   195
      Index           =   4
      Left            =   1560
      TabIndex        =   16
      Top             =   1980
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Decimales ingreso"
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      Top             =   1620
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "Símbolo:"
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción:"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Índice:"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "FrmMantMoneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lMoneda As Monedas_t
Dim lRc As Integer
Dim lOper As Integer
Dim lOpIdx As Integer

Friend Function FNew(Moneda As Monedas_t) As Integer
   lOper = O_NEW
   
   Me.Show vbModal
   
   Moneda = lMoneda
   FNew = lRc
End Function
Friend Function FEdit(Moneda As Monedas_t) As Integer
   lMoneda = Moneda
   lOper = O_EDIT
   
   Me.Show vbModal
   
   Moneda = lMoneda
   FEdit = lRc
End Function

Private Sub Bt_Cancel_Click()
   Unload Me
End Sub

Private Sub bt_OK_Click()
   Dim Q1 As String
   Dim i As Integer
   Dim Rs As Recordset

   If Trim(Tx_Descrip) = "" Then
      MsgBox1 "Debe ingresar una descripción", vbExclamation
      Tx_Descrip.SetFocus
      Exit Sub
   End If
   
   If Trim(Tx_Simbolo) = "" Then
      MsgBox1 "Debe ingresar un símbolo", vbExclamation
      Tx_Simbolo.SetFocus
      Exit Sub
   End If
   
   If lOper = O_NEW Then
      
      'Obtengo el código
      Q1 = "SELECT Max(idMoneda) as Max FROM Monedas "
      Set Rs = OpenRs(DbMain, Q1)
      lMoneda.Id = vFld(Rs("Max")) + 1
      Call CloseRs(Rs)

      Q1 = "INSERT INTO Monedas (idMoneda) VALUES(" & lMoneda.Id & ")"
      Call ExecSQL(DbMain, Q1)
   End If
   
   Q1 = "UPDATE Monedas SET Descrip='" & ParaSQL(Tx_Descrip) & "'"
   Q1 = Q1 & ", Simbolo='" & ParaSQL(Tx_Simbolo) & "'"
   Q1 = Q1 & ", DecInf=" & vFmt(Tx_DecInf)
   Q1 = Q1 & ", DecVenta=" & vFmt(Tx_DecVenta)
   Q1 = Q1 & ", Caracteristica=" & lOpIdx
   Q1 = Q1 & "  WHERE idMoneda=" & lMoneda.Id
   Call ExecSQL(DbMain, Q1)
   
   lMoneda.DecInf = vFmt(Tx_DecInf)
   lMoneda.DecVenta = vFmt(Tx_DecVenta)
   lMoneda.Descrip = Tx_Descrip
   lMoneda.Simbolo = Tx_Simbolo
   lMoneda.Caract = lOpIdx
   
   For i = 0 To UBound(gMonedas)
      If gMonedas(i).Id = lMoneda.Id Then
         gMonedas(i).DecInf = lMoneda.DecInf
         gMonedas(i).DecVenta = lMoneda.DecVenta
         gMonedas(i).Descrip = lMoneda.Descrip
         gMonedas(i).Simbolo = lMoneda.Simbolo
         gMonedas(i).Caract = lMoneda.Caract
         
         If gMonedas(i).DecInf > 0 Then
            gMonedas(i).FormatInf = NUMFMT & "." & String(gMonedas(i).DecInf, "0")
         Else
            gMonedas(i).FormatInf = NUMFMT
         End If
         
         If gMonedas(i).DecVenta > 0 Then
            gMonedas(i).FormatVenta = NUMFMT & "." & String(gMonedas(i).DecVenta, "0")
         Else
            gMonedas(i).FormatVenta = NUMFMT
         End If
         
         Exit For
      End If
   Next i
         
   lRc = vbOK
   
   Unload Me
   
End Sub

Private Sub Form_Load()
   lRc = vbCancel
   
   If lOper = O_NEW Then
      Me.Caption = "Agregar moneda"
   Else
      Me.Caption = "Modificar moneda"
   End If
   
   Call LoadAll
'   Call EnableForm(Me, gEmpresa.FCierre = 0)
   Call SetTxRO(Tx_Id, True)
      
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Cod As Integer
   
   If lOper = O_NEW Then
      Exit Sub
   End If
   
   Q1 = "SELECT * FROM Monedas WHERE idMoneda=" & lMoneda.Id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      Tx_Id = vFld(Rs("idMoneda"))
      Tx_Descrip = vFld(Rs("Descrip"), True)
      Tx_Simbolo = vFld(Rs("Simbolo"), True)
      Tx_DecInf = Format(vFld(Rs("DecInf")), NUMFMT)
      Tx_DecVenta = Format(vFld(Rs("DecVenta")), NUMFMT)
      Op_Moneda(vFld(Rs("Caracteristica"))).Value = True
      
      If vFld(Rs("EsFijo")) <> 0 Then
         Call SetRO(Tx_Descrip, True)
         Call SetRO(Tx_Simbolo, True)
         Call SetRO(Tx_DecInf, True)
         Call SetRO(Tx_DecVenta, True)
         Call SetRO(Tx_DecVenta, True)
         
         Fr_Caract.Enabled = False
      End If
   End If
   
   Call CloseRs(Rs)
         
End Sub

Private Sub Op_Moneda_Click(Index As Integer)
   lOpIdx = Index
   
End Sub

Private Sub Tx_DecInf_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub

Private Sub Tx_DecVenta_KeyPress(KeyAscii As Integer)
   Call KeyNum(KeyAscii)
End Sub
