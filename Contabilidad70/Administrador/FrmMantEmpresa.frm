VERSION 5.00
Begin VB.Form FrmMantEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empresa"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "FrmMantEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   1080
      TabIndex        =   5
      Top             =   420
      Width           =   7755
      Begin VB.TextBox tx_Rut 
         Height          =   315
         Left            =   1560
         MaxLength       =   13
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Tx_NCorto 
         Height          =   315
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   4215
      End
      Begin VB.CommandButton bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   6000
         TabIndex        =   3
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton bt_Cancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   6000
         TabIndex        =   4
         Top             =   660
         Width           =   1275
      End
      Begin VB.CheckBox Ch_NoActivo 
         Caption         =   "No Activo"
         Height          =   255
         Left            =   4680
         TabIndex        =   2
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "R.U.T.:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Corto:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   900
         Width           =   1035
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Index           =   0
      Left            =   240
      Picture         =   "FrmMantEmpresa.frx":000C
      Top             =   540
      Width           =   690
   End
End
Attribute VB_Name = "FrmMantEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lRut As String
Dim lNCorto As String
Dim lId As Long
Dim lRc As Integer
Dim lOper As Integer
Dim lEstado As Integer

Private Function Valida() As Boolean
   Dim Row As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim RUT As String
   
   Valida = False
   
   RUT = Trim(Tx_RUT)
   
   If RUT = "" Then
      MsgBox1 "Debe ingresar el RUT de la empresa", vbExclamation
      Exit Function
   End If
   
   If RUT <> "" And Trim(Tx_NCorto) = "" Then
      MsgBox1 "Debe ingresar nombre corto", vbExclamation
      Exit Function
   End If
   
   If gAppCode.Demo = True Then
      If RUT <> "1-9" And RUT <> "2-7" And RUT <> "3-5" Then
         MsgBox1 "En la versín DEMO sólo puede usar los siguietes RUTs:" & vbCrLf & vbCrLf & "1-9, 2-7 y 3-5", vbExclamation
         Exit Function
      End If
   End If
         
   
   'Ver en BDatos
   Q1 = "SELECT Rut FROM Empresas WHERE Rut='" & vFmtCID(Tx_RUT) & "'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False And lOper = OPER_NEW Then
      MsgBox1 "Ya existe una empresa con este RUT.", vbExclamation
      Call CloseRs(Rs)
      Exit Function
      
   End If
   Call CloseRs(Rs)
   
   'Ver en BDatos
   Q1 = "SELECT idEmpresa FROM Empresas WHERE NombreCorto='" & Tx_NCorto & "'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If (lOper = OPER_NEW) Or (lOper = OPER_EDIT And vFld(Rs("idEmpresa")) <> lId) Then
         MsgBox1 "Ya existe este nombre corto asociado a otra empresa", vbExclamation
         Call CloseRs(Rs)
         Exit Function
      End If
   End If
   Call CloseRs(Rs)
   
   Valida = True
   
End Function

Private Sub bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub bt_OK_Click()
   Dim Q1 As String
   Dim Row As Integer
   Dim Rs As Recordset
   Dim Estado As Integer
   Dim FldArray(1) As AdvTbAddNew_t
   
   If Valida() = False Then
      Exit Sub
   End If
   
   If lOper = OPER_NEW Then
'      Set Rs = DbMain.OpenRecordset("Empresas", dbOpenTable)
'      Rs.AddNew
'      lId = Rs("IdEmpresa")
'      Rs("Rut") = vFmtCID(tx_Rut)
'      Rs.Update
'      Rs.Close
      
      FldArray(0).FldName = "Rut"
      FldArray(0).FldValue = vFmtCID(Tx_RUT)
      FldArray(0).FldIsNum = False
      
      FldArray(1).FldName = "NombreCorto"
      FldArray(1).FldValue = ParaSQL(Tx_NCorto)
      FldArray(1).FldIsNum = False
                  
      lId = AdvTbAddNewMult(DbMain, "Empresas", "IdEmpresa", FldArray)
   
'      lId = AdvTbAddNew(DbMain, "Empresas", "IdEmpresa", "Rut", vFmtCID(Tx_RUT))
   End If
   
   If Not Ch_NoActivo.Visible Then
      Estado = 0
   Else
      Estado = IIf(Ch_NoActivo = 0, 0, 1)
   End If
      
   Q1 = "UPDATE Empresas SET NombreCorto='" & ParaSQL(Tx_NCorto) & "', Estado = " & Estado
   Q1 = Q1 & " WHERE IdEmpresa=" & lId
   Call ExecSQL(DbMain, Q1)
   
   lRut = Tx_RUT
   lNCorto = Tx_NCorto
   lRc = vbOK
   lEstado = Estado
   
   Unload Me
   
End Sub

Public Function FNew(id As Long, RUT As String, NCorto As String) As Integer
   lOper = OPER_NEW
   
   lId = id
   lRut = RUT
   lNCorto = NCorto
   
   Me.Show vbModal
   
   id = lId
   RUT = lRut
   NCorto = lNCorto
   FNew = lRc
   
End Function

Public Function FEdit(id As Long, RUT As String, NCorto As String, Estado As Integer) As Integer
   lOper = OPER_EDIT
   lId = id
   lRut = RUT
   lNCorto = NCorto
   lEstado = Estado
   
   Me.Show vbModal
 
   NCorto = lNCorto
   Estado = lEstado
   FEdit = lRc
   
End Function

Private Sub Form_Load()

   lRc = vbCancel
   
   If lOper = OPER_NEW Then
      Me.Caption = "Nueva empresa"
      Ch_NoActivo.Visible = False
      If lId = 0 And lRut <> "" And lNCorto <> "" Then
         Tx_RUT = FmtCID(lRut)
         Tx_NCorto = lNCorto
      End If
   Else
      Me.Caption = "Modificar Empresa"
      Tx_RUT.Enabled = False
      Tx_RUT = lRut
      Tx_NCorto = lNCorto
      Ch_NoActivo = lEstado

   End If
   
'   If gAppCode.Demo Then
'      Call SetRO(Tx_NCorto, True)
'   End If
   
   
End Sub

Private Sub Tx_Rut_KeyPress(KeyAscii As Integer)
   Call KeyCID(KeyAscii)
End Sub

Private Sub Tx_Rut_LostFocus()
   
   If Tx_RUT = "" Then
      Exit Sub
      
   End If
   
   If vFmtCID(Tx_RUT) = 0 Then
      Tx_RUT = ""
      Tx_RUT.SetFocus
      Exit Sub
   End If
   
'   If Not MsgValidRut(Tx_Rut) Then
'      Tx_Rut.SetFocus
'      Exit Sub
'
'   End If
'
   Tx_RUT = FmtCID(vFmtCID(Tx_RUT))
   
End Sub
Private Sub Tx_RUT_Validate(Cancel As Boolean)
   
   If Tx_RUT = "" Then
      Exit Sub
   End If
   
   If Trim(Tx_RUT) = "0-0" Then
      MsgBox1 "RUT Inválido.", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   
   If Not MsgValidRut(Tx_RUT) Then
      Tx_RUT.SetFocus
      Cancel = True
      Exit Sub
      
   End If
   
   
End Sub
