VERSION 5.00
Begin VB.Form FrmActFijoInfoAdic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Información Adicional Activo Fijo"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   8460
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8460
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información Adicional"
      Height          =   2295
      Left            =   360
      TabIndex        =   10
      Top             =   1740
      Width           =   7695
      Begin VB.CommandButton Bt_FechaProy 
         Height          =   315
         Left            =   4020
         Picture         =   "FrmActFijoInfoAdic.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1560
         Width           =   230
      End
      Begin VB.TextBox Tx_FechaProy 
         Height          =   315
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   2
         Top             =   1560
         Width           =   1395
      End
      Begin VB.TextBox Tx_NombreProy 
         Height          =   315
         Left            =   2640
         MaxLength       =   60
         TabIndex        =   1
         Top             =   960
         Width           =   4875
      End
      Begin VB.TextBox Tx_PatenteRol 
         Height          =   315
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   0
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(según proceda)"
         Height          =   195
         Left            =   5700
         TabIndex        =   14
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Proyecto Inversion:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1620
         Width           =   1860
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Proyecto Inversión:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1020
         Width           =   1965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Placa Patente, Rol o Inscripción:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Activo Fijo"
      Height          =   1035
      Left            =   1620
      TabIndex        =   7
      Top             =   420
      Width           =   6435
      Begin VB.TextBox Tx_Descrip 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   8
         Top             =   420
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   360
      Picture         =   "FrmActFijoInfoAdic.frx":030A
      ScaleHeight     =   630
      ScaleWidth      =   825
      TabIndex        =   6
      Top             =   480
      Width           =   885
   End
End
Attribute VB_Name = "FrmActFijoInfoAdic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lIdActFijo As Long
Dim lDescActFijo As String
Dim lPatenteRol As String
Dim lNombreProy As String
Dim lFechaProy As Long
Dim lDepLey21210Inst As Boolean

Dim lRc As Integer
Dim lOper As Integer

Public Function FEdit(ByVal IdActFijo As Long, ByVal DescActFijo As String, ByVal DepLey21210Inst As Boolean, PatenteRol As String, NombreProy As String, FechaProy As Long) As Integer

   lIdActFijo = IdActFijo
   lDescActFijo = DescActFijo
   lDepLey21210Inst = DepLey21210Inst
   lPatenteRol = PatenteRol
   lNombreProy = NombreProy
   lFechaProy = FechaProy
   
   lOper = O_EDIT
   
   Me.Show vbModal
   
   If lRc = vbOK Then
      PatenteRol = lPatenteRol
      NombreProy = lNombreProy
      FechaProy = lFechaProy
   End If
   
   FEdit = lRc
End Function

Public Sub FView(ByVal IdActFijo As Long)

   lIdActFijo = IdActFijo
   lOper = O_VIEW
   
   Me.Show vbModal
   
End Sub

Private Sub Bt_Cancel_Click()
   lRc = vbCancel
   
   Unload Me
End Sub

Private Sub Bt_FechaProy_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FechaProy)
      
   Set Frm = Nothing
   
End Sub

Private Sub bt_OK_Click()
   
   If Valida Then
   
      Call SaveAll
      
      lRc = vbOK
      Unload Me
   End If
   
End Sub

Private Sub Form_Load()

   If lOper = O_VIEW Then
   
      Call SetRO(Tx_PatenteRol, True)
      Call SetRO(Tx_NombreProy, True)
      Call SetRO(Tx_FechaProy, True)
      
      Bt_OK.Visible = False
      Bt_Cancel.Caption = "Cerrar"
      
   End If
   
   Tx_Descrip = lDescActFijo
   Tx_PatenteRol = lPatenteRol
   Tx_NombreProy = lNombreProy
   Call SetTxDate(Tx_FechaProy, lFechaProy)
   
End Sub
Private Sub Tx_FechaProy_GotFocus()
   Call DtGotFocus(Tx_FechaProy)
End Sub

Private Sub Tx_FechaProy_LostFocus()
   Dim Fecha As Long
   
   If Trim$(Tx_FechaProy) = "" Then
      Exit Sub
   End If
   
   Fecha = GetTxDate(Tx_FechaProy)
   
   Call DtLostFocus(Tx_FechaProy)
     
End Sub

Private Sub Tx_FechaProy_KeyPress(KeyAscii As Integer)
   Call KeyDate(KeyAscii)
   
End Sub
Private Sub Bt_Fecha_Click()
   Dim Frm As FrmCalendar

   Set Frm = New FrmCalendar
   Call Frm.TxSelDate(Tx_FechaProy)
      
   Set Frm = Nothing
   
End Sub

Private Function Valida() As Boolean
   Dim Fecha As Long
   
   Valida = False
   
'   If Trim(Tx_PatenteRol) = "" Then
'      MsgBox1 "Falta ingresar Patente, Rol o Inscripción, según proceda.", vbExclamation
'      Exit Function
'   End If
'
   If Not lDepLey21210Inst Then
      Valida = True
      Exit Function
   End If
   
   Fecha = GetTxDate(Tx_FechaProy)
   
   If Trim(Tx_NombreProy) = "" Then
      MsgBox1 "Dado que este bien está acogido a la Depreciación Ley 21210, Instantánea e Inmediata, debe ingresar un Nombre de Proyecto.", vbExclamation
      Exit Function
   End If
   
   '(fecha debe ser entre 01/10/2019 al 31/12/2021)
   If Fecha < DateSerial(2019, 10, 1) Or Fecha > DateSerial(2021, 12, 31) Then
      MsgBox1 "Dado que este bien está acogido a la Depreciación Ley 21210, Instantánea e Inmediata, debe ingresar una Fecha de Proyecto entre el 01/10/2019 y el 31/12/2021.", vbExclamation
      Exit Function
   End If
   
   Valida = True
   
End Function

Private Sub SaveAll()
   Dim Q1 As String
   
   lPatenteRol = Tx_PatenteRol
   lNombreProy = Tx_NombreProy
   lFechaProy = GetTxDate(Tx_FechaProy)

   Q1 = "UPDATE MovActivoFijo SET "
   Q1 = Q1 & "  PatenteRol = '" & ParaSQL(lPatenteRol) & "'"
   Q1 = Q1 & ", NombreProy = '" & ParaSQL(lNombreProy) & "'"
   Q1 = Q1 & ", FechaProy = " & lFechaProy
   Q1 = Q1 & " WHERE IdActFijo = " & lIdActFijo
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
   
End Sub
