VERSION 5.00
Begin VB.Form FrmUnlock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desbloquear Procesos"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "FrmUnlock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tx_Info 
      Height          =   915
      Left            =   1500
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "FrmUnlock.frx":000C
      Top             =   4200
      Width           =   3855
   End
   Begin VB.CommandButton Bt_Unlock 
      Caption         =   "Desbloquear..."
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   3720
      Width           =   1395
   End
   Begin VB.CommandButton Bt_Close 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   660
      Width           =   1395
   End
   Begin VB.ListBox Ls_PC 
      Height          =   3375
      Left            =   1500
      TabIndex        =   0
      Top             =   660
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   480
      Picture         =   "FrmUnlock.frx":0076
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Equipos:"
      Height          =   195
      Left            =   1500
      TabIndex        =   4
      Top             =   420
      Width           =   615
   End
End
Attribute VB_Name = "FrmUnlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Close_Click()
   Unload Me
End Sub

Private Sub Bt_Unlock_Click()
   Dim Msg As String, PC As String, Q1 As String
   Dim Rc As Long

   If Ls_PC.ListIndex < 0 Then
      MsgBox1 "Seleccione algún equipo.", vbExclamation
      Exit Sub
   End If

   PC = Trim(Ls_PC)
   If PC = "" Then
      Ls_PC.RemoveItem Ls_PC.ListIndex
      Exit Sub
   End If

   Msg = "¡ ATENCION !" & vbLf & vbLf
   Msg = Msg & "Desbloquear los procesos bloqueados por un equipo que está funcionando puede producir serios problemas en el sistema." & vbLf
   Msg = Msg & "Sólo debe hacerlo si el equipo está apagado." & vbLf & vbLf
   Msg = Msg & "¿Desea desbloquear los procesos del equipo " & PC & " de todas formas?"

   If MsgBox1(Msg, vbExclamation Or vbYesNo Or vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If

   MousePointer = vbHourglass
   DoEvents

   Q1 = "DELETE FROM LockAction WHERE PcName='" & PC & "'"
   Rc = ExecSQL(DbMain, Q1)
   
   Call LoadAll

   MousePointer = vbDefault

End Sub

Private Sub Form_Load()

   If gAppCode.Demo Then
      Me.Caption = Me.Caption & " - DEMO"
   End If

   Bt_Unlock.Visible = ChkPriv(PRV_ADM_SIS)

   Tx_Info = ReplaceStr(Tx_Info, "eeee", gEmpresa.NombreCorto)
   Tx_Info = ReplaceStr(Tx_Info, "aaaa", gEmpresa.Ano)

   Call LoadAll

End Sub

Private Sub LoadAll()
   Dim Q1 As String
      
   Ls_PC.Clear
   
   Q1 = "SELECT PcName, Count(*) as N FROM LockAction GROUP BY PcName ORDER BY PcName"
   Call FillCombo(Ls_PC, DbMain, Q1, -2)
   
   Bt_Unlock.Enabled = Abs(Ls_PC.ListCount > 0)
   
End Sub

Private Sub Ls_PC_Click()
   Beep
End Sub
