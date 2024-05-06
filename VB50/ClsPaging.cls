VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CurReg As Long               'primer registro del rango de registros de la página actual
Public NumReg As Long               'cantidad de registros en la grilla
Public ToRightPressed As Boolean    'para indicar que se presionó Bt_ToRight

Dim Bt_ToLeft As CommandButton
Dim Bt_ToRight As CommandButton

Private Sub Class_Initialize()

   Call Clear

End Sub

Public Sub Clear()

   CurReg = 1
   NumReg = 0

End Sub

Public Sub Init(Frm_Bt_ToLeft As CommandButton, Frm_Bt_ToRight As CommandButton)

   Set Bt_ToLeft = Frm_Bt_ToLeft
   Set Bt_ToRight = Frm_Bt_ToRight
   
   If gDbType = SQL_ACCESS Then
      Bt_ToLeft.Visible = False
      Bt_ToRight.Visible = False
   Else
      Bt_ToLeft.ToolTipText = "Anterior conjunto de registros (" & gPageNumReg & ")"
      Bt_ToRight.ToolTipText = "Siguiente conjunto de registros (" & gPageNumReg & ")"
   End If

End Sub

Public Sub ActivateButtons(ByVal NoRecords As Boolean)

   If gDbType = SQL_ACCESS Then
      Exit Sub
   End If
   
   If CurReg = 1 Then
      Bt_ToLeft.Enabled = False
   Else
      Bt_ToLeft.Enabled = True
   End If

   If (NoRecords Or NumReg < gPageNumReg) And ToRightPressed Then
      Bt_ToRight.Enabled = False
   Else
       Bt_ToRight.Enabled = True
   End If

End Sub

Public Sub ToLeft()

   If gDbType = SQL_ACCESS Then
      Exit Sub
   End If

   CurReg = CurReg - gPageNumReg

   If CurReg < 1 Then
      CurReg = 1
   End If

End Sub

Public Sub ToRight()

   If gDbType = SQL_ACCESS Then
      Exit Sub
   End If

   CurReg = CurReg + NumReg
   ToRightPressed = True

End Sub
