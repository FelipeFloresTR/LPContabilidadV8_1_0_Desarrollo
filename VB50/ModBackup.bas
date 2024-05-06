Attribute VB_Name = "ModBackup"
Option Explicit

Private bInitMsg As Boolean

Public Sub ShowMsgBackup()
   Dim DtM As Long, Dt As Long, bShow As Boolean
   Dim Frm As FrmBackup
   
   If bInitMsg Then
      Exit Sub
   End If
   
   If gAppCode.Demo Then
      bInitMsg = True
      Exit Sub
   End If
   
   If Now < W.FStart + TimeSerial(0, 1, 0) Then
      Exit Sub
   End If
   
   DtM = Val(GetIniString(gIniFile, "Config", "MsgBackup", "0"))
   Dt = Int(Now)
   
   bShow = False
   
   If Abs(DtM - Dt) > 14 Then
      bShow = True
   ElseIf Dt > DtM + 5 And Weekday(Dt, vbMonday) = 5 Then
      bShow = True
   End If

   If bShow Then
      Set Frm = New FrmBackup
      Frm.Show vbModal
      Set Frm = Nothing
      Call SetIniString(gIniFile, "Config", "MsgBackup", Dt)
   End If

   bInitMsg = True
   
End Sub
