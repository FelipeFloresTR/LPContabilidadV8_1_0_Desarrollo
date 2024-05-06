Attribute VB_Name = "Calendar"
Option Explicit

' Funciones que sirven para manejar fechas en TextBox
' Son consistentes para usarlas con Calendar.frm

' 31 oct 2018: se cambia el uso de .Tag por .DataField por uso del .Tag en PrtForm y PrtForm2

Public Const DATEFMT As String = "dd mmm yyyy"
Public Const EDATEFMT As String = "dd\/mm\/yyyy"   ' Para GetDate(),  12 jul 2021: se agrega el \ para que Windows no cambie el / por lo que tiene definido para la fecha
Public Const SDATEFMT As String = "dd\/mm\/yy"      ' 12 jul 2021: se agrega el \ para que Windows no cambie el / por lo que tiene definido para la fecha
Public Const HOURFMT As String = "hh:mm"
Public Const HOURFMTAP As String = "Hh:mm A/P"

' Funciones para fechas en TextBox
' Usar con SetTxDate y GetTxDate
Public Function DtGotFocus(Tx_Date As Control) As Long
   Dim LDate As Long, bMod As Boolean
      
   bMod = Tx_Date.DataChanged
   
   If Trim(Tx_Date.DataField) = "" Then
      LDate = VFmtDate(Tx_Date)
   Else
      LDate = Val(Tx_Date.DataField)
   End If
   
   DtGotFocus = LDate
   
   If Tx_Date.Enabled = False Or Tx_Date.Locked = True Then
      Exit Function
   End If

   If LDate > 0 Then
      Tx_Date = Format(LDate, EDATEFMT)
   Else
      Tx_Date = ""
   End If
   
   Tx_Date.DataChanged = bMod
   
End Function
'reconoce fechas de la forma "dd/mm/yyyy", en que "/" puede
'ser cualquier caracter. Reconocimiento posicional.Formatea
'como "dd mmm yyyy"
Public Function DtLostFocus(Tx_Date As Control, Optional DispMsg As Integer = True, Optional bMonth As Boolean = 0) As Long
   Dim Dt As Long
   Dim TxtDate As String, bMod As Boolean
      
   bMod = Tx_Date.DataChanged
      
   If Tx_Date.Enabled = False Or Tx_Date.Locked = True Then
      DtLostFocus = Val(Tx_Date.DataField)
      Exit Function
   End If
            
   TxtDate = Tx_Date
   
   If Trim(TxtDate) = "" Then
      Tx_Date.DataField = ""
      Dt = 0
      
      'DtLostFocus = 0
   Else
      Dt = GetDate(TxtDate, "dmy")
      
      If bMonth Then
         Call SetTxMonth(Tx_Date, Dt)
      Else
         Call SetTxDate(Tx_Date, Dt)
      End If
         
      'DtLostFocus = ValFmtDate(TxtDate, DispMsg)

      'If Tx_Date.Text <> TxtDate Then
      '   Tx_Date.Text = TxtDate
      'End If
   End If
   
   Tx_Date.DataChanged = bMod
   DtLostFocus = Dt
         
End Function
'reconoce fechas de la forma "dd/mm/yyyy", en que "/" puede
'ser cualquier caracter. Reconocimiento posicional.Formatea
'como "dd mmm yyyy"
Public Function ValFmtDate(TxtDate As String, Optional DispMsg As Integer = True) As Long
   Dim SYear As String
   Dim SMonth As String
   Dim SDay As String
   Dim SDate As String
   Dim LDate As Long
      
   SDate = Trim(TxtDate)
   
   LDate = GetDate(SDate)
   If LDate <= 0 And DispMsg Then
      MsgBox1 "Fecha inválida.", vbExclamation + vbOKOnly
      ValFmtDate = 0
   Else
      ValFmtDate = LDate
      TxtDate = Format(LDate, DATEFMT)
   End If
   
End Function
Public Function FmtDate(ByVal Fecha As Double, Optional ByVal Fmt As String = DATEFMT) As String

   If Fecha > 0 Then
      FmtDate = Format(Fecha, Fmt)
   Else
      FmtDate = ""
   End If
   
End Function
'reconoce fechas y retorna el long asociado
' No siempre funciona bien. Mejor usar SetTxDate y GetTxDate
Public Function VFmtDate(ByVal SDate As String) As Long
   Dim SYear As String
   Dim SMonth As String
   Dim SDay As String
   Dim Dt As Date
   
   On Error Resume Next
   
   Debug.Print "*** La función VFmtDate no siempre funciona bien ***"
   
   If Trim(SDate) = "" Then
      VFmtDate = 0
   Else
      Dt = Format(SDate)
      VFmtDate = Int(Dt)
   End If

   On Error GoTo 0
End Function

' Estas funciones son para ser usadas con Calendar.frm
Public Sub InitTxDate(Tx As TextBox, ByVal Dt As Double)

   Call SetTxDate(Tx, Dt)
   Tx.DataChanged = False

End Sub

' Estas funciones son para ser usadas con Calendar.frm
Public Sub SetTxDate(Tx As TextBox, ByVal Dt As Double, Optional ByVal Fmt As String = DATEFMT)

   If Dt > 0 Then
      
      Tx.DataField = Int(Dt)
      Tx = Format(Dt, Fmt)
   Else
      Tx.DataField = ""
      Tx = ""
   End If

   If Tx.Enabled = True And Tx.Locked = False And Tx.MaxLength = 0 Then
      Tx.MaxLength = Len(Fmt) + 1
   End If

End Sub
' Estas funciones son para ser usadas con Calendar.frm
Public Sub SetTxMonth(Tx As TextBox, ByVal Dt As Double, Optional ByVal Fmt As String = "mmm yyyy")

   If Dt > 0 Then
      If Day(Dt) <> 1 Then
         Dt = DateSerial(Year(Dt), Month(Dt), 1)
      End If
      
      Tx.DataField = Int(Dt)
      Tx = Format(Dt, Fmt)
   Else
      Tx.DataField = ""
      Tx = ""
   End If

End Sub

Public Function GetTxDate(Tx As TextBox, Optional ByVal bOblig As Boolean = 0) As Long

   If Trim(Tx) = "" Then
      GetTxDate = 0
      Call ColFilled(Tx, bOblig)
   ElseIf Tx.DataField <> "" Then
      GetTxDate = Val(Tx.DataField)
   Else
      GetTxDate = VFmtDate(Tx)
   End If

End Function
Public Function GetTxMonth(Tx As TextBox) As Long
   Dim Mon As Long

   If Trim(Tx) = "" Then
      Mon = 0
   ElseIf Tx.DataField <> "" Then
      Mon = Val(Tx.DataField)
   Else
      Mon = VFmtDate(Tx)
   End If

   If Mon > 0 Then
      Mon = DateSerial(Year(Mon), Month(Mon), 1)
   End If

   GetTxMonth = Mon

End Function

Public Function CalcEdad(ByVal FNacim As Double, Optional ByVal Ahora As Double = 0) As Integer
   Dim Anos As Integer

   If FNacim <= 1000 Then
      CalcEdad = 0
      Exit Function
   End If

   If Ahora = 0 Then
      Ahora = Int(Now)
   End If
   
   Anos = Year(Ahora) - Year(FNacim) - 1

   If Format(Ahora, "mmdd") >= Format(FNacim, "mmdd") Then
      Anos = Anos + 1
   End If

   CalcEdad = Anos
   
End Function
Public Function DateFilled(Tx As TextBox, ByVal Msg As String) As Boolean

   If GetTxDate(Tx) <= 0 Then
      Call ColFilled(Tx, True)
      MsgBox1 Msg, vbExclamation
      Tx.SetFocus
      Exit Function
   Else
      Call ColFilled(Tx)
      DateFilled = True
   End If

End Function

#If NOCALEND <> 1 Then
Public Sub ShowCalendar(Tx As TextBox)
   Dim Frm As FrmCalendar
     
   Set Frm = New FrmCalendar
   
   Call Frm.TxSelDate(Tx)
   
   Set Frm = Nothing

End Sub
#End If
