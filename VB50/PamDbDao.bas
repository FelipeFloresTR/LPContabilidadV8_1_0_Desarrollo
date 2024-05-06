Attribute VB_Name = "PamDbDao"
' Funciones de Bases de Datos para MS Access - Database - DATACON=1
Option Explicit

'
' Abre un Recordset para consulta
'
Public Function OpenRsDao(Db As Database, ByVal Qry As String, Optional ByVal bErrMsg As Boolean = True, Optional ByVal RsType As Integer = dbOpenSnapshot, Optional ByVal Mrk As String = "", Optional ByVal RsOption As Integer = 0, Optional bRepApostr As Boolean = 1) As dao.Recordset
   Dim Rs As dao.Recordset, ConnStr As String, Errno As Long

   Set OpenRsDao = Nothing

   If Trim(Qry) = "" Then
      Exit Function
   End If
    
   If bRepApostr Then
      Qry = ReplaceStr(Qry, "''", "NULL")
   End If
   
   On Error Resume Next
   Set OpenRsDao = Nothing

   'Set Rs = Db.OpenRecordset(Qry, dbOpenForwardOnly, dbConsistent, dbReadOnly) ' , dbOptimistic )
   
   Err.Clear
   SqlErr = 0
   SqlError = ""
   
   ConnStr = Db.Connect
   dao.Errors.Refresh
   
   Set Rs = Db.OpenRecordset(Qry, RsType, dbConsistent Or RsOption, dbReadOnly)  ' , dbOptimistic )
   
   SqlErr = Err
   SqlError = Error
      
   If SqlErr Then
      SqlError = SqlError & vbTab & GetDbErr(Db, Errno)
   
      If Errno <> 0 Then
         SqlErr = Errno
      End If
   
      Call AddLog("OpenRs: Error " & SqlErr & ", " & SqlError & " [" & Qry & "] [Dec=" & Format(1234.56, DBLFMT2) & "-" & True & "]")
      If bErrMsg Then
         MsgBox1 SqlError & vbLf & "[" & Qry & "]", vbExclamation
         'MsgBox1 "Error " & SqlErr & ", " & SqlError & vbLf  & vbLf & "[" & Qry & "]", vbExclamation
      End If
      Set Rs = Nothing
      Exit Function
   ElseIf gDebug = 579 Then
      Call AddDebug("OpenRs: OK Qry=[" & Qry & "]")
   End If
   
   If Rs Is Nothing Then
      Exit Function
   End If
   
   Call PamDb.AddRs(Rs, Qry, Mrk) ' *** PARA DEBUG
   
   Set OpenRsDao = Rs
   
End Function

Public Function ExecSQLDao(Db As Database, ByVal Qry As String, Optional ByVal bErrMsg As Boolean = True, Optional ByVal nTry As Byte = 0) As Long
   Dim Rc As Long, nRec As Long
   Dim LogErr As String, Errno As Long
   Dim Tm As Double
   Dim ConnStr As String
   
   ExecSQLDao = -1

   If Trim(Qry) = "" Or Db Is Nothing Then
      Exit Function
   End If
   
   Qry = ReplaceStr(Qry, "''", "NULL")
   Qry = ReplaceStr(Qry, Chr(164), "''") ' 15 dic 2017: para que sea más estándar

   Tm = Now

   On Error Resume Next
   Err.Clear
   SqlnRec = -1

   dao.Errors.Refresh
   ConnStr = Db.Connect
   
   nRec = -1
   Db.Execute Qry ' Access
   nRec = Db.RecordsAffected
   
   SqlnRec = nRec
   
   If nRec = -1 And Err = 0 Then
      nRec = 1  ' para que no sea 0
   End If
   
   If Errno <> 0 Then
      SqlErr = Errno
   Else
      SqlErr = Err
   End If
   
   SqlError = Error & LogErr
   ExecSQLDao = nRec
   
   If SqlErr Then
      If UCase(Left(Qry, 5)) <> "DROP " Then
         SqlError = SqlError & GetDbErr(Db, Errno)
         
         If SqlError = "" Then
            SqlError = "Error"
         End If
         
         Call AddLog("ExecSql: bMsg=" & bErrMsg & ", " & LogErr & "; Error " & SqlErr & ", " & SqlError & "; [" & Qry & "] [Dec=" & Format(1234.56, DBLFMT2) & "-" & True & "]")
         
         If bErrMsg Then
            MsgBox1 LogErr & vbCrLf & SqlError & vbLf & "[" & Qry & "]", vbExclamation
            'MsgBox1 "Error " & SqlErr & ", " & SqlError & NL & "[" & Qry & "]", vbExclamation
         End If
         
         ExecSQLDao = -1
      End If
   Else
      Call AddDebug("ExecSql: OK [" & Qry & "]")
   End If

   'Debug.Print "ExecSQL: tiempo " & Format(Now - Tm, "nn:ss") & " [m:s]"

End Function

Function vFldDao(Fld As dao.Field, Optional ByVal bDeSql As Boolean = True) As Variant
   Dim bString As Boolean, bBoolean As Boolean
   
   bString = (Fld.Type = dbText Or Fld.Type = dbMemo Or Fld.Type = dbChar)
   bBoolean = (Fld.Type = dbBoolean)

   If IsNull(Fld) Then
      
      If bString Then
         vFldDao = ""
      Else
         vFldDao = 0
      End If
   
   ElseIf bString Then
   
      If bDeSql Then
         vFldDao = DeSQL(Fld.Value)
      Else
         vFldDao = Fld.Value
      End If
      
   ElseIf bBoolean Then
   
      vFldDao = Abs(Fld.Value)
      
   Else
      vFldDao = Fld.Value
   End If
      
End Function
