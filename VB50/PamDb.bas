Attribute VB_Name = "PamDb"
'************ Funciones para Trabajo con Bases de Datos ********
Option Explicit

' DATACON
' *** OJO *** : estas constantes solo funcionan en este archivo por lo tanto
' en el resto del programa se debe usar #If DATACON = 1 o = 2 Then
#Const DAO_CONN = 1     ' Database - Access  *** OJO: estas ctes sólo funcionan en este archivo
#Const ADO_CONN = 2     ' Connection

' Connect String para OpenDatabase
' "ODBC;Driver=SQL Server;Server=Pam;Database=GestSap;UID=usuario;PWD=clave;"

' Connect String para Connection
' "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=USUARIO;Initial Catalog=BASE_DE_DATOS;Data Source=EQUIPO"
' "Provider= Microsoft.Jet.OLEDB.4.0;Data Source=j:\WebPaola\Centro.mdb;"
Public Const API_KEY_TECNOBACKDESA = "RCCYCj9xZl5Wld7jNLEc4Jct4xdaB1v4pLKc1ybd"
Public Const API_KEY_TECNOBACK = "03vf7bY0lY9dIKQxwPYNm3al0MCo5RaX5b19HGaP"

Public SqlErr As Long
Public SqlError As String
Public SqlnRec As Long  ' cantidad de registros afectados por el último ExecSql

Public Const SQLDATEFMT = "'yyyymmdd'"

' Tipos de SQL - DbType
Public Const SQL_ACCESS  As Integer = 1
Public Const SQL_SERVER  As Integer = 2   ' Sql Server
Public Const SQL_MYSQL  As Integer = 3
Public gDbType As Integer

Public Const GL_LWILD = 1
Public Const GL_RWILD = 2
Public Const GL_WILD = (GL_LWILD Or GL_RWILD)
Public Const GL_AND = 4
Public Const GL_OR = 8

Private Type LstRs_t
   Rs    As Object ' Recordset
   Qry   As String
   Mrk   As String
End Type
Private Const NRS = 30
Dim LstRs(NRS) As LstRs_t
Private nRsOpen As Integer ' Para saber cuantos Recordsets hay abiertos

Private Const NTR = 30
Private LstTrans(NTR) As String
Private iTrans As Integer

Private FldType() As Integer
Private FldName() As String

' Para ser usada en ImportFile
Public Type Campo_t
   Campo    As String
   Tipo     As String
   Largo    As Integer
   Dec      As Byte
   Fmt      As String
   Def      As String
   ExCampo  As String   ' Le pone un 1 si usa el Default
End Type

'para ser usada en la AdvTbAddNewMult
Public Type AdvTbAddNew_t
   FldName     As String
   FldValue    As String
   FldIsNum    As Boolean
End Type

Public Type BCP_t
   Cmd      As String
   FInfo    As String
   FnLog    As String   ' Log general del bcp
   TblErr   As String   ' Log de la tabla
   Msg      As String
End Type


Public Sub CheckRs(Optional ByVal bMsg As Boolean = 1)
   Dim i As Integer

   If nRsOpen > 0 Then
      Debug.Print "*************************************"
      Debug.Print "*** Quedaron " & nRsOpen & " recordsets abiertos ****"
      
      For i = 0 To NRS
         If Not LstRs(i).Rs Is Nothing Then
            Debug.Print "*** Mrk=[" & LstRs(i).Mrk & "] Rs[" & LstRs(i).Qry & "]"
         End If
         
      Next i
      Debug.Print "**************************"
      
      If bMsg And W.InDesign Then
         MsgBox1 "Quedaron " & nRsOpen & " recordsets abiertos", vbExclamation
      End If
      
   End If
   
   If iTrans Then
      Debug.Print "***** Quedaron " & iTrans & " Transacciones abiertas *****"
      
      If bMsg And W.InDesign Then
         MsgBox1 "Quedaron " & iTrans & " Transacciones abiertas", vbExclamation
      End If
   End If
   
End Sub
Public Sub CloseDb(Db As Object)
   
   If Not Db Is Nothing Then
      
      Db.Close
      
      Set Db = Nothing
      
   End If

End Sub

Public Sub CloseRs(Rs As Object) ' Recordset
   
   If Not Rs Is Nothing Then
      
      Rs.Close
      
      Call RemoveRs(Rs)  ' *** PARA DEBUG

      Set Rs = Nothing
            
   End If

End Sub
#If DATACON = DAO_CONN Then
Public Function ExecSQL(Db As Database, ByVal Qry As String, Optional ByVal bErrMsg As Boolean = True, Optional ByVal nTry As Byte = 0) As Long
#Else
Public Function ExecSQL(Db As Connection, ByVal Qry As String, Optional ByVal bErrMsg As Boolean = True, Optional ByVal nTry As Byte = 0) As Long
#End If
   Dim Rc As Long, nRec As Long, DbType As Integer
   Dim LogErr As String, Errno As Long
   Dim Tm As Double
   Dim ConnStr As String
   Dim iAux As Integer
   
   ExecSQL = -1

   If Trim(Qry) = "" Or Db Is Nothing Then
      Exit Function
   End If

   If gDbType <= 0 Then
      Debug.Print "*** Falta asignar gDbType = SqlType(Db)"
      DbType = SqlType(Db)
   Else
      DbType = gDbType
   End If
   
   If DbType = SQL_MYSQL Then
      Qry = ReplaceStr(Qry, Chr(164), "\'") ' 15 dic 2017: para que sea más estándar
   ElseIf DbType = SQL_ACCESS Or DbType = SQL_SERVER Then
      Qry = ReplaceStr(Qry, "''", "NULL")
      Qry = ReplaceStr(Qry, Chr(164), "''") ' 15 dic 2017: para que sea más estándar
   End If

   Tm = Now

   On Error Resume Next
   ERR.Clear
   SqlnRec = -1

#If DATACON = DAO_CONN Then
   dao.Errors.Refresh
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If
   
   ' Probamos las dos alternativas, dependiendo del tipo de conexiòn
   If InStr(1, ConnStr, "odbc", vbTextCompare) = 0 Then
      nRec = -1
#If DATACON = DAO_CONN Then
      Db.Execute Qry ' Access
      nRec = Db.RecordsAffected
#Else
      Db.Execute Qry, nRec
#End If
'#If DATACON = DAO_CONN Then
'      If Err Then
'         Err.Clear
'
'         Rc = -1
'         Db.Execute Qry, dbSQLPassThrough
'         Rc = Db.RecordsAffected
'      End If
'#End If

   Else ' ODBC ?
   
#If DATACON = DAO_CONN Then
      
      nRec = -1
      Db.Execute Qry, dbSQLPassThrough
      nRec = Db.RecordsAffected
      
      If ERR Then
         LogErr = ", Error-1: " & ERR & ", " & GetDbErr(Db, Errno)
         Rc = -1
         
         ' Si no es SQL Server, volvemos a intentar
         If InStr(1, Db.Connect, "SQL Server", vbTextCompare) = 0 Then
            ERR.Clear
#End If
            nRec = -1
            Db.Execute Qry, nRec ' Access
            ' nRec = Db.RecordsAffected *** 4 ene 2008
#If DATACON = DAO_CONN Then
         End If
      End If
#End If

   End If
   
'   For iAux = 0 To 104
'      Debug.Print iAux, Db.Properties(iAux).Name, Db.Properties(iAux)
'   Next iAux
   
   SqlnRec = nRec
   
   If nRec = -1 And ERR = 0 Then
      nRec = 1  ' para que no sea 0
   End If
   
   If Errno <> 0 Then
      SqlErr = Errno
   Else
      SqlErr = ERR
   End If
   
   SqlError = Error & LogErr
   ExecSQL = nRec
   
   If SqlErr Then
      If UCase(Left(Qry, 5)) <> "DROP " Then
         SqlError = SqlError & GetDbErr(Db, Errno)
         
         If SqlError = "" Then
            SqlError = "Error"
         End If
         
         Call AddLog("ExecSql: bMsg=" & bErrMsg & ", " & LogErr & "; Error " & SqlErr & ", " & SqlError & "; [" & Qry & "] [Dec=" & Format(1234.56, DBLFMT2) & "-" & True & "]")
         
'         If SqlType(Db, ConnStr) = SQL_SERVER And InStr(1, SqlError, "LOCK resource", vbTextCompare) Then
         If DbType = SQL_SERVER And InStr(1, SqlError, "LOCK resource", vbTextCompare) Then
         
            If nTry <= Val(GetIniString(gCfgFile, "Config", "nLockTry", "5")) Then
               Call Sleep(1000 * Val(GetIniString(gCfgFile, "Config", "nLockSec", "15")))
               ExecSQL = ExecSQL(Db, Qry, bErrMsg, nTry + 1)
               Exit Function
            ElseIf MsgBox1("SQL Server no dispone de más recursos LOCK." & vbCrLf & "¿Desea re-intentar la consulta?", vbYesNo Or vbQuestion) = vbYes Then
               ExecSQL = ExecSQL(Db, Qry, bErrMsg)
               Exit Function
            End If
         End If
         
         If bErrMsg Then
            MsgBox1 LogErr & vbCrLf & SqlError & vbLf & "[" & Qry & "]", vbExclamation
            'MsgBox1 "Error " & SqlErr & ", " & SqlError & NL & "[" & Qry & "]", vbExclamation
         End If
         
         ExecSQL = -1
      End If
   Else
      Call AddDebug("ExecSql: OK [" & Qry & "]")
   End If

   'Debug.Print "ExecSQL: tiempo " & Format(Now - Tm, "nn:ss") & " [m:s]"

End Function

' Se supone que sirve para Combobox y Listbox
' idSel = -1 ==> Se selecciona el primero
' Qry del tipo  SELECT Texto, Codigo FROM ....
'
#If DATACON = DAO_CONN Then
Public Function FillCombo(Cmb As Control, Db As Database, ByVal Qry As String, ByVal idSel As Long, Optional ByVal bFCase As Boolean = 0, Optional ByVal MaxElements As Integer = -1, Optional ByVal Fmt As String = "") As Long
#Else
Public Function FillCombo(Cmb As Control, Db As Connection, ByVal Qry As String, ByVal idSel As Long, Optional ByVal bFCase As Boolean = 0, Optional ByVal MaxElements As Integer = -1, Optional ByVal Fmt As String = "") As Long
#End If
   Dim Rs As Recordset
   Dim Txt As String
   Dim nf As Integer, iSel As Integer

   FillCombo = 0

   Set Rs = OpenRs(Db, Qry)
   If Rs Is Nothing Then
      FillCombo = -2
      Exit Function
   End If
   
   nf = Rs.Fields.Count
   iSel = -1
   
   Do Until Rs.EOF
   
      If MaxElements > 0 Then
         If Cmb.ListCount > MaxElements Then
            FillCombo = -3
            Exit Do
         End If
      End If
      
      Txt = vFld(Rs(0))
      If Fmt <> "" Then
         Txt = Format(Txt, Fmt)
      ElseIf bFCase Then
         Txt = FCase(Txt)
      End If
      
      Cmb.AddItem Txt
      
      If nf > 1 Then
      
         Cmb.ItemData(Cmb.NewIndex) = vFld(Rs(1))
   
         If idSel >= 0 And idSel = vFld(Rs(1)) Then
'            Cmb.ListIndex = Cmb.NewIndex    2 nov 2010 pam
            iSel = Cmb.NewIndex     ' Mejor lo hacemos al final para evitar Rs anidados
         End If
      End If

      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)

   If idSel >= 0 And iSel >= 0 Then
      Cmb.ListIndex = iSel
   ElseIf idSel = -1 And Cmb.ListIndex < 0 And Cmb.ListCount > 0 Then
      Cmb.ListIndex = 0
   End If

End Function

#If DATACON = DAO_CONN Then
Public Sub FillComboDate(Cmb As Control, Db As Database, Qry As String, idSel As Long)
#Else
Public Sub FillComboDate(Cmb As Control, Db As Connection, Qry As String, idSel As Long)
#End If
   Dim Rs As Recordset, i As Integer, nf As Integer

   Set Rs = OpenRs(Db, Qry)
   If Rs Is Nothing Then
      Exit Sub
   End If
   
   nf = Rs.Fields.Count
   
   Do Until Rs.EOF
      If nf > 2 Then ' se asume que el tercer campo es la cantidad de registros
         Cmb.AddItem FmtFecha(vFld(Rs(0))) & "  (" & vFld(Rs(2)) & ")"
      Else
         Cmb.AddItem FmtFecha(vFld(Rs(0)))
      End If
      
      i = Cmb.NewIndex
      Cmb.ItemData(i) = Val(vFld(Rs(1)))
      
      If idSel >= 0 And idSel = vFld(Rs(1)) Then
         Cmb.ListIndex = 0
      End If

      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)

   If idSel = -1 And Cmb.ListCount > 0 Then
      Cmb.ListIndex = 0
   End If

End Sub

'Public Sub CambiarClaveOLD()
'Dim ano As Long
'Dim sArchivo As String
'Dim i As Integer
'Dim DbName As String
'
'DbName = gDbPath & "\" & BD_COMUN
'
'Call SetDbSecurityCambio(DbName, PASSW_LEXCONT_NEW, gCfgFile, SG_SEGCFG, gComunConnStr, "", "")
'Call SetDbSecurityCambio(DbName, PASSW_LEXCONT_NEW2, gCfgFile, SG_SEGCFG, gComunConnStr, "", "")
'Call SetDbSecurityCambio(DbName, PASSW_LEXCONT, gCfgFile, SG_SEGCFG, gComunConnStr, "", "")
'Call OpenDbAdm
'
'ano = 2014
'For i = 0 To 10
'    ano = ano + 1
'    If ExistFile(gDbPath & "\Empresas\" & ano) Then
'
'    sArchivo = Dir(gDbPath & "\Empresas\" & ano & "\")
'    Do While sArchivo <> ""
'    'List1.AddItem sArchivo
'    DbName = gDbPath & "\Empresas\" & ano & "\" & sArchivo
'
'    Call SetDbSecurityCambio(DbName, PASSW_PREFIX_NEW & Replace(sArchivo, ".mdb", ""), gCfgFile, SG_SEGCFG, gEmpresa.ConnStr, "EMP", Replace(sArchivo, ".mdb", ""))
'    Call SetDbSecurityCambio(DbName, PASSW_PREFIX & Replace(sArchivo, ".mdb", ""), gCfgFile, SG_SEGCFG, gEmpresa.ConnStr, "EMP", Replace(sArchivo, ".mdb", ""))
'    Call OpenDbEmp(Replace(sArchivo, ".mdb", ""), ano)
'    'CloseDb (DbMain)
'    sArchivo = Dir
'    Loop
'
'    End If
'
'Next i
'
'
'
'End Sub
' Abre un recordset para consulta
'
#If DATACON = DAO_CONN Then
Public Function OpenRs(Db As Database, ByVal Qry As String, Optional ByVal bErrMsg As Boolean = True, Optional ByVal RsType As Integer = dbOpenSnapshot, Optional ByVal Mrk As String = "", Optional ByVal RsOption As Integer = 0, Optional bRepApostr As Boolean = 1) As Recordset
#Else
Public Function OpenRs(Db As Connection, ByVal Qry As String, Optional ByVal bErrMsg As Boolean = True, Optional ByVal RsType As Integer = 0, Optional ByVal Mrk As String = "", Optional ByVal RsOption As Integer = 0, Optional bRepApostr As Boolean = 1) As Recordset
#End If
   Dim Rs As Recordset, ConnStr As String, Errno As Long

   Set OpenRs = Nothing

   If Trim(Qry) = "" Then
      Exit Function
   End If
    
   If gDbType <= 0 Then
      Debug.Print "*** Falta asignar gDbType = SqlType(Db)"
   End If
    
   If bRepApostr Then
      Qry = ReplaceStr(Qry, "''", "NULL")
   End If
   
   On Error Resume Next
   Set OpenRs = Nothing

   'Set Rs = Db.OpenRecordset(Qry, dbOpenForwardOnly, dbConsistent, dbReadOnly) ' , dbOptimistic )
   
   ERR.Clear
   SqlErr = 0
   SqlError = ""
#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
   dao.Errors.Refresh
   
   If Left(ConnStr, 5) <> "ODBC;" Or Left(ConnStr, 4) = "Jet " Then
      Set Rs = Db.OpenRecordset(Qry, RsType, dbConsistent Or RsOption, dbReadOnly)  ' , dbOptimistic )
   End If
   
   If Rs Is Nothing And (Left(ConnStr, 5) = "ODBC;" Or InStr(1, ConnStr, "Sql Server", vbTextCompare) > 0) Then
      dao.Errors.Refresh
      ERR.Clear
      'Set Rs = Db.OpenRecordset(Qry, dbOpenForwardOnly, dbConsistent Or dbSQLPassThrough, dbReadOnly) ' , dbOptimistic )
      Set Rs = Db.OpenRecordset(Qry, RsType, dbConsistent Or dbSQLPassThrough, dbReadOnly)  ' , dbOptimistic )
      
'      While Rs.StillExecuting
'         DoEvents
'      Wend

   End If
   
   SqlErr = ERR
   SqlError = Error
   
#Else
   Db.CursorLocation = adUseClient
   
   Db.CommandTimeout = 0
    ''3426560
   'Db.ConnectionTimeout = 0
   '3426560
   Set Rs = Db.Execute(Qry)
   If Db.Errors.Count > 0 Then
      DoEvents
      SqlErr = Db.Errors(0).Number
      SqlError = Db.Errors(0).Description
   End If
#End If
   
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
   
   Call AddRs(Rs, Qry, Mrk) ' *** PARA DEBUG
   
   Set OpenRs = Rs
   
End Function

'
'  El parametro bool ya no se usa. True  -> ParaSQL()
'                                  False -> DeSQL()
'
Function ParaSQL(ByVal Buf As String, Optional ByVal bAllowEmpty As Boolean = 0, Optional ByVal DbType As Integer = -1) As String
   Dim i As Integer

   Buf = Trim(Buf)

   If Buf = "" Then
   
      If bAllowEmpty = False Then
         ParaSQL = " "
      End If
      
      Exit Function
   End If
   
   'Buf = ReplaceStr(Buf, "'", Chr(253))        ' reemplaza ' por 253
   'Buf = ReplaceStr(Buf, Chr(34), Chr(254))   ' reemplaza " por 252
   'Buf = ReplaceStr(Buf, "|", Chr(255))        ' reemplaza | por 251
   'Buf = ReplaceStr(Buf, Chr(10), Chr(248))   ' reemplaza nl por 250
   'Buf = ReplaceStr(Buf, Chr(13), Chr(247))   ' reemplaza cr por 249

   For i = 0 To UBound(ChrMask)
      Buf = ReplaceStr(Buf, ChrMask(i, 0), ChrMask(i, 1))
   Next i

   If DbType = -1 Then
      DbType = gDbType
   End If

   If DbType = SQL_MYSQL Then
      Buf = ReplaceStr(Buf, "\", "\\") ' 4 oct 2013: por el MySQL
      Buf = ReplaceStr(Buf, Chr(164), "\'") ' 9 ene 2018: por el MySQL
   End If

   ParaSQL = Buf

End Function
Function UParaSQL(ByVal Buf As String) As String

'   UParaSQL = UCase(ParaSQL(Buf, Bool))
   UParaSQL = UCase(ParaSQL(Buf))
End Function

Function SqlvFld(Fld As Field, Optional ByVal bDeSql As Boolean = True) As Variant

   If IsNull(Fld) Then
      SqlvFld = " NULL "
   Else
      SqlvFld = vFld(Fld, bDeSql)
   End If
   
End Function

'
' Permite consultar por un campo sin preocuparse si es NULL
'
Function vFld(Fld As Field, Optional ByVal bDeSql As Boolean = True) As Variant
   Dim bString As Boolean, bBoolean As Boolean
      
   #If DATACON = DAO_CONN Then
      bString = (Fld.Type = dbText Or Fld.Type = dbMemo Or Fld.Type = dbChar)
      bBoolean = (Fld.Type = dbBoolean)
   #Else
      bString = (Fld.Type = adChar Or Fld.Type = adVarChar Or Fld.Type = adLongVarChar Or Fld.Type = adLongVarWChar Or Fld.Type = adVarWChar Or Fld.Type = adWChar)
      bBoolean = (Fld.Type = adBoolean)
   #End If

   If IsNull(Fld) Then
      
      If bString Then
         vFld = ""
      Else
         vFld = 0
      End If
   
   ElseIf bString Then
   
      If bDeSql Then
         vFld = DeSQL(Fld.Value)
      Else
         vFld = Fld.Value
      End If
      
   ElseIf bBoolean Then
   
      vFld = Abs(Fld.Value)
      
   Else
      vFld = Fld.Value
   End If
      
End Function

' Para CheckBox
Function ValSiNo2(Fld As Field, Optional ByVal valNoChk As Long = 0) As Integer

   If IsNull(Fld) Then
      ValSiNo2 = 2
   Else
      Dim v As Integer
      v = Fld
      If Fld = valNoChk Then
         ValSiNo2 = 0
      Else
         ValSiNo2 = 1
      End If
   End If
   
End Function
' Sirve para CheckBox y RadioButton
Function ChkSiNo2(Ctrl As CheckBox, Optional ByVal valChk As String = "1", Optional ByVal valNoChk As String = "0") As String
   
   Select Case Ctrl.Value
      Case vbChecked:
         ChkSiNo2 = valChk
      Case vbUnchecked:
         ChkSiNo2 = valNoChk
      Case Else:
         ChkSiNo2 = " NULL "
      
   End Select
      
End Function

'
' Funciona con SplitBuf
'
Function JoinBuf(Rs As Recordset, Fld As String, n As Integer) As String
   Dim Buf As String, Txt As String
   Dim i As Integer, Ini As Integer, Fin As Integer
   Dim b As Integer, l As Integer

   n = Abs(n)

   Buf = ""
   For i = 1 To n
      Txt = vFld(Rs(Fld & i), True)

      If InStr(Txt, "@") Then

         b = 1
         l = Len(Txt)

         If Left(Txt, 1) = "@" Then
            b = 2
            l = l - 1
         End If

         If Right(Txt, 1) = "@" And l > 0 Then
            l = l - 1
         End If

         'Buf = Buf & Mid(Txt, Ini, Len(Txt) - Fin)
         Buf = Buf & Mid(Txt, b, l)
      Else
         Buf = Buf & Txt
      End If
   Next i
    
   Buf = ReplaceStr(Buf, ChrMask(0, 1), ChrMask(0, 0))

   JoinBuf = DeSQL(Buf)

End Function
'
' Separa el buffer en varios buffers más cortos para INSERT o UPDATE en base de datos
'
Function SplitBuf(ByVal Fld As String, ByVal Buf As String, ByVal n As Integer) As String
   Dim i As Integer
   Dim Q1 As String, Sep As String
   Dim l As Integer, Ln As Integer
   Dim Aux As String

   Ln = 255
   Sep = ""

   Buf = ParaSQL(Buf)
   l = Len(Buf)

   Q1 = ""
   If Fld = "" Then ' es INSERT
      For i = 1 To n
         Aux = Mid(Buf, (i - 1) * Ln + 1, Ln)
         If Aux = "" Then
            Aux = " "
         End If
         
         Q1 = Q1 & ", '" & Sep & Aux & Sep & "' "
      Next i

   Else              'es UPDATE
      For i = 1 To n
         Aux = Mid(Buf, (i - 1) * Ln + 1, Ln)
         If Aux = "" Then
            Aux = " "
         End If
         
         Q1 = Q1 & ", " & Fld & i & "= '" & Sep & Aux & Sep & "' "
      Next i

   End If

   SplitBuf = Mid(Q1, 2)

End Function

Public Sub AddRs(Rs As Object, Qry As String, Mrk As String)
   Dim i As Integer
   Static n As Integer
   
   If Rs Is Nothing Then
      Exit Sub
   End If
   
   If n = 0 Then
      n = 3
   End If
   
   For i = 0 To NRS
      If LstRs(i).Rs Is Nothing Then
         Set LstRs(i).Rs = Rs
         LstRs(i).Qry = Qry
         LstRs(i).Mrk = Mrk
         Exit For
      End If
   Next i
         
   nRsOpen = nRsOpen + 1

   If nRsOpen >= n Then
      Debug.Print "* Rs abiertos: " & nRsOpen
      n = n + 1
   End If
               
End Sub
Private Sub RemoveRs(Rs As Object) ' Recordset
   Dim i As Integer
   
   If Rs Is Nothing Then
      Exit Sub
   End If
   
   For i = 0 To NRS
      If LstRs(i).Rs Is Rs Then
         Set LstRs(i).Rs = Nothing
         LstRs(i).Qry = ""
         LstRs(i).Mrk = ""
         Exit For
      End If
   Next i
      
   nRsOpen = nRsOpen - 1
         
End Sub
#If DATACON = DAO_CONN Then
Public Function TabFields(Db As Database, ByVal TabName As String, Optional ByVal bTabName As Boolean = 1, Optional ByVal bComma As Boolean = 1) As String
#Else
Public Function TabFields(Db As Connection, ByVal TabName As String, Optional ByVal bTabName As Boolean = 1, Optional ByVal bComma As Boolean = 1) As String
#End If
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer

   Q1 = "SELECT * FROM " & TabName
   Set Rs = OpenRs(Db, Q1)

   If bTabName Then
      TabName = TabName & "."
   Else
      TabName = ""
   End If

   Q1 = ""
   ' se le ponen , para poder hacer los ReplaceStr en forma facil  ",campo1,campo2,campo3,"
   For i = 0 To Rs.Fields.Count - 1
      Q1 = Q1 & "," & TabName & Rs.Fields(i).Name
   Next i
   
   If bComma Then
      Q1 = Q1 & ","
   Else
      Q1 = Mid(Q1, 2)
   End If
   
   Call CloseRs(Rs)

   TabFields = Q1
   
End Function

'
' Esta funcion prepara una fecha para actualizar, insertar o comparar una fecha
'
' XFmtDate: Sirve para preparar una fecha para una consulta
'           SQL y formatearla según el servidor.
'
' Ej: SELECT * FROM Tb WHERE FInicio > " & SqlDate(Fecha)
'     UPDATE Tabla Set Fecha = " & SqlDate(Fecha)
'
#If DATACON = DAO_CONN Then
Function SqlDate(Db As Database, ByVal Fecha As Double) As String
#Else
Function SqlDate(Db As Connection, ByVal Fecha As Double) As String
#End If
   Dim Fmt As String, ConnStr As String

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If

   If Left(ConnStr, 5) <> "ODBC;" And InStr(1, ConnStr, "Sql Server", vbTextCompare) = 0 Then  ' Ms Access ??
      'Fmt = "#" & Format(Fecha, "mm/dd/yyyy") & "#"
      Fmt = str(Fecha)
   Else
      ' ODBC
      Fmt = "{ts " & Format(Fecha, "'yyyy-mm-dd hh:nn:ss'") & "}"
      'Fmt = Format(Fecha, SQLDATEFMT)
   End If

   SqlDate = Fmt

End Function
'
' Esta funcion prepara una hora para actualizar, insertar o comparar una fecha
'
' XFmtDate: Sirve para preparar una fecha para una consulta
'           SQL y formatearla según el servidor.
'
' Ej: SELECT * FROM Tb WHERE FInicio > " & SqlDate(Fecha)
'     UPDATE Tabla Set Fecha = " & SqlDate(Fecha)
'
#If DATACON = DAO_CONN Then
Function SqlTime(Db As Database, ByVal Hora As Double) As String
#Else
Function SqlTime(Db As Connection, ByVal Hora As Double) As String
#End If
   Dim Fmt As String, ConnStr As String

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If

   If Left(ConnStr, 5) <> "ODBC;" And InStr(1, ConnStr, "Sql Server", vbTextCompare) = 0 Then  ' Ms Access ??
      'Fmt = "#" & Format(Fecha, "mm/dd/yyyy") & "#"
      Fmt = str(Hora)
   Else
      ' ODBC
      Fmt = "{ts " & Format(Hora, "'hh:nn:ss'") & "}"
      'Fmt = Format(Fecha, SQLDATEFMT)
   End If

   SqlTime = Fmt

End Function

#If DATACON = DAO_CONN Then
Function SqlDateF(Db As Database, ByVal FldName As String) As String
#Else
Function SqlDateF(Db As Connection, ByVal FldName As String) As String
#End If
   Dim Fmt As String, ConnStr As String

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If

   If Left(ConnStr, 5) <> "ODBC;" And InStr(1, ConnStr, "Sql Server", vbTextCompare) = 0 Then  ' Ms Access ??
      'Fmt = "#" & Format(Fecha, "mm/dd/yyyy") & "#"
      SqlDateF = "int(" & FldName & ")"
   Else
      ' ODBC
      SqlDateF = "cast(floor(cast(" & FldName & " as float)) as datetime )"
      'Fmt = Format(Fecha, SQLDATEFMT)
   End If

End Function

'
' Esta funcion prepara una fecha para actualizar, insertar o comparar una fecha
'
' XFmtDate: Sirve para preparar una fecha para una consulta
'           SQL y formatearla según el servidor.
'
' Ej: SELECT * FROM Tb WHERE FInicio > " & SqlDate(Fecha)
'     UPDATE Tabla Set Fecha = " & SqlDate(Fecha)
'
#If DATACON = DAO_CONN Then
Function SqlMonth(Db As Database, ByVal Fld As String) As String
#Else
Function SqlMonth(Db As Connection, ByVal Fld As String) As String
#End If
   Dim Fmt As String
   Dim ConnStr As String

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If

   If Left(ConnStr, 5) <> "ODBC;" Then
      'Fmt = "#" & Format(Fecha, "mm/dd/yyyy") & "#"
      Fmt = "( Year(" & Fld & ") * 100 + Month(" & Fld & ") )"
   Else
      ' ODBC
      'Fmt = "( DatePart( yyyy, " & Fld & ") * 100 + DatePart( mm, " & Fld & ") )"
      'Fmt = "( {fn Year(" & Fld & ")} * 100 + {fn Month(" & Fld & ")} )"
      Fmt = "( {fn Year(" & Fld & "-2)} * 100 + {fn Month(" & Fld & "-2)} )"  ' el -2 es por el SQL Server
   End If

   SqlMonth = Fmt

End Function
'Sirve para usar una Fecha en una consulta SQL en el caso que se guarde en un Long
Function SqlMonthLng(ByVal LngFechaFld As String) As String
   
   If gDbType = SQL_SERVER Then
      SqlMonthLng = "Month(" & LngFechaFld & " - 2)"
   Else
      SqlMonthLng = "Month(" & LngFechaFld & ")"
   End If
   
End Function
'Sirve para usar una Fecha en una consulta SQL en el caso que se guarde en un Long
Function SqlYearLng(ByVal LngFechaFld As String) As String
   
   If gDbType = SQL_SERVER Then
      SqlYearLng = "Year(" & LngFechaFld & " - 2)"
   Else
      SqlYearLng = "Year(" & LngFechaFld & ")"
   End If
   
End Function
'
' Sirve para consultar/asignar la fecha/hora actual
'
' Ej: SELECT * FROM Tb WHERE FInicio > " & SqlNow()
'     UPDATE Tabla Set Fecha = " & SqlNow()
' OJO: Con SQL Server pueder dar problemas +/- 2 días
'
#If DATACON = DAO_CONN Then
Function SqlNow(Db As Database) As String
#Else
Function SqlNow(Db As Connection) As String
#End If
   Dim Q1 As String, ConnStr As String

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If

   If SqlType(Db) = SQL_SERVER Then

'   If InStr(1, ConnStr, "Sql Server", vbTextCompare) Then
      ' ODBC
      Q1 = "GetDate()"
   ElseIf Left(ConnStr, 5) <> "ODBC;" Then
      Q1 = "NOW"
   ElseIf gDbType = SQL_MYSQL Then
      Q1 = "Now()"
   Else
      ' ODBC
      Q1 = "{fn NOW()}"
   End If

   SqlNow = Q1

End Function
'
' Sirve para consultar/asignar la fecha (solo fecha) actual
'
' Ej: SELECT * FROM Tb WHERE FInicio > " & SqlToday()
'     UPDATE Tabla Set Fecha = " & SqlToday()
'
#If DATACON = DAO_CONN Then
Function SqlToday(Db As Database) As String
#Else
Function SqlToday(Db As Connection) As String
#End If
   Dim Q1 As String, ConnStr As String

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If

   If Left(ConnStr, 5) <> "ODBC;" Then
      Q1 = "Int(NOW)"
   Else
      ' ODBC
      Q1 = "{fn CurDate()}"  ' sólo la fecha
   End If

   SqlToday = Q1

End Function
'
' Sirve para consultar/asignar la fecha actual, cuando las fechas se guardan como Long
'
' Ej: SELECT * FROM Tb WHERE FInicio > " & SqlNow()
'     UPDATE Tabla Set Fecha = " & SqlNow()
'
#If DATACON = DAO_CONN Then
Function SqlNowI(Db As Database) As String
#Else
Function SqlNowI(Db As Connection) As String
#End If
   Dim Q1 As String, ConnStr As String

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If

   If Left(ConnStr, 5) <> "ODBC;" And InStr(1, ConnStr, "Sql Server", vbTextCompare) = 0 Then ' Access ?
      Q1 = "Int(NOW)"
   ElseIf gDbType = SQL_MYSQL Then
      Q1 = "(TO_DAYS(CurDate())-693959)" ' retorna la cantidad de días desde 1 ene 1900
   Else
      ' ODBC SQL Server
      'Q1 = "Cast( {fn Now()} AS Integer)"
      ' floor(convert( float, GetDate()))
      Q1 = "(DateDiff( dy, '19000101', GetDate() ) + 2)"  ' el + 2 es por el SQL Server
   End If

   SqlNowI = Q1

End Function

'
' Forma la sentencia SQL dependiendo del motor
'
#If DATACON = DAO_CONN Then 'ADO
Public Function UpdateSQL(Db As Database, ByVal TbName As String, ByVal sSet As String, ByVal sFrom As String, ByVal sWhere As String, Optional ByVal bErrMsg As Boolean = True) As Long
#Else
Public Function UpdateSQL(Db As Connection, ByVal TbName As String, ByVal sSet As String, ByVal sFrom As String, ByVal sWhere As String, Optional ByVal bErrMsg As Boolean = True) As Long
#End If
   Dim Q1 As String, ConnStr As String, T As Integer

   T = SqlType(Db, ConnStr)

   If T = SQL_SERVER Or T = SQL_MYSQL Then
      Q1 = "UPDATE " & TbName & " SET " & sSet & " FROM " & sFrom & " " & sWhere
   Else
      Q1 = "UPDATE " & sFrom & " SET " & sSet & " " & sWhere
   End If
   
   UpdateSQL = ExecSQL(Db, Q1, bErrMsg)

End Function
'
' Forma la sentencia SQL dependiendo del motor
'
' sWhere = " WHERE ....
' sWhere = " INNER JOIN Tabla2 ON ... WHERE ....
'
#If DATACON = DAO_CONN Then
Public Function DeleteSQL(Db As Database, ByVal TbName As String, ByVal sWhere As String, Optional ByVal bErrMsg As Boolean = True) As Long
#Else
Public Function DeleteSQL(Db As Connection, ByVal TbName As String, ByVal sWhere As String, Optional ByVal bErrMsg As Boolean = True) As Long
#End If
   Dim Q1 As String, ConnStr As String, T As Integer, bJoin As Boolean
   
   bJoin = (InStr(1, Left(sWhere, 13), " JOIN ", vbTextCompare) > 0)
   
   T = SqlType(Db, ConnStr)
   If T = SQL_SERVER Or T = SQL_MYSQL Then
      
      If bJoin Then
         Q1 = "DELETE " & TbName
      Else
         Q1 = "DELETE"
      End If
   
   Else ' Access
      If bJoin Then
         Q1 = "DELETE " & TbName & ".*"
      Else
         Q1 = "DELETE *"
      End If
      
   End If
   
   If Len(sWhere) < 5 Then
      Debug.Print "DeleteSQL: Atención: no where el Where **** "
   End If
   
   Q1 = Q1 & " FROM " & TbName & " " & sWhere
   
   DeleteSQL = ExecSQL(Db, Q1, bErrMsg)

End Function
' DELETE con Join
#If DATACON = DAO_CONN Then
Public Function DeleteJSQL(Db As Database, ByVal TbName As String, ByVal sFrom As String, ByVal sWhere As String, Optional ByVal bErrMsg As Boolean = True) As Long
#Else
Public Function DeleteJSQL(Db As Connection, ByVal TbName As String, ByVal sFrom As String, ByVal sWhere As String, Optional ByVal bErrMsg As Boolean = True) As Long
#End If
   Dim Q1 As String, ConnStr As String, T As Integer

   T = SqlType(Db, ConnStr)

   If Len(sWhere) < 5 Then
      Debug.Print "DeleteJSQL: Atención: no where el Where **** "
   End If

   If T = SQL_SERVER Or T = SQL_MYSQL Then
      Q1 = "DELETE FROM " & TbName
   Else
      Q1 = "DELETE " & TbName & ".*"
   End If
   
   Q1 = Q1 & " FROM " & sFrom & " " & sWhere
   
   DeleteJSQL = ExecSQL(Db, Q1, bErrMsg)

End Function
' Genera un Like a partir de las palabras en Buf
' genera algo del tipo  (campo LIKE 'valor1' OR campo LIKE 'valor2')
#If DATACON = DAO_CONN Then 'ADO
Function GenLike(Db As Database, ByVal Buf As String, ByVal Fld As String, Optional Opt As Byte = GL_WILD) As String
#Else
Function GenLike(Db As Connection, ByVal Buf As String, ByVal Fld As String, Optional Opt As Byte = GL_WILD) As String
#End If
   Dim QN As String, Logic As String
   Dim i As Integer, j As Integer
   Dim W1 As String, W2 As String
   Dim ConnStr As String

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If
   
   If Left(ConnStr, 5) = "ODBC;" Or InStr(1, ConnStr, "Sql Server", vbTextCompare) > 0 Or InStr(1, ConnStr, "msdasql", vbTextCompare) > 0 Or InStr(1, ConnStr, "SQLNCLI", vbTextCompare) > 0 Or InStr(1, ConnStr, "ORA", vbTextCompare) > 0 Then
      QN = "%"    '"%"
   Else
      QN = "*"
   End If
   
   If (Opt And GL_LWILD) Then
      W1 = QN
   End If
   
   If (Opt And GL_OR) Then
      Logic = "OR "
   Else
      Logic = "AND "
   End If
   
   If (Opt And GL_RWILD) Then
      W2 = QN
   End If

   QN = ""
   Buf = Trim(Buf)
   j = 1
   Do
      i = InStr(j, Buf, " ")
      If i > 0 Then
         QN = QN & Fld & " LIKE '" & W1 & Trim(Mid(Buf, j, i - j)) & W2 & "' " & Logic
      Else
         QN = QN & Fld & " LIKE '" & W1 & Trim(Mid(Buf, j)) & W2 & "'"
         Exit Do
      End If

      j = i + 1
   Loop

   GenLike = "( " & QN & " )"

End Function
' Registra una conexión ODBC para SQL Server
' DSN:      Data source name
' DbName:   Nombre de la base de datos
' Svr:      Nombre del servidor

Sub RegODBC(ByVal DSN As String, ByVal DbName As String, ByVal Svr As String)
   Dim Dat As String
   Dim hkProtocol As Long
   Dim Rc As Long

   ' Registramos la conexion ODBC en el Registry para WinNT

   On Error Resume Next

   Rc = RegCreateKey(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBC.INI\" & DSN, hkProtocol)
   If Rc <> 0 Then
      MsgBox "No se pudo crear la conexión ODBC " & DSN & " (Rc=" & Rc & ")", vbExclamation
      Exit Sub
   End If
   
   Dat = DbName
   Call RegSetValueEx(hkProtocol, "Database", 0, REG_SZ, Dat, Len(Dat))
   Dat = DbName & " en SQL Server"
   Call RegSetValueEx(hkProtocol, "Description", 0, REG_SZ, Dat, Len(Dat))
   Dat = "sqlsrv32.dll"
   Call RegSetValueEx(hkProtocol, "Driver", 0, REG_SZ, Dat, Len(Dat))
   Dat = "No"
   Call RegSetValueEx(hkProtocol, "OEMTOANSI", 0, REG_SZ, Dat, Len(Dat))
   Dat = Svr
   Call RegSetValueEx(hkProtocol, "Server", 0, REG_SZ, Dat, Len(Dat))
   Dat = "No"
   Call RegSetValueEx(hkProtocol, "Trusted_Connection", 0, REG_SZ, Dat, Len(Dat))
   Dat = "No"
   Call RegSetValueEx(hkProtocol, "UseProcForPrepare", 0, REG_SZ, Dat, Len(Dat))

   Call RegCloseKey(hkProtocol)
   
   Rc = RegCreateKey(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBC.INI\ODBC Data Sources", hkProtocol)
   If Rc <> 0 Then
      Exit Sub
   End If
   
   Dat = "SQL Server"
   Call RegSetValueEx(hkProtocol, DSN, 0, REG_SZ, Dat, Len(Dat))
   
   Call RegCloseKey(hkProtocol)
   
   ' Registramos la conexion ODBC en el ODBC.INI para Win95/98
   
   Rc = WritePrivateProfileString("ODBC 32 bit Data Sources", DSN, "SQL Server (32 bit)", "ODBC.ini")
   Rc = WritePrivateProfileString(DSN, "Driver32", "sqlsrv32.dll", "ODBC.ini")
   
   Rc = FlushProfile(0, 0, 0, "ODBC.ini")
 
End Sub
#If DATACON = DAO_CONN Then 'ADO
Public Function CompactDb2(Db As Database, ByVal bMsg As Boolean, Optional ByVal ConnStr As String = "") As Long
   Dim DbPath As String

   If Not Db Is Nothing Then

      DbPath = Db.Name
      
      Call CloseDb(Db)
      
      CompactDb2 = CompactDb1(DbPath, bMsg, ConnStr)
   End If
      
End Function
#End If
#If DATACON = DAO_CONN Then 'ADO
Public Function CompactDb1(ByVal DbPath As String, ByVal bMsg As Boolean, Optional ByVal ConnStr As String = "") As Long
   Dim DbName As String
   Dim DbPathS As String, DbPathC As String, DbPathD As String
   Dim LenS As Long, LenD As Long
   Dim i As Integer
   
   i = rInStr(DbPath, "\")

'   For i = Len(DbPath) To 1 Step -1
'      If Mid(DbPath, i, 1) = "\" Then
'         DbName = Mid(DbPath, i + 1)
'         DbPath = Left(DbPath, i - 1)
'         Exit For
'      End If
'   Next i

   If i <= 0 Then
      CompactDb1 = 1
   End If
   
   DbName = Mid(DbPath, i + 1)
   DbPath = Left(DbPath, i - 1)
   
   i = DbEnUso(DbPath, DbName)
   If i Then
      CompactDb1 = i
      
      If i = 2 And bMsg Then
         MsgBox1 "La base está en uso, no puede compactarse.", vbExclamation
      End If
   
      Exit Function
   End If
   

   If Len(ConnStr) > 2 And Left(ConnStr, 1) <> ";" Then ' 2 feb 2016: si no trae ; se lo ponemos
      ConnStr = ";" & ConnStr
   End If

   ' preparamos los nombres de archivos
   DbPathS = DbPath & "\" & DbName     ' base fuente
   
   i = 1
   Do
      DbPathC = DbPath & "\C" & i & "_" & DbName   ' base backup
      If ExistFile(DbPathC) = False Then
         Exit Do
      End If
      i = i + 1
   Loop
   
   DbPathD = DbPath & "\~C" & DbName    ' base destino
      
   On Error Resume Next
   LenS = FileLen(DbPathS)
      
   CompactDatabase DbPathS, DbPathD, dbLangGeneral & ConnStr, , dbLangGeneral & ConnStr
   
   If ERR Then
   
      If bMsg Then
         Call MsgErr(NL & "Base en " & DbPathS)
      End If
      
   Else ' Todo OK
      Kill DbPathC
      If ERR = 0 Or ERR = 53 Then ' 53: File not found
         ERR.Clear
         LenD = FileLen(DbPathD)
         Name DbPathS As DbPathC ' Dejamos un backup de lo que había
         If ERR = 0 Then
            Name DbPathD As DbPathS
         End If
      End If
      
      If bMsg Then
         If ERR = 0 Then
            MsgBox1 "La base fue compactada con éxito, su tamaño se redujo en un " & Format((LenS - LenD) / LenS, "0.0%") & vbCrLf & vbCrLf & "La base original quedó en " & DbPathC, vbInformation
         Else
            Call MsgErr("La base no fue compactada.")
         End If
      End If
      
   End If

   CompactDb1 = ERR
   
End Function
#End If
#If DATACON = DAO_CONN Then 'ADO
Public Function CompactDb(ByVal DbPath As String, ByVal DbName As String, ByVal bMsg As Boolean, Optional ByVal bDelTmpFiles As Boolean = 0) As Long
   Dim DbPathS As String, DbPathB As String, DbPathC As String
   Dim LenS As Long, LenC As Long

   DbPathS = DbPath & "\" & DbName     ' Original
   DbPathB = DbPath & "\B_" & DbName   ' Backup
   DbPathC = DbPath & "\C_" & DbName   ' Compactada
      
   On Error Resume Next
   LenS = FileLen(DbPathS)
   CompactDatabase DbPathS, DbPathC
   
   If ERR Then
      If bMsg Then
         Call MsgErr(NL & "Base en " & DbPathS)
      End If
   Else ' Todo OK
      Kill DbPathB
      If ERR = 0 Or ERR = ERR_FILENOTFND Then ' 53: File not found
         LenC = FileLen(DbPathC)
         ERR.Clear
         Name DbPathS As DbPathB ' Dejamos un backup de lo que había
         If ERR = 0 Then
            Name DbPathC As DbPathS
         End If
      End If
      
      If bMsg Then
         If ERR = 0 Then
            MsgBox1 "La base fue compactada con éxito, su tamaño se redujo en un " & Format((LenS - LenC) / LenS, "0.0%") & NL & "La base original quedó en " & DbPathB, vbInformation
         Else
            Call MsgErr("La base no fue compactada.")
         End If
      End If
      
   End If

   CompactDb = ERR

   If ERR = 0 And bDelTmpFiles Then
      Kill DbPathB
   End If

End Function
#End If
#If DATACON = DAO_CONN Then 'ADO
Public Function DbEnUso(ByVal DbPath As String, ByVal DbName As String) As Integer
   Dim i As Integer

   i = InStr(DbName, ".")
   If i <= 0 Then
      DbEnUso = 1
      Exit Function
   End If

   If ExistFile(DbPath & "\" & Left(DbName, i) & "ldb") Then
      DbEnUso = 2
   End If

End Function


Public Function RepairDb1(ByVal DbPath As String, ByVal bMsg As Boolean) As Long
   Dim DbName As String
   Dim DbPathS As String, DbPathR As String
   Dim i As Integer
   
   i = rInStr(DbPath, "\")
   
'   For i = Len(DbPath) To 1 Step -1
'      If Mid(DbPath, i, 1) = "\" Then
'         DbName = Mid(DbPath, i + 1)
'         DbPath = Left(DbPath, i - 1)
'         Exit For
'      End If
'   Next i
   
   If i <= 0 Then
      RepairDb1 = 1
      Exit Function
   End If
   
   DbName = Mid(DbPath, i + 1)
   DbPath = Left(DbPath, i - 1)
   
   i = DbEnUso(DbPath, DbName)
   If i Then
      RepairDb1 = i
      
      If i = 2 And bMsg Then
         MsgBox1 "La base está en uso, no puede repararse.", vbExclamation
      End If
   
      Exit Function
   End If
   
   

   ' preparamos los nombres de archivos
   DbPathS = DbPath & "\" & DbName     ' base fuente
   
   i = 1
   Do
      DbPathR = DbPath & "\R" & i & "_" & DbName   ' base backup
      If ExistFile(DbPathR) = False Then
         Exit Do
      End If
      i = i + 1
   Loop
   
   ERR.Clear
   FileCopy DbPathS, DbPathR
   If ERR And bMsg Then
      MsgErr "Error al intentar copiar la base de datos antes de repararla." & vbCrLf & DbPathS
      Exit Function
   End If
         
   On Error Resume Next
      
   DBEngine.RepairDatabase DbPathS
   
   If ERR Then
   
      If bMsg Then
         Call MsgErr(NL & "Base en " & DbPathS)
      End If
      
   Else ' Todo OK
'      Kill DbPathR
'      If Err = 0 Or Err = 53 Then ' 53: File not found
'         Err.Clear
'         Name DbPathS As DbPathR ' Dejamos un backup de lo que había
'         If Err = 0 Then
'            Name DbPathD As DbPathS
'         End If
'      End If
      
      If bMsg Then
         If ERR = 0 Then
            MsgBox1 "La base fue reparada con éxito." & vbCrLf & vbCrLf & "La base original quedó en " & DbPathR, vbInformation
         Else
            Call MsgErr("La base no pudo ser reparada.")
         End If
      End If
      
   End If

   RepairDb1 = ERR
   
End Function
#End If


#If DATACON = DAO_CONN Then 'ADO
Public Function RepairDb(ByVal NameMdb As String, Optional ByVal bMsg As Boolean = 1) As Boolean
   Dim nPathMdb As String, oPathMdb As String
   
   RepairDb = False
   
   On Error Resume Next
   
   Call AddLog("RepairDb: " & NameMdb & ", DbEngine.version: " & DBEngine.Version)
   
   If Val(DBEngine.Version) > 3.51 Then
      MsgBox1 "No es posible intentar reparar la base de datos. La versión de la base de datos no lo permite.", vbExclamation
      Exit Function
   
   End If
   
   nPathMdb = NameMdb & "n"
   oPathMdb = NameMdb & "o"
   
   Kill nPathMdb
   
   ERR.Clear
   FileCopy NameMdb, nPathMdb
   If ERR And bMsg Then
      MsgErr "Error al intentar copiar la base de datos" & vbCrLf & NameMdb
      Exit Function
   End If
   
   ERR.Clear
   DBEngine.RepairDatabase nPathMdb
      
   If ERR = 0 Then
      
      Kill oPathMdb
      
      ERR.Clear
      Name NameMdb As oPathMdb
      If ERR And bMsg Then
         MsgErr "Error al renombrar la base" & vbCrLf & NameMdb
         Exit Function
      End If
      
      Name nPathMdb As NameMdb
   
      MsgBox1 "La base de datos fue reparada con éxito.", vbInformation
      
      RepairDb = True
   Else
      
      If bMsg Then
         MsgErr "No fue posible reparar la base de datos, intente comprimirla y luego repararla." & vbCrLf & NameMdb
         Call AddLog("RepairDb: Error al reparar: " & ERR.Number & ", " & ERR.Description & ", " & NameMdb)
      End If
   
      Kill nPathMdb
   
   End If
   
End Function
#End If


' busca el máximo del campo 0
#If DATACON = DAO_CONN Then 'ADO
Public Function GetMax(Db As Database, ByVal Qry As String) As Long
#Else
Public Function GetMax(Db As Connection, ByVal Qry As String) As Long
#End If
   Dim Rs As Recordset

   Set Rs = OpenRs(Db, Qry)
   If Rs.EOF Then
      GetMax = -1
   Else
      GetMax = vFld(Rs(0))
   End If

   Call CloseRs(Rs)

End Function
' Exporta una consulta a un archivo.
' Mrk es una marca que pone al comienzo del archivo para reconocerlo con la función Import table
#If DATACON = DAO_CONN Then 'ADO
Public Function ExportQry(Db As Database, ByVal Qry As String, ByVal FName As String, Optional Mrk As String = "", Optional ByVal bAppend As Boolean = 0, Optional ByVal bTit As Boolean = 0, Optional ByVal FmtDate As String = "") As Long
#Else
Public Function ExportQry(Db As Connection, ByVal Qry As String, ByVal FName As String, Optional Mrk As String = "", Optional ByVal bAppend As Boolean = 0, Optional ByVal bTit As Boolean = 0, Optional ByVal FmtDate As String = "") As Long
#End If
   Dim Rs As Recordset
   Dim n As Long
   Dim Fd As Long
   Dim Buf As String
   Dim i As Integer, nf As Integer
   Dim Dt As Double

   On Error Resume Next

   Fd = FreeFile()
   
   If bAppend Then
      Open FName For Append As #Fd
   Else
      Open FName For Output As #Fd
   End If
   
   If ERR Then
      ExportQry = -ERR
      Call MsgErr(FName)
      Exit Function
   End If
   
   Set Rs = OpenRs(Db, Qry)
   If Rs Is Nothing Then
      ExportQry = -SqlErr
      Close #Fd
      Exit Function
   End If
   
   nf = Rs.Fields.Count - 1
   
   If bAppend = False Or bTit = True Then
      Buf = ""
      For i = 0 To nf
         Buf = Buf & vbTab & Rs(i).Name
      Next i
      Print #Fd, Mrk & Mid(Buf, 2)
   End If
   
   n = 0
   Do Until Rs.EOF
      Buf = ""
      
      For i = 0 To nf
         Select Case Rs(i).Type
            #If DATACON = DAO_CONN Then
            Case dbSingle, dbDouble:
            #Else
            Case adDouble, adSingle, adDecimal:
            #End If
               Buf = Buf & vbTab & Str0(vFld(Rs(i)))
      
            #If DATACON = DAO_CONN Then
            Case dbDate, dbTime, dbBoolean:
            #Else
            Case adDate, adDBDate, adDBTime, adBoolean:
            #End If
               Dt = vFld(Rs(i))
               
               If FmtDate <> "" Then
                  Buf = Buf & vbTab & Format(Dt, FmtDate)
               Else
                  Buf = Buf & vbTab & Str0(Dt)
               End If
            
            Case Else
               Buf = Buf & vbTab & ReplaceStr(vFld(Rs(i)), vbTab, " ") ' por si viene un tab en el dato
         End Select
      
      Next i

      Print #Fd, Mid(Buf, 2)
      n = n + 1
      If (n Mod 500) = 0 Then
         DoEvents
      End If

      Rs.MoveNext
   Loop

   Call CloseRs(Rs)

   Close #Fd

   ExportQry = n

End Function
' Obtiene la fecha actual de la base de datos
#If DATACON = DAO_CONN Then 'ADO
Public Function GetDbNow(Db As Database) As Double
#Else
Public Function GetDbNow(Db As Connection) As Double
#End If
   Dim Rs As Recordset
   Dim n As Double, n1 As Double
   Dim Q1 As String, ConnStr As String

   If Db Is Nothing Then
      GetDbNow = Now
      Exit Function
   End If

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If

   If ConnStr = "" Then    ' Access
      n = Now
      Q1 = Db.Name
      Q1 = Left(Q1, Len(Q1) - 4) & ".ldb"
      n1 = FileDateTime(Q1)
      If n1 > n Then
         n = n1
      End If
            
      n = FileDateTime(Db.Name)
      If n1 > n Then
         n = n1
      End If
      
   Else
      Q1 = "SELECT " & SqlNow(Db) ' & " FROM Param"
      Set Rs = OpenRs(Db, Q1, False)
      If Rs Is Nothing Then
         n = Now
      Else
         n = vFld(Rs(0))
         Call CloseRs(Rs)
      End If
   End If
   
   If n <= 0 Then
      n = Now
   End If
   
   GetDbNow = n
   
End Function
' Permite importar desde un archivo de texto con datos de largo fijo
' El formato de los campos se obtiene de la tabla FormatoImport
' FldSep : el número de caracteres que separan los campos
' Los números que ya tienen punto, hay que poner 0 en Decim
' Decim se usa cuando los números vienen sin simbolo decimal
' por ejemplo 0000245 con dos decimales, es decir, 2.45
' MilDec contiene los separadores que vienen de miles y decimal "md", por ej: ".,"
#If DATACON = DAO_CONN Then 'ADO
Public Function ImportFile(Db As Database, ByVal TabName As String, ByVal FName As String, nRec As Long, LaRec As Label, FInfo As String, Optional TabDest As String = "", Optional ByVal FldSep As Integer = 0, Optional ByVal RecLabelUpdate As Integer = 100, Optional ByVal SkipLines As Integer = 0, Optional ByVal MilDec As String = "") As Long
#Else
Public Function ImportFile(Db As Connection, ByVal TabName As String, ByVal FName As String, nRec As Long, LaRec As Label, FInfo As String, Optional TabDest As String = "", Optional ByVal FldSep As Integer = 0, Optional ByVal RecLabelUpdate As Integer = 100, Optional ByVal SkipLines As Integer = 0, Optional ByVal MilDec As String = "") As Long
#End If
   Dim Rs As Recordset
   Dim QB As String, Buf As String, Fld As String, Q1 As String, Q2 As String
   Dim Tabla() As Campo_t
   Dim i As Integer, n As Integer, j As Integer, Fd As Long, l As Long, Rc As Long, r As Long
   Dim Msg As String
   Dim RecLen As Integer, iFld As Integer, lFld As Integer, bDefUsed As Integer
   Dim UFile As UnixFile_t
'   dim Dcmal As String
   
   ImportFile = 0

   If RecLabelUpdate <= 10 Then
      RecLabelUpdate = 10
   End If

   If ExistFile(FName) = False Then
      Rc = -7
      Msg = "No existe el archivo " & FName & ", Rc= " & Rc
      Call AddLog("ImportFile: " & Msg)
      MsgBox1 Msg, vbExclamation
      ImportFile = Rc
      Exit Function
   End If

   ReDim Tabla(12)

   Q1 = "SELECT Campo, Tipo, Largo, Decim, Formato, Predet, ExCampo FROM FormatoImport WHERE Tabla='" & TabName & "' ORDER BY Orden"
   Set Rs = OpenRs(Db, Q1)
   If Rs Is Nothing Then
      Rc = -6
      Msg = "No hay datos en FormatoImport para " & TabName & ", Rc= " & Rc
      Call AddLog("ImportFile: No hay datos en FormatoImport para " & TabName & ", Rc= " & Rc)
      MsgBox1 Msg, vbExclamation
      ImportFile = Rc
      Exit Function
   End If
   
   n = 0
   RecLen = 0
   QB = ""
   Do Until Rs.EOF
      If n > UBound(Tabla) Then
         ReDim Preserve Tabla(n + 5)
      End If
      
      Tabla(n).Campo = Trim(vFld(Rs("Campo")))
      Tabla(n).Tipo = Trim(vFld(Rs("Tipo")))
      Tabla(n).Largo = vFld(Rs("Largo"))
      Tabla(n).Dec = vFld(Rs("Decim"))
      Tabla(n).Fmt = vFld(Rs("Formato"))
      Tabla(n).Def = vFld(Rs("Predet"))
      Tabla(n).ExCampo = vFld(Rs("ExCampo"))
   
      RecLen = RecLen + Tabla(n).Largo + Tabla(n).Dec
   
      If LCase(Left(Tabla(n).Campo, 5)) <> "_skip" Then ' si parte con _skip no lo importa
         QB = QB & "," & Tabla(n).Campo
      End If
      
      n = n + 1
         
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)

   If n < 1 Then
      Rc = -1
      ImportFile = Rc
      Call AddLog("ImpFile: La tabla FormatoImport no incluye información para la importación de la tabla " & TabName & ", Rc= " & Rc)
      Exit Function
   End If

   If TabDest = "" Then
      TabDest = TabName
   End If

   QB = "INSERT INTO " & TabDest & " (" & Mid(QB, 2) & ") VALUES ("
   n = n - 1
   ReDim Preserve Tabla(n)

   'On Error Resume Next
   UFile.Fd = FreeFile()
   'Open FName For Input As #UFile.Fd
   Open FName For Binary Access Read As #UFile.Fd
   If ERR Then
      MsgErr FName
      Rc = -2
      Call AddLog("ImpFile: No se pudo abrir el archivo " & FName & ", Error " & ERR & ", " & Error & ", Rc= " & Rc)
      ImportFile = Rc
      Exit Function
   End If

   FInfo = FName & ": " & Format(FileDateTime(FName), "dd mmm yyyy hh:nn:ss") & " - " & Format(FileLen(FName), NUMFMT) & " bytes"
   Call AddLog("ImportFile: " & FInfo)

   l = 0
   r = 0
   LaRec = r

   Do Until UnixEoF(UFile)
   
      l = l + 1
      Buf = ""
      Q1 = QB
      
      'Line Input #Fd, Buf
      Buf = UnixLineInput(UFile)

      If l <= SkipLines Then
         GoTo NextRec
      End If

      If Len(Buf) = 1 Then
         If Asc(Buf) = Val("&H1A") And UnixEoF(UFile) Then  ' EOF ?
            Exit Do
         End If
      ElseIf Len(Buf) = 0 Then
         If UnixEoF(UFile) Then   ' EOF ?
            Exit Do
         End If
      End If

      Q2 = ""
      j = 1
      For i = 0 To n
         
         iFld = j
         lFld = Tabla(i).Largo
         Fld = Trim(Mid(Buf, j, Tabla(i).Largo))
         j = j + Tabla(i).Largo
         
         'PS
         If Tabla(i).Tipo = "N" Or Tabla(i).Tipo = "M" Then
            If IsNumeric(Fld) = False And Trim(Fld) = "" Then
            'Se asume número por defecto
               Fld = Tabla(i).Def
            End If
         
            If Tabla(i).Dec Then
               Fld = Fld & "." & Mid(Buf, j, Tabla(i).Dec)
               
               '**PS
   '            Dcmal = Replace(Mid(Buf, j, Tabla(i).Dec), ".", "") 'PS
   '            Dcmal = Replace(Mid(Buf, j, Tabla(i).Dec), ",", "") 'PS
   '            If Trim(Dcmal) = "" Then
   '               'Se asume cero
   '               Dcmal = String(Tabla(i).Dec, "0")
   '            End If
   '            Fld = Fld & "." & Dcmal
               '****
               
               j = j + Tabla(i).Dec
               lFld = lFld + Tabla(i).Dec
            ElseIf MilDec <> "" Then
               If InStr(Fld, Left(MilDec, 1)) Then ' viene separador de miles
                  Fld = Replace(Fld, Left(MilDec, 1), "")
               End If
               
               If InStr(Fld, Mid(MilDec, 2, 1)) And Mid(MilDec, 2, 1) <> "." Then ' viene un sep decimal que no es .
                  Fld = Replace(Fld, Mid(MilDec, 2, 1), ".")
               End If
            Else ' si vienen decimales con , los cambiamos por .
               Fld = Replace(Fld, ",", ".")
            End If
         End If
         
         Select Case Tabla(i).Tipo
            Case "N", "M":   ' Numérico, M: Numerico SAP con signo al final
               If IsNumeric(Fld) = False Then
               
                  Rc = -3
                  ImportFile = Rc
                  Msg = "Línea " & l & ", col. " & i + 1 & " [I" & iFld & "-L" & lFld & "], Campo " & Tabla(i).Campo & ": el valor [" & Fld & "] no es numérico."
                  Call AddLog("ImportFile: " & Msg & " Rc= " & Rc)
                  If MsgBox1(FName & vbCrLf & Msg & vbCrLf & "¿ Continua ?, se asumirá cero.", vbExclamation Or vbYesNo Or vbDefaultButton2) <> vbYes Then
                     Close Fd
                     Exit Do
                  End If
                  Rc = 0
               End If
               
               If Tabla(i).Tipo = "M" Then
                  If Right(Fld, 1) = "-" Then ' Formato SAP
                     Fld = "-" & Left(Fld, Len(Fld) - 1)
                  End If
               End If
               
               Fld = Str0(Val(Fld)) ' tiene que se Val( ) !!
               
            Case "A":   ' Alfanumérico
               Fld = "'" & ParaSQL(Fld) & "'"
            
            Case "F":   ' Fecha
               If Fld <> String(Tabla(i).Largo, "0") And Trim(Fld) <> "" Then
                        
                  Rc = GetFixDate(Fld, Tabla(i).Fmt, Tabla(i).Def, bDefUsed)
                  If Rc < 0 Then
                     Rc = -3
                     ImportFile = Rc
                     Msg = "Línea " & l & ", col. " & i + 1 & " [I" & iFld & "-L" & lFld & "], Campo " & Tabla(i).Campo & ": el valor [" & Fld & "] no es una fecha."
                     Call AddLog("ImportFile: " & Msg & " Rc=" & Rc)
                     If MsgBox1(FName & vbCrLf & Msg & vbCrLf & "¿ Continúa ? Se asumirá '1 ene 1900'.", vbExclamation Or vbYesNo Or vbDefaultButton2) <> vbYes Then
                        Exit Do
                     End If
                     
                     bDefUsed = 0
                     Rc = 0
                  End If
                  Fld = Rc
                  
                  If bDefUsed <> 0 And Tabla(i).ExCampo <> "" Then
                     Q1 = ReplaceStr(Q1, "," & Tabla(i).Campo & ",", "," & Tabla(i).Campo & "," & Tabla(i).ExCampo & ",")
                     Fld = Fld & ", 1"
                  End If
                  
               Else
                  Fld = 0
               End If
               
            Case Else:
               Msg = "Línea " & l & ", col. " & i + 1 & " [I" & iFld & "-L" & lFld & "], Campo " & Tabla(i).Campo & ": el tipo [" & Tabla(i).Tipo & "] no es soportado "
               Call AddLog("ImpFile: " & Msg)
               Debug.Print "ImportFile: " & Msg
               
         End Select
         
         If LCase(Left(Tabla(i).Campo, 5)) <> "_skip" Then ' si parte con _skip no lo importa
            Q2 = Q2 & "," & Fld
         End If
         
         j = j + FldSep
      
      Next i
   
      Buf = Q1 & Mid(Q2, 2) & ")"
      Rc = ExecSQL(Db, Buf)
      If SqlErr Then
         Rc = -4
         ImportFile = Rc
         Msg = "Línea " & l & ": Error al insertar. Rc= " & Rc
         Call AddLog("ImportFile: " & Msg)
         If MsgBox1(FName & vbCrLf & Msg & vbLf & "¿Continúa?", vbExclamation Or vbYesNo Or vbDefaultButton2) <> vbYes Then
            Close (Fd)
            Exit Do
         End If
         
      ElseIf Rc < 0 Then
         Rc = -5
         ImportFile = Rc
         Msg = "Línea " & l & ": Error al insertar, puede ser que no cumpla con la llave única. Rc= " & Rc
         Call AddLog("ImportFile: " & Msg)
         If MsgBox1(FName & vbCrLf & Msg & vbLf & "¿Continúa?", vbExclamation Or vbYesNo Or vbDefaultButton2) <> vbYes Then
            Close (Fd)
            Exit Do
         End If
      Else
         r = r + 1
      End If
      
      If l Mod RecLabelUpdate = 0 Then
         LaRec = Format(r, NUMFMT)
         DoEvents
      End If

NextRec:
   Loop
   
   Close #UFile.Fd
   'File.Close
   
   Call AddLog("ImportFile: Se importaron " & r & " registros en la tabla " & TabName & ".")
   
   nRec = r
   LaRec = Format(r, NUMFMT)
   
End Function


' Importa un archivo de datos separados por Tabs
' Retorna el número de regitros importados
#If DATACON = DAO_CONN Then 'ADO
Public Function ImportTable(Db As Database, ByVal TbName As String, ByVal Flds As String, ByVal FName As String, Optional ByVal LSkip As Byte = 1, Optional ByVal bAcceptNull As Boolean = True, Optional Mrk As String = "") As Long
#Else
Public Function ImportTable(Db As Connection, ByVal TbName As String, ByVal Flds As String, ByVal FName As String, Optional ByVal LSkip As Byte = 1, Optional ByVal bAcceptNull As Boolean = True, Optional Mrk As String = "") As Long
#End If
   Dim Rs As Recordset
   Dim Q1 As String, Q2 As String, sValor As String, Msg As String, Buf As String, MsgNull As String
   Dim nf As Integer, i As Integer, F As Integer
   Dim Fd As Long, Nr As Long, l As Long, Rc As Long
   Dim Valor As Double
   
   Q1 = "SELECT " & Flds & " FROM " & TbName
   Set Rs = OpenRs(Db, Q1)
   nf = Rs.Fields.Count - 1
   ReDim FldType(nf) As Integer
   ReDim FldName(nf) As String

   For F = 0 To nf
      FldType(F) = Rs.Fields(F).Type
      FldName(F) = Rs.Fields(F).Name
   Next F
   
   Call CloseRs(Rs)

   Q1 = "INSERT INTO " & TbName & " ( " & Flds & ") VALUES ("

   Fd = FreeFile()
   Open FName For Input As #Fd

   For i = 1 To LSkip
      Line Input #Fd, Buf
      If i = 1 And Mrk <> "" Then
         If Left(Buf, Len(Mrk)) <> Mrk Then
            Close #Fd
            ImportTable = -1
            Exit Function
         End If
      End If
   Next i

   Msg = "Error en el archivo " & FName
   l = LSkip

   MsgNull = ""

   Do Until EOF(Fd)
      Line Input #Fd, Buf
      l = l + 1
      Q2 = ""
      i = 1
      
      For F = 0 To nf
      
         If bAcceptNull = False Then
            MsgNull = "Linea " & l & ": Falta dato para el campo " & FldName(F)
         End If
         
         #If DATACON = DAO_CONN Then
         If FldType(F) = dbChar Or FldType(F) = dbText Then
         #Else
         If FldType(F) = adChar Or FldType(F) = adVarChar Then
         #End If
            
            'If LineGetStr(Buf, i, msg, MsgNull, SValor, (bAcceptNull = False And f = nf)) Then
            sValor = GetBufStr(Buf)
            If bAcceptNull = False And sValor = "" Then
               MsgBox1 MsgNull, vbExclamation
               If Buf = "" Then
                  GoTo NextLine
               End If
            Else
               Q2 = Q2 & ", '" & ParaSQL(sValor) & "'"
            End If
         Else
'            If LineGetVal(Buf, i, msg, MsgNull, Valor, (bAcceptNull = False And f = nf)) Then
            sValor = GetBufStr(Buf)
            Valor = Val(sValor)
            If bAcceptNull = False And sValor = "" Then
               MsgBox1 MsgNull, vbExclamation
               If Buf = "" Then
                  GoTo NextLine
               End If
            Else
               Q2 = Q2 & ", " & str(Valor)
            End If
         End If
      Next F
      
      Rc = ExecSQL(Db, Q1 & Mid(Q2, 2) & " )")
      If Rc > 0 Then
         Nr = Nr + 1
         
         If l Mod 50 = 0 Then
            DoEvents
         End If
      End If
      
NextLine:
   Loop

   Close #Fd

   ImportTable = Nr

End Function
' Importa un archivo que debe tener los campos en el mismo orden que la tabla, separados por Tabs
' Se puede usar con ExportQry. Se salta la primera línea
#If DATACON = DAO_CONN Then 'ADO
Public Function ImportUsingBCP(DbTo As Database, ByVal Table As String, ByVal FName As String) As Long
#Else
Public Function ImportUsingBCP(DbTo As Connection, ByVal Table As String, ByVal FName As String) As Long
#End If
   Dim Rc As Long
   Dim Q1 As String

   Q1 = "BULK INSERT " & DbTo.Name & ".." & Table
   Q1 = Q1 & " FROM '" & FName & "'"
   Q1 = Q1 & " WITH ("
   Q1 = Q1 & " BATCHSIZE = 1000"
   Q1 = Q1 & ", FIRSTROW = 2"
   Q1 = Q1 & ", FIELDTERMINATOR = '\t'"
   Q1 = Q1 & ", ROWTERMINATOR = '\n'"
   Q1 = Q1 & " )"
      
   Rc = ExecSQL(DbTo, Q1)
   
'BULK INSERT [['nombreBaseDatos'.]['propietario'].]{'nombreTabla' FROM archivoDatos}
'[WITH
'(
'[ BATCHSIZE [= tamañoLote]]
'[[,] CHECK_CONSTRAINTS]
'[[,] CODEPAGE [= 'ACP' | 'OEM' | 'RAW' | páginaCódigos']]
'[[,] DATAFILETYPE [=
'{'char' | 'native'| 'widechar' | 'widenative'}]]
'[[,] FIELDTERMINATOR [= 'terminadorCampo']]
'[[,] FIRSTROW [= primeraFila]]
'[[,] FORMATFILE [= 'RutaArchivoFormato']]
'[[,] KEEPIDENTITY]
'[[,] KEEPNULLS]
'[[,] KILOBYTES_PER_BATCH [= kilobytesPorLote]]
'[[,] LASTROW [= últimaFila]]
'[[,] MAXERRORS [= erroresMáximos]]
'[[,] ORDER ({columna [ASC | DESC]} [,.n])]
'[[,] ROWS_PER_BATCH [= filasPorLote]]
'[[,] ROWTERMINATOR [= 'terminadorFila']]
'[[,] TABLOCK]
')
']


End Function
#If DATACON = DAO_CONN Then
Function AlterField(Db As Database, ByVal TblName As String, ByVal FldName As String, ByVal FldType As Integer, Optional ByVal FldSize As Integer = -1, Optional ByVal SetStr As String = "", Optional ByVal NewSets As Byte = 0, Optional ByVal DefValue As String = "", Optional ByVal bRequired As Boolean = 0, Optional ByVal bAllowZeroLen As Boolean = 1) As Integer
   Dim Q1 As String, Q2 As String, FldName1 As String, IdxName As String
   Dim Rc As Long
   Dim Tbl As TableDef
   Dim Fld As Field, OFld As Field, iFld As Field
   Dim Rs As Recordset
   Dim Idx As Index, nIdx As Index

   Call AddDebug("2038: >> AlterField: '" & TblName & "." & FldName & "'.")

   On Error Resume Next

   AlterField = -3

   Set Tbl = Db.TableDefs(TblName)

   Set OFld = Tbl.Fields(FldName)
   
   If OFld Is Nothing Then
      Call AddLog("2049: AlterField: No existe el campo " & TblName & "." & FldName)
      Debug.Print "No existe el campo " & TblName & "." & FldName
      Exit Function
   End If
   
   If OFld.Type = FldType Then
      If FldSize > 0 Then
         If OFld.Size = FldSize Then
            Call AddDebug("2057: << AlterField: No cambia el tipo ni tamaño.")
            Exit Function
         End If
      Else
         Call AddDebug("2061: << AlterField: No cambia el tipo.")
         Exit Function
      End If
   End If
         
   FldName1 = FldName & "_" & Format(Now, "hnnss") & "_"
   
   ERR.Clear
   If FldSize > 0 Then
      Set Fld = Tbl.CreateField(FldName1, FldType, FldSize)
   Else
      Set Fld = Tbl.CreateField(FldName1, FldType)
   End If
   
   If ERR Then
      Call AddLog("2076: AlterField: Error " & ERR & ", " & ERR.Description & ", al crear el campo '" & TblName & "." & FldName1 & "'.")
   End If
   
   Fld.OrdinalPosition = OFld.OrdinalPosition
   
   If (NewSets And 1) <> 0 Then
      Fld.DefaultValue = DefValue
   Else
      Fld.DefaultValue = OFld.DefaultValue
   End If
   
   If (NewSets And 2) <> 0 Then
      Fld.Required = bRequired
   Else
      Fld.Required = OFld.Required
   End If

   If (NewSets And 4) <> 0 Then
      Fld.AllowZeroLength = bAllowZeroLen
   Else
      Fld.AllowZeroLength = OFld.AllowZeroLength
   End If
   
   ERR.Clear
   Tbl.Fields.Append Fld

   If ERR = 0 Then
      Tbl.Fields.Refresh
   Else
      AlterField = ERR.Number
      Call AddLog("2106: AlterField: Error " & ERR.Number & ", " & ERR.Description & ", en Append para '" & TblName & "." & FldName1 & "'.")
      Exit Function
   End If
      
   If SetStr = "" Then
      SetStr = FldName
   End If
   
   Q1 = "UPDATE " & TblName & " SET " & FldName1 & " = " & SetStr
   Rc = ExecSQL(Db, Q1)
   
   If Rc < 0 Then
      Call AddLog("2118: AlterField: Update, Rc=" & Rc & ", para '" & TblName & "." & FldName & ".")
      AlterField = Rc
      Exit Function
   End If
   
   ERR.Clear
   For Each Idx In Tbl.Indexes
      For Each iFld In Idx.Fields
         If UCase(iFld.Name) = UCase(FldName) Then
   
            Set nIdx = Tbl.CreateIndex(Idx.Name)
            nIdx.Fields = ReplaceStr(Idx.Fields, FldName, FldName1)
            nIdx.IgnoreNulls = Idx.IgnoreNulls
            nIdx.Primary = Idx.Primary
            nIdx.Required = Idx.Required
            nIdx.Unique = Idx.Unique
   
            Tbl.Indexes.Delete Idx.Name
            Tbl.Indexes.Refresh
            Tbl.Indexes.Append nIdx
            Tbl.Indexes.Refresh
      
            Set Idx = nIdx
            Exit For
         End If
      Next iFld
   
   Next Idx

   ERR.Clear

   Set OFld = Nothing
   Tbl.Fields.Delete FldName
   Tbl.Fields.Refresh
      
   Fld.Name = FldName

   Tbl.Fields.Refresh

   Set Fld = Nothing
   Set Tbl = Nothing

   If ERR.Number Then
      Call AddLog("2161: AlterField: Error " & ERR.Number & ", " & ERR.Description & ", para '" & TblName & "." & FldName & "'.")
   Else
      Call AddDebug("2163: << AlterField: " & TblName & "." & FldName)
   End If

   AlterField = ERR.Number

End Function
#End If

' Para que esta funcion opere con ODBC se debe crear el siguiente SP
'CREATE PROCEDURE [Sp_AddNew]
'   @Tabla Varchar(20) ,
'   @Campo Varchar(20)
'
'  AS
'
'  Begin
'    declare @Rc Int
'
'   set @Rc = 0
'
'   Exec('INSERT INTO ' + @Tabla + '  ( ' + @Campo + ' ) VALUES ( 0 )' )
'
'   If @@error = 0
'   begin
'      select @@Identity AS ID
'      set @Rc = 0
'   End
'   Else
'   begin
'      set @Rc = @@Error
'   End
'
'   return( @Rc)
'
'  End
'
#If DATACON = DAO_CONN Then 'ADO
Public Function TbAddNew(Db As Database, ByVal Tbl As String, ByVal FldId As String, ByVal FldName As String, Optional ByVal Value2FldName As String = "") As Long
#Else
Public Function TbAddNew(Db As Connection, ByVal Tbl As String, ByVal FldId As String, ByVal FldName As String, Optional ByVal Value2FldName As String = "") As Long
#End If
'Public Function TbAddNew(Db As Database, ByVal Tbl As String) As Long
   Dim Q1 As String ' , FldName As String
   Dim Rs As Recordset
   Dim Fld As Field
   Dim ConnStr As String

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect  ' OJO: si la tabla es Access y la tabla está linkeada con clave, podría funcionar, mejor usar TbAddNew3
#Else
   ConnStr = Db.ConnectionString
#End If

   On Error Resume Next

   TbAddNew = -1

   If Left(ConnStr, 5) <> "ODBC;" Then
      'For Each Fld In Db.TableDefs(Tbl).Fields
      '   If Fld.DataUpdatable Then ' es autonumber
      '      FldName = Fld.Name
      '      Exit For
      '   End If
      'Next Fld
            
      #If DATACON = DAO_CONN Then
'      Set Rs = Db.OpenRecordset(Tbl, dbOpenTable)
      Set Rs = DbOpenTable1(Db, Tbl)
      #Else
      Set Rs = Db.Execute(Tbl, adCmdTable)
      #End If
      
      If Rs Is Nothing Then
         SqlErr = ERR
         SqlError = ERR.Description
         Call AddLog("TbAddNew: error: " & SqlErr & ", " & SqlError & ", en OpenRecorset/Execute con " & Tbl)
         MsgBox1 "Error " & SqlErr & ", " & SqlError, vbExclamation
         TbAddNew = -1
         Exit Function
      End If
      
      Rs.AddNew
      TbAddNew = vFld(Rs(FldId))
      
      If FldName <> "" And Value2FldName <> "" Then ' por si tiene un índice que no permite null en FldName
         Rs(FldName) = Value2FldName
      End If
      
      Rs.Update
      
      If ERR Then
         SqlErr = ERR
         SqlError = ERR.Description
         Call AddLog("TbAddNew: " & Tbl & ", error: " & SqlErr & ", " & SqlError)
         TbAddNew = -1
      End If
      Call CloseRs(Rs)
   Else
      'For Each Fld In Db.TableDefs(Tbl).Fields
      '   If Fld.DataUpdatable = False Then ' no es autonumber
      '      FldName = Fld.Name
      '      Exit For
      '   End If
      'Next Fld
         
      Q1 = "SP_AddNew " & Tbl & ", " & FldName
      Set Rs = OpenRs(Db, Q1)
      TbAddNew = vFld(Rs("Id"))
      Call CloseRs(Rs)
   End If

End Function

' Para que esta funcion opere con ODBC se debe crear el siguiente SP
'CREATE PROCEDURE [Sp_AddNew]
'   @SqlCmd Varchar(1000)
'  AS
'
'  Begin
'    declare @Rc Int
'
'   set @Rc = 0
'
'   Exec( @SqlCmd )
'   --Exec( 'INSERT INTO Empleados (RutEmpleado) VALUES (''1-9'')'  )
'
'   If @@error = 0
'   begin
'      select @@Identity AS ID
'      set @Rc = 0
'   End
'   Else
'   begin
'      set @Rc = @@Error
'   End
'
'   return( @Rc)
'
'  End
'
' Requiere que exista el procedimiento almacenado Sp_AddNew (ver arriba), pero se debe probar que funcione
#If DATACON = DAO_CONN Then 'ADO
Public Function TbAddNew2(Db As Database, ByVal InsertCmd As String) As Long
#Else
Public Function TbAddNew2(Db As Connection, ByVal InsertCmd As String) As Long
#End If
   Dim Rs As Recordset
   Dim Fld As Field
   Dim Q1 As String, ConnStr As String
   Dim id As Long, sErr As String

'#If DATACON = DAO_CONN Then
'   ConnStr = Db.Connect
'#Else
'   ConnStr = Db.ConnectionString
'#End If

   On Error Resume Next

   TbAddNew2 = -22

   If SqlType(Db, ConnStr) = SQL_SERVER Then
'   If Left(ConnStr, 5) = "ODBC;" Or InStr(1, ConnStr, "Sql Server", vbTextCompare) <> 0 Then
      
      If InStr(1, InsertCmd, "NULL", vbTextCompare) Then
         Debug.Print "*** TbAddNew2: NO SE PUEDEN PONER NULL"
         InsertCmd = ReplaceStr(InsertCmd, "NULL", "' '")
      End If
      
      InsertCmd = ReplaceStr(InsertCmd, "'", "''")
      Q1 = "SP_AddNew '" & InsertCmd & "'"
      Set Rs = OpenRs(Db, Q1, , , , , False)
            
      If Not Rs Is Nothing Then
         
         If Rs.EOF = False Then
            id = vFld(Rs("ID"))
'            sErr = vFld(Rs("Error"))
         
            TbAddNew2 = id
         ElseIf Len(Q1) > 250 Then
            Debug.Print "*** Verificar el largo del comando en el SP ****"
         End If
         Call CloseRs(Rs)
      End If
   End If

End Function
' NO Requiere que exista el procedimiento almacenado Sp_AddNew
#If DATACON = DAO_CONN Then 'ADO
Public Function TbAddNew3(Db As Database, ByVal InsertCmd As String, ByVal SelCmd As String) As Long
#Else
Public Function TbAddNew3(Db As Connection, ByVal InsertCmd As String, ByVal SelCmd As String) As Long
#End If
   Dim Rs As Recordset
   Dim Fld As Field
   Dim Q1 As String, Valor As Long, Rc As Long
   Dim id As Long, sErr As String
   Dim i As Integer, n As Integer
   Static bInit As Boolean
   
   If bInit = False Then
      Randomize Now
      bInit = True
   End If

   id = -22
   TbAddNew3 = id

   n = 1 + Rnd() * 5
   For i = 0 To n
      Valor = 10000000# + i + CDbl(Now) * 10000# + (Rnd() * (2 ^ 28))
   Next i
   
   Q1 = ReplaceStr(InsertCmd, "%_RND_%", Valor)
   Rc = ExecSQL(Db, Q1)
   
   If Rc < 0 And InStr(InsertCmd, "%_RND_%") > 0 Then
      Exit Function
   End If
   
   Q1 = ReplaceStr(SelCmd, "%_RND_%", Valor)
   Set Rs = OpenRs(Db, Q1)
   If Rs.EOF = False Then
      id = vFld(Rs(0))
   End If
   Call CloseRs(Rs)
   
   TbAddNew3 = id
   
End Function
' Esto sirve solo para SQL Server
#If DATACON = DAO_CONN Then 'ADO
Public Function TbAddNew4(Db As Database, ByVal InsQry As String, ByVal idFld As String, Optional ByVal Msg As Boolean = False) As Long

#Else
Public Function TbAddNew4(Db As Connection, ByVal InsQry As String, ByVal idFld As String, Optional ByVal Msg As Boolean = False) As Long
#End If
   Dim Rs As Recordset

   If gDbType <> SQL_SERVER Then
      Debug.Print "La base no corresponde."
      Exit Function
   End If
   
   If W.InDesign Then
      If InStr(1, InsQry, " VALUES ", vbTextCompare) <= 0 Then
         Debug.Print "*** TbAddNew4: No se encontró ' VALUES ' en InsQry ***"
      End If
   End If
   
   InsQry = ReplaceStr(InsQry, " VALUES ", " OUTPUT Inserted." & idFld & " VALUES ")

   Set Rs = OpenRs(Db, InsQry, Msg)
   If Not Rs Is Nothing Then
      If Rs.EOF = False Then
         TbAddNew4 = Rs(0)
      End If
      Call CloseRs(Rs)
   End If

End Function

' Abre bases Access o DBF
#If DATACON = DAO_CONN Then
Public Function OpenDbfDb(ByVal DbPath As String) As Database
   Dim Rs As Recordset
   Dim Db As Database
   
   Set OpenDbfDb = Nothing
   
   On Error Resume Next
   
   DbPath = LCase(Trim(DbPath))
      
   If DbPath = "" Then
      Exit Function
   End If
   
   If Right(DbPath, 4) = ".mdb" Then
      Set Db = OpenDatabase(DbPath, False, False)
   Else ' DBF
      Set Db = OpenDatabase(DbPath, False, False, "dBASE IV;")
            
   End If
   
   If ERR = 0 Then
      Set OpenDbfDb = Db
   Else
      MsgErr DbPath
   End If

End Function
#End If
#If DATACON = DAO_CONN Then 'dao
Public Function LinkDbfTable(Db As Database, ByVal TbPath As String, ByVal TbFile As String, Optional ByVal NewName As String = "", Optional DbType As String = "dBASE IV ", Optional ByVal bForce As Boolean = 0, Optional bMsg As Boolean = 1) As Boolean
   Dim Tbl As TableDef
   Dim i As Integer
   Dim TableName As String, ConnStr As String, Msg As String
   
   On Error Resume Next

   LinkDbfTable = False
   
   TbFile = Trim(TbFile)
   
   i = InStr(TbFile, ".")
   If i Then
      TableName = Left(TbFile, i - 1)
   Else
      TableName = TbFile
   End If

   If Len(TableName) > 8 Then
      MsgBox1 "El nombre del archivo '" & TableName & "' supera los 8 caracteres.", vbExclamation
      LinkDbfTable = False
      Exit Function
   End If

   If Trim(NewName) = "" Then
      NewName = TableName
   End If
   
   If Db.TableDefs(NewName).Connect = "" Then ' No es una tabla linkeada, se perderian datos
      If ERR = 0 Then
         LinkDbfTable = False
         Exit Function
      End If
      ERR.Clear
   End If
  
   ConnStr = DbType & ";DATABASE=" & AbsPath(TbPath)

   Set Tbl = Db.TableDefs(NewName)

   If Not Tbl Is Nothing Then

      If bForce = False Then

         Tbl.RefreshLink

         If ERR = 0 And StrComp(Tbl.Connect, ConnStr, vbTextCompare) = 0 And StrComp(Tbl.SourceTableName, TbFile, vbTextCompare) = 0 Then
            If ExistFile(TbPath & "\" & TbFile) Then
               LinkDbfTable = True
               Exit Function
            End If
         End If
         
      End If

   End If

   Db.TableDefs.Delete NewName
   ERR.Clear
   
   Set Tbl = New TableDef
   Tbl.Connect = ConnStr
   Tbl.SourceTableName = TbFile
   Tbl.Name = NewName
   
   Db.TableDefs.Append Tbl
   Db.TableDefs.Refresh
   
   LinkDbfTable = (ERR = 0)
   
   If ERR Then
      Msg = "Error al vincular el archivo '" & TbPath & "\" & TbFile & "'."
      
      If bMsg Then
         MsgErr Msg
      End If
      Debug.Print "FALLÓ LinkDbf(" & TbPath & "\" & TbFile & ") Err=" & ERR & ", " & Error
      
      Call AddLog("LinkDbf: " & Msg & " Err=" & ERR & ", " & Error)
   
   End If
      
End Function
#End If
#If DATACON = DAO_CONN Then 'dao
' OJO: ConnString no debe tener un ; en el primer caracter
Public Function LinkMdbTable(Db As Database, ByVal MdbPath As String, ByVal TableName As String, Optional ByVal NewName As String = "", Optional bForce As Boolean = 0, Optional bMsg As Boolean = 1, Optional ConnString As String = "", Optional ByVal bForceIfNotLinked As Boolean = False) As Boolean
   Dim Tbl As TableDef
   Dim i As Integer
   Dim ConnStr As String, Msg As String, Conn1 As String, TConnect As String, TPWD As String
   Dim Q1 As String
   
   On Error Resume Next
   
   LinkMdbTable = False
   
   If Trim(NewName) = "" Then
      NewName = TableName
   End If
      
   If Left(ConnString, 1) = ";" Then
      ConnString = Mid(ConnString, 2)
   End If
      
   'FCA: se agrega verificación de bForce para que lo haga de todas maneras, si se requiere
      
   If Db.TableDefs(NewName).Connect = "" And Not bForceIfNotLinked Then ' No es una tabla linkeada, se perderian datos
      If ERR = 0 Then
         LinkMdbTable = False
         Exit Function
      End If
      ERR.Clear

   End If
   
   Set Tbl = Db.TableDefs(NewName)

   If Not Tbl Is Nothing Then

      If bForce = False Then
      
         TConnect = Tbl.Connect
         TPWD = "PWD=" & GetTxConnectInfo(TConnect, "PWD") & ";"
      
         ' Si estaba linkeado sin clave pero ahora la base tiene clave, y la clave es diferente => FORCE
         If ConnString <> "" And StrComp(TPWD, ConnString, vbTextCompare) <> 0 Then
            bForce = True
         Else
         
            ' El Count sirve para verificar si está linkeada con la password correcta
            
            If ERR = 0 And StrComp(GetTxConnectInfo(TConnect, "DATABASE"), AbsPath(MdbPath), vbTextCompare) = 0 _
               And StrComp(Tbl.SourceTableName, TableName, vbTextCompare) = 0 Then
               
               If SameMdb(GetTxConnectInfo(TConnect, "DATABASE"), MdbPath, True) Then ' por si cambiaron la unidad y la base sigue existiendo  z:\datos\lpremu.mdb
'               If ExistFile(MdbPath) Then
                  LinkMdbTable = True
                  Exit Function
               End If
               
            End If
            
         End If
         
      End If
            
   End If
   
   Debug.Print "LinkMdbTable: relinkeando " & NewName & " - psw=" & (ConnString <> "")
   
   ConnStr = ";DATABASE=" & AbsPath(MdbPath) & ";" & ConnString

'se cambia dbmain por db gcb201021
   Q1 = "DROP TABLE " & NewName
   Call ExecSQL(Db, Q1, W.InDesign)

   ERR.Clear
   Db.TableDefs.Delete NewName   ' Si ya existía, la eliminamos
   If ERR.Number <> 0 And ERR.Number <> 3265 And ERR.Number <> 3011 Then  ' si no existe, todo bien
      Msg = "Error al eliminar la tabla vinculada '" & NewName & "'."

      If bMsg Then
         MsgErr Msg
      End If
   
      Call AddLog("LinkMdb: " & Msg & " Err=" & ERR & ", " & Error)
      LinkMdbTable = False
      
      Exit Function
   End If
   
   ERR.Clear
   
   Set Tbl = New TableDef
   Tbl.Connect = ConnStr
   Tbl.SourceTableName = TableName
   Tbl.Name = NewName
   
   Db.TableDefs.Append Tbl
   Db.TableDefs.Refresh
   
   LinkMdbTable = (ERR = 0)
   
   If ERR Then
      Msg = "Error al vincular la tabla '" & TableName & "' ubicada en " & MdbPath & "."

      If bMsg Then
         MsgErr Msg
      End If
      
      Debug.Print "FALLÓ LinkMdb(" & TableName & ") Err=" & ERR & ", " & Error
      
      Call AddLog("LinkMdb: " & Msg & " Err=" & ERR & ", " & Error)
   
   End If
      
End Function
#End If
#If DATACON = DAO_CONN Then 'dao
Public Function UnLinkTable(Db As Database, ByVal TableName As String) As Boolean
   
   On Error Resume Next
      
   If Db.TableDefs(TableName).Connect = "" Then ' No es una tabla linkeada, se perderian datos
      If ERR = 0 Then
         UnLinkTable = False
         Exit Function
      ElseIf ERR = 3265 Then
         UnLinkTable = True
         Exit Function
      End If
      ERR.Clear

   End If
   
   Db.TableDefs.Delete TableName   ' Si ya existía, la eliminamos
   
   UnLinkTable = (ERR = 0)
            
End Function
#End If
#If DATACON = DAO_CONN Then 'dao
Public Function IsLinkedTable(Db As Database, ByVal TableName As String) As Boolean
   Dim ConnStr As String
   
   On Error Resume Next
   
   ConnStr = Db.TableDefs(TableName).Connect
      
   IsLinkedTable = InStr(1, ConnStr, "DATABASE=", vbTextCompare) > 0
            
End Function
#End If
#If DATACON = DAO_CONN Then 'dao
' ConnString debe ser del tipo: "DSN=GestSap;UID=usuario;PWD=clave;"
Public Function LinkODBCTable(Db As Database, ByVal ConnString As String, ByVal TableName As String, Optional ByVal NewName As String = "", Optional bMsg As Boolean = 1) As Boolean
   Dim Tbl As TableDef
   Dim i As Integer, Msg As String
   
   On Error Resume Next
   
   LinkODBCTable = False
   
   If Trim(NewName) = "" Then
      NewName = TableName
   End If
   
   If Db.TableDefs(NewName).Connect = "" Then ' No es una tabla linkeada, se perderian datos
      If ERR = 0 Then
         LinkODBCTable = False
         Exit Function
      End If
      ERR.Clear
   End If
   
   Db.TableDefs.Delete NewName   ' Si ya existía, la eliminamos
   ERR.Clear
   
   Set Tbl = New TableDef
   Tbl.Connect = "ODBC;" & ConnString
   Tbl.SourceTableName = TableName
   Tbl.Name = NewName

   Db.TableDefs.Append Tbl
   Db.TableDefs.Refresh
   
   LinkODBCTable = (ERR = 0)

   If ERR Then
      Msg = "Error al vincular la tabla '" & TableName & "'."

      If bMsg Then
         MsgErr Msg
      End If
   
      Debug.Print "FALLO LinkODBC(" & TableName & ") Err=" & ERR & ", " & Error
   
      Call AddLog("LinkODBC: " & Msg & " Err=" & ERR & ", " & Error)
   
   End If

End Function
#End If

#If DATACON = DAO_CONN Then 'dao
Public Function NormalizeCodSQL(Db As Database, ByVal TableName As String, ByVal FldName As String, ByVal Ln As Integer, Optional ByVal Wh As String = "") As Long
#Else
Public Function NormalizeCodSQL(Db As Connection, ByVal TableName As String, ByVal FldName As String, ByVal Ln As Integer, Optional ByVal Wh As String = "") As Long
#End If
   Dim Q1 As String, Rc As Long

   Q1 = "UPDATE " & TableName
   Q1 = Q1 & " SET " & FldName & " = right('" & String(Ln + 2, "0") & "' + LTrim(RTrim(" & FldName & ")), " & Ln & ")"
   Q1 = Q1 & " WHERE IsNumeric(" & FldName & ") = 1"  ' sólo si es número
   If Wh <> "" Then
      Q1 = Q1 & " AND " & Wh
   End If
   
   Rc = ExecSQL(Db, Q1)

   NormalizeCodSQL = Rc

End Function



#If DATACON = DAO_CONN Then 'ADO
Public Function GetNewAutoNumber(Db As Database, ByVal TblName As String, ByVal idFldName As String, ByVal AuxTxtFld As String, Optional ByVal AuxIntFld As String = "") As Long
#Else
Public Function GetNewAutoNumber(Db As Connection, ByVal TblName As String, ByVal idFldName As String, ByVal AuxTxtFld As String, Optional ByVal AuxIntFld As String = "") As Long
#End If
   Dim Q1 As String, Rs As Recordset
   Dim Rc As Long
   Dim Valor As Long, sValor As String

   Valor = 1000000 + (CDbl(Now) - Int(Now)) * 10000000 + (Rnd() * (2 ^ 27))
   sValor = "[" & Valor & "$"

   If AuxTxtFld <> "" Then
      Q1 = "INSERT INTO " & TblName & " (" & AuxTxtFld & ") values ( '" & sValor & "' )"
   Else
      Q1 = "INSERT INTO " & TblName & " (" & AuxIntFld & ") values ( " & Valor & " )"
   End If
   Rc = ExecSQL(Db, Q1)

   If AuxTxtFld <> "" Then
      Q1 = "SELECT " & idFldName & " as __id FROM " & TblName & " WHERE " & AuxTxtFld & " = '" & sValor & "'"
   Else
      Q1 = "SELECT " & idFldName & " as __id FROM " & TblName & " WHERE " & AuxIntFld & " = " & Valor
   End If
   Set Rs = OpenRs(Db, Q1)
   If Rs.EOF = False Then
      GetNewAutoNumber = vFld(Rs("__id"))
   Else
      GetNewAutoNumber = -1
   End If
   Call CloseRs(Rs)
   
End Function

' Funcion util para agregar un registro que debe tener un campo único (no AutoNumber)
' UFldName: nombre del campo único (numérico)
' UValue:   valor unico a poner en el campo UFldName, si es un string pasarlo como "'valor'"
' FldName:  otro campo long que se utiliza para verificar la unicidad del registro
' sWhere    condicion, incluyendo AND, que define la unicidad del campo UFldName
#If DATACON = DAO_CONN Then 'ADO
Public Function AddUniqueRecord(Db As Database, ByVal TblName As String, ByVal UFldName As String, ByVal UValue As String, ByVal IFldName As String, Optional ByVal sWhere As String = "", Optional ByVal SFldName As String = "") As Boolean
#Else
Public Function AddUniqueRecord(Db As Connection, ByVal TblName As String, ByVal UFldName As String, ByVal UValue As String, ByVal IFldName As String, Optional ByVal sWhere As String = "", Optional ByVal SFldName As String = "") As Boolean
#End If
   Dim Q1 As String, SValue As String
   Dim Rc As Long, RNum As Long
   Dim Rs As Recordset
   
   AddUniqueRecord = False
   
   Call PamRandomize
   
   RNum = Rnd() * (2 ^ 28)
   
   
   'Debug.Print "RNum = " & RNum
      
   Q1 = "INSERT INTO " & TblName & "( " & UFldName & ", " & IFldName
   
   If SFldName <> "" Then
      Q1 = Q1 & ", " & SFldName
      SValue = W.PcName
   End If
   
   Q1 = Q1 & " ) VALUES ( " & UValue & ", " & RNum
   
   If SFldName <> "" Then
      Q1 = Q1 & ",'" & ParaSQL(SValue) & "'"
   End If
   Q1 = Q1 & " )"
   
   Rc = ExecSQL(Db, Q1)

   If Rc >= 0 And SqlnRec = 1 Then
      DoEvents
      Q1 = "SELECT " & UFldName & ", " & IFldName
      
      If SFldName <> "" Then
         Q1 = Q1 & ", " & SFldName
      End If
      
      Q1 = Q1 & " FROM " & TblName
      Q1 = Q1 & " WHERE " & UFldName & " = " & UValue & " " & sWhere

      Set Rs = OpenRs(Db, Q1)
      If Rs.EOF = False Then
         AddUniqueRecord = True
      
         Do Until Rs.EOF
         
            If vFld(Rs(IFldName)) <> RNum Then
               AddUniqueRecord = False
               Exit Do
            End If
            
            If SFldName <> "" Then
               If StrComp(vFld(Rs(SFldName)), SValue) Then
                  AddUniqueRecord = False
                  Exit Do
               End If
            End If
            
            Rs.MoveNext
         Loop
         Call CloseRs(Rs)
         
         If AddUniqueRecord = False Then
            Q1 = " WHERE " & UFldName & " = " & UValue & " AND " & IFldName & "=" & RNum
            If SFldName <> "" Then
               Q1 = Q1 & " And " & SFldName & "='" & ParaSQL(SValue) & "'"
            End If
            
            Rc = DeleteSQL(Db, TblName, Q1)
         End If
         
      Else
         Call CloseRs(Rs)
      End If
      
   End If

End Function

Public Function GetTxConnectInfo(ByVal ConnStr As String, ByVal Key As String) As String
   Dim i As Integer, j As Integer
   
   Key = ";" & Key & "="
   
   i = InStr(1, ConnStr, Key, vbTextCompare)
   If i Then
      i = i + Len(Key)
      j = InStr(i, ConnStr, ";", vbBinaryCompare)

      If j = 0 Then
         GetTxConnectInfo = Trim(Mid(ConnStr, i))
      Else
         GetTxConnectInfo = Trim(Mid(ConnStr, i, j - i))
      End If
   
   Else
      GetTxConnectInfo = ""
   End If

End Function


#If DATACON = DAO_CONN Then 'ADO
Public Function GetConnectInfo(ByVal Db As Database, ByVal Key As String) As String
#Else
Public Function GetConnectInfo(ByVal Db As Connection, ByVal Key As String) As String
#End If
   Dim i As Integer, j As Integer
   Dim ConnStr As String, UKey As String, DSN As String, Buf As String, Buff As String * 51
   Dim Rc As Long, hkProtocol As Long
      
   GetConnectInfo = ""
#If DATACON = DAO_CONN Then 'ADO
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If

   UKey = ";" & UCase(Key) & "="
   i = InStr(1, ConnStr, ";DSN=", vbTextCompare)
   
   If i <> 0 And InStr(1, ConnStr, UKey, vbTextCompare) = 0 Then
      i = i + Len(";DSN=")
      j = InStr(i, ConnStr, ";", vbTextCompare)

      If j = 0 Then
         DSN = Trim(Mid(ConnStr, i))
      Else
         DSN = Trim(Mid(ConnStr, i, j - i))
      End If

      Rc = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\ODBC\ODBC.INI\" & DSN, 0, KEY_READ, hkProtocol)
      If Rc <> 0 Then
         Rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBC.INI\" & DSN, 0, KEY_READ, hkProtocol)
         If Rc <> 0 Then
            Exit Function
         End If
      End If
   
      Buff = ""
      Rc = RegQueryValueExS(hkProtocol, Key, 0, 0, Buff, 50)
      If Rc = 0 Then
         Buf = Left(Buff, StrLen(Buff))
         GetConnectInfo = Buf
      End If
   
      Call RegCloseKey(hkProtocol)
   
      Exit Function
   
   End If
   
   GetConnectInfo = GetTxConnectInfo(ConnStr, Key)
   
End Function
#If DATACON = DAO_CONN Then
Sub DbBeginTrans(Optional ByVal Mrk As String)

   LstTrans(iTrans) = Mrk
   iTrans = iTrans + 1
   
   BeginTrans
   
End Sub

Sub DbCommitTrans(Optional ByVal Mrk As String)

   iTrans = iTrans - 1
   If LstTrans(iTrans) <> Mrk Then
      Beep
      Debug.Print "** COMMIT: No calzan las marcas: [" & LstTrans(iTrans) & "] y [" & Mrk & "]"
   End If
   
   CommitTrans

End Sub

Sub DbRollBack(Optional ByVal Mrk As String)

   iTrans = iTrans - 1
   If LstTrans(iTrans) <> Mrk Then
      Beep
      Debug.Print "** ROLLBACK: No calzan las marcas: [" & LstTrans(iTrans) & "] y [" & Mrk & "]"
   End If
   
   Rollback

End Sub
#End If
#If DATACON = DAO_CONN Then 'ADO
Public Function SqlChar(Db As Database, ByVal Char As String) As String
#Else
Public Function SqlChar(Db As Connection, ByVal Char As String) As String
#End If
   Dim ConnStr As String

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If

   If Left(ConnStr, 5) = "ODBC;" Or InStr(1, ConnStr, "Sql Server", vbTextCompare) <> 0 Then
      Select Case Trim(Char)
         Case "&":   ' Concat
            SqlChar = "+"
         Case "*":   ' Wilcard
            SqlChar = "%"
         Case Else:
            SqlChar = "????"
      End Select
   
   Else
      SqlChar = Char
   End If

End Function
'#If DATACON = DAO_CONN Then 'ADO
'Public Function SqlVal(Db As Database, ByVal FldName As String) As String
'#Else
'Public Function SqlVal(Db As Connection, ByVal FldName As String) As String
'#End If
'   Dim ConnStr As String
'
'#If DATACON = DAO_CONN Then
'   ConnStr = Db.Connect
'#Else
'   ConnStr = Db.ConnectionString
'#End If
'
'   If Left(ConnStr, 5) = "ODBC;" Or InStr(1, ConnStr, "Sql Server", vbTextCompare) > 0 Then
'      SqlVal = "( 1 * " & FldName & ")"
'   Else
'      SqlVal = "val(" & FldName & ")"
'   End If
'
'End Function

Public Function SqlVal(ByVal Expr As String, Optional ByVal DbType As Byte = 0) As String

   If DbType = 0 Then
      DbType = gDbType
   End If
   
   If DbType = SQL_ACCESS Then
      SqlVal = "val(" & Expr & ")"
   Else
      SqlVal = "( 1 * " & Expr & ")"
   End If

End Function

#If DATACON = DAO_CONN Then 'ADO
Public Function GetDbErr(Db As Database, Optional Errno As Long = 0) As String
#Else
Public Function GetDbErr(Db As Connection, Optional Errno As Long = 0) As String
#End If
   Dim Buf As String, Buf1 As String

   Buf = ""
   Errno = 0

#If DATACON = DAO_CONN Then 'DAO
   Dim ErrLoop As dao.Error
   ' Enumerate Errors collection and display properties of
   ' each Error object.
   For Each ErrLoop In dao.Errors
      With ErrLoop
         
         If Len(Buf) = 0 Then
            Errno = .Number
         End If
         
         Buf1 = "Error " & .Number & ", " & .Description & " (Source: " & .Source & ", topic: " & .HelpContext & ")"
         
         If InStr(1, Buf, Buf1, vbTextCompare) = 0 Then
           If Buf <> "" Then
              Buf = Buf & vbLf
           End If
         
           Buf = Buf & Buf1
         End If
         'strError = strError & "  " & .Description & vbCr
         'strError = strError & "  (Source: " & .Source & ")" & vbCr
         'strError = strError & "Press F1 to see topic  " & .HelpContext & vbCr
         'strError = strError & "  in the file " & .HelpFile & "."
      End With
       'MsgBox strError
   Next
#Else
   Dim ErrLoop As ADODB.Error
   For Each ErrLoop In Db.Errors
      With ErrLoop
         
         If Len(Buf) = 0 Then
            Errno = .Number
         End If
         
         Buf1 = "Error " & Hex(.Number) & ", " & .Description & " (Source: " & .Source & ", topic: " & .HelpContext & ")"
         
         If InStr(1, Buf, Buf1, vbTextCompare) = 0 Then
           If Buf <> "" Then
              Buf = Buf & vbLf
           End If
         
           Buf = Buf & Buf1
         End If
         'strError = strError & "  " & .Description & vbCr
         'strError = strError & "  (Source: " & .Source & ")" & vbCr
         'strError = strError & "Press F1 to see topic  " & .HelpContext & vbCr
         'strError = strError & "  in the file " & .HelpFile & "."
      End With
       'MsgBox strError
   Next
   
   'Buf = "Error " & Err & ", " & Error

#End If

   GetDbErr = Buf
   
End Function

' http://msdn2.microsoft.com/en-us/library/ms191516.aspx
' BCP base..tabla format nul -f tabla.bcp -Uuser -Ppass -Sserver

#If DATACON = DAO_CONN Then 'DAO
Public Function GenBCP(Db As Database, ByVal TableName As String, ByVal FName As String, Optional ByVal FldSep As String = "\t", Optional ByVal RecSep As String = "\r\n", Optional fType As Boolean = 0) As Long
   Dim TbDef As TableDef
   Dim Fld As Field
   Dim F As Integer, l As String
   Dim Fd As Long
   Dim Sep As String, Buf As String, DType As String
      
   GenBCP = 0
   On Error Resume Next
   
   Set TbDef = Db.TableDefs(TableName)
   If ERR Then
      GenBCP = ERR.Number
      Exit Function
   End If
   
   Fd = FreeFile()
   Open FName For Output As #Fd
   If ERR Then
      GenBCP = ERR.Number
      Exit Function
   End If
   
   Print #Fd, "6.0"
   Print #Fd, "" & TbDef.Fields.Count
   Sep = """" & FldSep & """"
   
   For F = 0 To TbDef.Fields.Count - 1
   
      If F = TbDef.Fields.Count - 1 Then
         Sep = """" & RecSep & """"
      End If
   
      Set Fld = TbDef.Fields(F)
      
      Select Case Fld.Type
         Case dbText:
            l = Fld.Size
            DType = "SQLCHAR"
            
         Case dbMemo:
            l = Fld.Size
            DType = "SQLCHAR"
            
         Case dbDate:
            l = 26
            DType = "SQLDATETIME"
         
         Case dbByte:
            l = 6
            DType = "SQLTINYINT"
         
         Case dbInteger:
            l = 6
            DType = "SQLSMALLINT"
         
         Case dbLong:
            l = 12
            DType = "SQLINT"
            
         Case dbCurrency:
            l = 24
            DType = "SQLMONEY"
            
         Case dbSingle:
            l = 25
            DType = "SQLFLT4"
         
         Case dbDouble:
            l = 25
            DType = "SQLFLT8"
         
         Case Else:
            l = "??? " & Fld.Type
         
      End Select
      
      If fType = False Then
         DType = "SQLCHAR"
      End If
      
      Buf = (F + 1) & vbTab & DType & vbTab & "0" & vbTab & l & vbTab & Sep & vbTab & (F + 1) & vbTab & Fld.Name
      
      Print #Fd, Buf
      If ERR Then
         Close #Fd
         GenBCP = ERR.Number
         Exit Function
      End If
      
   Next F
   
   Close #Fd

End Function
#End If

' Genera el comando para ejecutar el BCP.exe
'    Bcp = "BCP Indura..Obras in " & w.apppath & "\Datos\Obras.txt -m5 -f" & w.apppath & "\Obras.bcp -e" & w.apppath & "\Log\ObrasErr.txt -b2000 -F2 -U" & gUser.Userid & " -P" & gUser.Passw & " -S" & GetConnectInfo(DbMain, "SERVER")
#If DATACON = DAO_CONN Then 'ADO
Public Function GenBcpCmd2(Db As Database, Bcp As BCP_t, ByVal TableName As String, ByVal FnDatos As String, ByVal FnBcp As String, ByVal FnErr As String, ByVal FirstRow As Integer, ByVal ChkFiles As Boolean, Optional ByVal Pag As String = "RAW") As Boolean
#Else
Public Function GenBcpCmd2(Db As Connection, Bcp As BCP_t, ByVal TableName As String, ByVal FnDatos As String, ByVal FnBcp As String, ByVal FnErr As String, ByVal FirstRow As Integer, ByVal ChkFiles As Boolean, Optional ByVal Pag As String = "RAW") As Boolean
#End If
   Dim Msg As String
   Dim i As Integer
   
   Bcp.Cmd = ""
   Bcp.FInfo = ""
   Bcp.TblErr = ""
   
   If ChkFiles Then
   
      On Error Resume Next
      GenBcpCmd2 = False
      
      If Dir(FnDatos) = "" Then
         Msg = "No existe el archivo de datos. " & FnDatos
         Call AddLog("GenBcp: " & Msg)
         MsgBox1 Msg, vbExclamation
         Exit Function
      End If
   
      If FileLen(FnDatos) <= 0 Then
         Msg = "El archivo de datos está vació. " & FnDatos
         Call AddLog("GenBcp: " & Msg)
         MsgBox1 Msg, vbExclamation
         Exit Function
      End If
   
      Bcp.FInfo = FnDatos & ": " & Format(FileDateTime(FnDatos), "yyyy-mm-dd hh:nn:ss") & " - " & Format(FileLen(FnDatos), NUMFMT) & " bytes"
   
      If Dir(FnBcp) = "" Then
         Msg = "No existe el archivo de formato " & FnBcp
         Call AddLog("GenBcp: " & Msg)
         MsgBox1 Msg, vbExclamation
         Exit Function
      End If
   
      For i = Len(FnBcp) To 1 Step -1
         If Mid(FnBcp, i, 1) = "\" Then
            i = i + 1
            Exit For
         End If
      Next i
   
      Bcp.FInfo = Bcp.FInfo & "; " & Mid(FnBcp, i) & ": " & Format(FileDateTime(FnBcp), "yyyy-mm-dd hh:nn:ss") & " - " & Format(FileLen(FnBcp), NUMFMT) & " bytes"
   
   End If

   ' -C página de código: ACP (Ansi/Windows), OEM (default), RAW

   Bcp.TblErr = FnErr
   Bcp.Cmd = "BCP " & GetConnectInfo(Db, "DATABASE") & ".." & TableName & " in " & FnDatos & " -m5 -f" & FnBcp & " -e" & FnErr & " -b2000 -C" & Pag & " -F" & FirstRow & " -U" & GetConnectInfo(Db, "UID") & " -P" & GetConnectInfo(Db, "PWD") & " -S" & GetConnectInfo(Db, "SERVER")
   
   GenBcpCmd2 = (Len(Bcp.Cmd) > 0)
   
End Function
' Genera el comando para ejecutar el BCP.exe
'    Bcp = "BCP Indura..Obras in " & w.apppath & "\Datos\Obras.txt -m5 -f" & w.apppath & "\Obras.bcp -e" & w.apppath & "\Log\ObrasErr.txt -b2000 -F2 -U" & gUser.Userid & " -P" & gUser.Passw & " -S" & GetConnectInfo(DbMain, "SERVER")
#If DATACON = DAO_CONN Then 'ADO
Public Function GenBcpCmd(Db As Database, ByVal TableName As String, ByVal FnDatos As String, ByVal FnBcp As String, ByVal FnErr As String, ByVal FirstRow As Integer, ByVal ChkFiles As Boolean, FInfo As String, Optional ByVal Pag As String = "RAW") As String
#Else
Public Function GenBcpCmd(Db As Connection, ByVal TableName As String, ByVal FnDatos As String, ByVal FnBcp As String, ByVal FnErr As String, ByVal FirstRow As Integer, ByVal ChkFiles As Boolean, FInfo As String, Optional ByVal Pag As String = "RAW") As String
#End If
   Dim Msg As String
   Dim i As Integer
   
   FInfo = ""

   If ChkFiles Then
   
      On Error Resume Next
      GenBcpCmd = ""
      
      If Dir(FnDatos) = "" Then
         Msg = "No existe el archivo de datos. " & FnDatos
         Call AddLog("GenBcp: " & Msg)
         MsgBox1 Msg, vbExclamation
         Exit Function
      End If
   
      If FileLen(FnDatos) <= 0 Then
         Msg = "El archivo de datos está vació. " & FnDatos
         Call AddLog("GenBcp: " & Msg)
         MsgBox1 Msg, vbExclamation
         Exit Function
      End If
   
      FInfo = FnDatos & ": " & Format(FileDateTime(FnDatos), "yyyy-mm-dd hh:nn:ss") & " - " & Format(FileLen(FnDatos), NUMFMT) & " bytes"
   
      If Dir(FnBcp) = "" Then
         Msg = "No existe el archivo de formato " & FnBcp
         Call AddLog("GenBcp: " & Msg)
         MsgBox1 Msg, vbExclamation
         Exit Function
      End If
   
      For i = Len(FnBcp) To 1 Step -1
         If Mid(FnBcp, i, 1) = "\" Then
            i = i + 1
            Exit For
         End If
      Next i
   
      FInfo = FInfo & "; " & Mid(FnBcp, i) & ": " & Format(FileDateTime(FnBcp), "yyyy-mm-dd hh:nn:ss") & " - " & Format(FileLen(FnBcp), NUMFMT) & " bytes"
   
   End If

   ' -C página de código: ACP (Ansi/Windows), OEM (default), RAW


   GenBcpCmd = "BCP " & GetConnectInfo(Db, "DATABASE") & ".." & TableName & " in " & FnDatos & " -m5 -f" & FnBcp & " -e" & FnErr & " -b2000 -C" & Pag & " -F" & FirstRow & " -U" & GetConnectInfo(Db, "UID") & " -P" & GetConnectInfo(Db, "PWD") & " -S" & GetConnectInfo(Db, "SERVER")

End Function
#If DATACON = DAO_CONN Then 'ADO
' Creada 29-AGO-2004
Public Sub SetDbSecurity_new(ByVal DbPath As String, ByVal Passw As String, ByVal CfgFile As String, ByVal SegCfg As String, oConnStr As String, Optional LngPassw As String = "")
   Dim Db As Database, i As Integer
   Dim Cfg As String, ConnStr1(2, 2) As String
   Dim bSeg As Boolean
   
   On Error Resume Next
      
   If LngPassw = "" Then
      LngPassw = Passw
   Else
      ConnStr1(2, 0) = LngPassw
      ConnStr1(2, 1) = ";PWD=" & LngPassw & ";"
      ConnStr1(2, 2) = ";PWD=" & LngPassw & ";"
   End If
   
   ConnStr1(0, 0) = ""
   ConnStr1(0, 1) = ""
   ConnStr1(0, 2) = ";PWD=" & LngPassw & ";"

   ConnStr1(1, 0) = Passw
   ConnStr1(1, 1) = ";PWD=" & Passw & ";"
   ConnStr1(1, 2) = ";PWD=" & LngPassw & ";"
   
   bSeg = (GetIniString(CfgFile, "Config", "Secur", "") <> SegCfg)
   
   For i = 0 To 2
      Set Db = OpenDatabase(DbPath, True, False, ConnStr1(i, 1))
      
      If Not Db Is Nothing Then
         ERR.Clear
      
         If bSeg Then  ' le pone clave
            Db.NewPassword ConnStr1(i, 0), LngPassw
            oConnStr = ConnStr1(i, 2)
         Else
            Db.NewPassword ConnStr1(i, 0), ""
            oConnStr = ""
         End If
    
         Call CloseDb(Db)
         Exit Sub
      End If
   Next i
      
End Sub
' Creada 03-JUN-2002
Public Sub SetDbSecurity(ByVal DbPath As String, ByVal Passw As String, ByVal CfgFile As String, ByVal SegCfg As String, ConnStr As String)
   Dim Db As Database
   Dim Cfg As String, ConnStr1 As String
   Dim Seg As Boolean
   
   On Error Resume Next
   
   Cfg = GetIniString(CfgFile, "Config", "Secur", "")
   'MsgBox "ruta : " & CfgFile & "clave CFG: " & Cfg & " CLAVE SegCfg: " & SegCfg
   If Cfg <> SegCfg Then
      Seg = True
      ConnStr1 = ""
      ConnStr = ";PWD=" & Passw & ";"
   Else
      Seg = False
      ConnStr1 = ";PWD=" & Passw & ";"
      ConnStr = ""
   End If

   ERR.Clear

   ' Probamos a abrir la base con lo contrario de la seguridad esperada
   Set Db = OpenDatabase(DbPath, True, False, ConnStr1)
   If Not Db Is Nothing Then ' si pudo abrir, entonces hay que cambiar
      
      If Seg Then
         Db.NewPassword "", Passw   ' le pone clave
      Else
         Db.NewPassword Passw, ""   ' le quita la clave
      End If
      
      If ERR Then
         Call AddLog("SetDbSecurity: Seg=" & Seg & ", " & DbPath & ", Error " & ERR & ", " & ERR.Description)
      End If
      
      Call CloseDb(Db)
   Else
      Call AddLog("SetDbSecurity: No se pudo quitar/poner clave a la base de datos, " & DbPath & ", Error " & ERR & ", " & ERR.Description)
      Debug.Print "*** No se pudo quitar/poner clave a la base de datos."
   End If
   
End Sub

#End If

'Public Sub SetDbSecurityCambio(ByVal DbPath As String, ByVal Passw As String, ByVal CfgFile As String, ByVal SegCfg As String, ConnStr As String, tipo As String, rut As String)
'   Dim Db As Database
'   Dim Cfg As String, ConnStr1 As String
'   Dim Seg As Boolean
'   Dim passOld As String
'
'   If tipo = "EMP" Then
'        passOld = PASSW_PREFIX & rut
'   Else
'        passOld = PASSW_LEXCONT
'   End If
'
'
'   On Error Resume Next
'
'   Cfg = GetIniString(CfgFile, "Config", "Secur", "")
'   'MsgBox "ruta : " & CfgFile & "clave CFG: " & Cfg & " CLAVE SegCfg: " & SegCfg
'   If Cfg <> SegCfg Then
'      Seg = True
'      ConnStr1 = ""
'      ConnStr = ";PWD=" & Passw & ";"
'   Else
'      Seg = False
'      ConnStr1 = ";PWD=" & Passw & ";"
'      ConnStr = ""
'   End If
'
'   ERR.Clear
'
'   ' Probamos a abrir la base con lo contrario de la seguridad esperada
'   Set Db = OpenDatabase(DbPath, True, False, ConnStr1)
'   If Not Db Is Nothing Then ' si pudo abrir, entonces hay que cambiar
'
'      If Seg Then
'         'Db.NewPassword "", passOld   ' le pone clave
'         'If Passw = PASSW_PREFIX_NEW & rut Or Passw = PASSW_LEXCONT_NEW Or Passw = PASSW_LEXCONT_NEW2 Then
'            Db.NewPassword Passw, passOld
'            Call AddLog("pass que tenia " & Passw & "  La vieja que se le puso " & passOld)
'         'End If
'      Else
'         'Db.NewPassword passOld, ""   ' le quita la clave
'         'If Passw = PASSW_PREFIX_NEW & rut Or Passw = PASSW_LEXCONT_NEW Or Passw = PASSW_LEXCONT_NEW2 Then
'            Db.NewPassword Passw, passOld
'            Call AddLog("pass que tenia " & Passw & "  La vieja que se le puso " & passOld)
'         'End If
'      End If
'
'
'      If ERR Then
'         Call AddLog("SetDbSecurity: Seg=" & Seg & ", " & DbPath & ", Error " & ERR & ", " & ERR.Description)
'      End If
'
'      Call CloseDb(Db)
'   Else
'      Call AddLog("SetDbSecurity: No se pudo quitar/poner clave a la base de datos, " & DbPath & ", Error " & ERR & ", " & ERR.Description)
'      Debug.Print "*** No se pudo quitar/poner clave a la base de datos."
'   End If
'
'End Sub
'#End If
Public Function FldIsString(Fld As Field) As Boolean
        Dim bType As Boolean

   #If DATACON = DAO_CONN Then
      bType = (Fld.Type = dbText Or Fld.Type = dbMemo Or Fld.Type = dbChar)
   #Else
      bType = (Fld.Type = adChar Or Fld.Type = adVarChar Or Fld.Type = adLongVarChar Or Fld.Type = adLongVarWChar Or Fld.Type = adVarWChar Or Fld.Type = adWChar)
   #End If

   FldIsString = bType

End Function

#If DATACON = DAO_CONN Then
Public Function CreateQry(Db As Database, ByVal QName As String, ByVal Qry As String) As Integer
   Dim QryDef As QueryDef
   Dim Q1 As String
   Dim Rc As Long

   On Error Resume Next

'   Q1 = "DROP TABLE " & QName
'   Rc = ExecSQL(Db, Q1)

   Call Db.QueryDefs.Delete(QName)

   ' Error 3265 - Item not found in this collection
   ' Error 3011 - Access database engine could not find object

   If ERR = 0 Or ERR = 3265 Or ERR = 3011 Then
      Set QryDef = Db.CreateQueryDef
      QryDef.Name = QName
      QryDef.SQL = Qry
      Db.QueryDefs.Append QryDef
      Db.QueryDefs.Refresh
   
      CreateQry = True
   Else
      CreateQry = False
   End If
   
End Function
#Else
Public Function CreateQry(Db As Connection, ByVal QName As String, ByVal Qry As String) As Integer
   Dim Q1 As String
   Dim Rc As Long

   On Error Resume Next

   Q1 = "DROP VIEW " & QName
   Rc = ExecSQL(Db, Q1)
  
   Q1 = "CREATE VIEW " & QName & " AS " & Qry
   Rc = ExecSQL(Db, Q1)
  
   If ERR = 0 Then
      CreateQry = True
   Else
      CreateQry = False
   End If
   
End Function

#End If

#If DATACON = DAO_CONN Then
' Duplica la tabla pero sin datos
Public Function DuplicTable(Db As Database, ByVal TblOrig As String, ByVal TblDest As String, Optional ByVal bIdx As Boolean = 1) As Long
   Dim Q1 As String, Rc As Long, i As Integer
   Dim TDef1 As TableDef, Idx1 As Index
   Dim TDef2 As TableDef, Idx2 As Index

   Q1 = "DROP TABLE " & TblDest
   Rc = ExecSQL(Db, Q1)

   Q1 = "SELECT * INTO " & TblDest & " FROM " & TblOrig & " WHERE 1=0"
   Rc = ExecSQL(Db, Q1)

   If bIdx Then
   
      Set TDef1 = Db.TableDefs(TblOrig)
      Set TDef2 = Db.TableDefs(TblDest)
      
      For Each Idx1 In TDef1.Indexes
         
         Set Idx2 = New Index
         
         Idx2.Fields = Idx1.Fields
         
         On Error Resume Next
         For i = 0 To Idx1.Properties.Count - 1
            Idx2.Properties(i) = Idx1.Properties(i)
         Next i
         On Error GoTo 0
         
         TDef2.Indexes.Append Idx2

      Next Idx1

   End If

   DuplicTable = Rc
   
End Function
#End If

' Calcula un checksum del resultado del query
#If DATACON = DAO_CONN Then
Public Function ChkSumQry(Db As Database, ByVal Qry As String) As Long
#Else
Public Function ChkSumQry(Db As Connection, ByVal Qry As String) As Long
#End If
   Dim Rs As Recordset
   Dim i As Integer, Chk As Long, n As Long

   Chk = 1000
   n = 1
   
   Set Rs = OpenRs(Db, Qry)
   
   Do Until Rs.EOF
   
      For i = 0 To Rs.Fields.Count - 1
         Chk = Chk + Hash(vFld(Rs(i))) * (i + 1)
      Next i

      n = n + 1
      Chk = Chk Mod 77777777

      Rs.MoveNext
   Loop
   
   Chk = (Chk + (n * Rs.Fields.Count)) Mod 77777777
   
   Call CloseRs(Rs)

   ChkSumQry = Chk

End Function

' Debe existir la tabla LockAction con los siguientes campos
'     idLock     AutoNumber
'     Fecha      Fecha/hora
'     PcName     char(30)
'     hInstance  long
'     idAction   integer
'     idItem     long
'     *** Indice único por idAction, idItem (Primary Key) ***
' Si pudo bloquear retorna "", sino, retorna el PC que bloquea
#If DATACON = DAO_CONN Then
Public Function LockAction(Db As Database, ByVal idAction As Integer, Optional ByVal IdItem As Long = 0) As Boolean
#Else
Public Function LockAction(Db As Connection, ByVal idAction As Integer, Optional ByVal IdItem As Long = 0) As Boolean
#End If
   Dim Q1 As String, PcName As String
   Dim Rs As Recordset
   Dim Rc As Long, Dt As Double
   
   PcName = "'" & GetComputerName() & "'"
   Dt = Now - TimeSerial(4, 0, 0) ' 22 feb 2010: si tiene más de 4 horas, se elimina primero
   
   Q1 = " WHERE idAction=" & idAction
   Q1 = Q1 & " AND idItem=" & IdItem
   Q1 = Q1 & " AND Fecha <" & str(Dt)
   
   Rc = DeleteSQL(Db, "LockAction", Q1)
     
'   Q1 = "DELETE l.* FROM LockAction l"
'   Q1 = Q1 & " WHERE idAction=" & idAction
'   Q1 = Q1 & " AND idItem=" & idItem
'   Q1 = Q1 & " AND Fecha <" & Dt
'   Rc = ExecSQL(Db, Q1, False) ' 22 feb 2010: si tiene más de 4 horas, se elimina primero
            
   Q1 = "INSERT INTO LockAction (Fecha, PcName, hInstance, idAction, idItem )"
   Q1 = Q1 & " VALUES( " & SqlNow(Db) & "," & PcName & "," & App.hInstance & "," & idAction & "," & IdItem & ")"
   Rc = ExecSQL(Db, Q1, False)

   Q1 = "SELECT idLock, PcName, hInstance FROM LockAction"
   Q1 = Q1 & " WHERE idAction=" & idAction
   Q1 = Q1 & " AND idItem=" & IdItem
   Q1 = Q1 & " AND PcName=" & PcName
   Q1 = Q1 & " AND hInstance=" & App.hInstance
   
   Set Rs = OpenRs(Db, Q1)

   If Rs.EOF = True Then
      LockAction = False
   Else
      LockAction = True
   End If

   Call CloseRs(Rs)

End Function
' Para desbloquear alguna accion si un PC se cayo, hay que llamar esta función con bUseInstance = 0
#If DATACON = DAO_CONN Then
Public Sub UnLockAction(Db As Database, ByVal idAction As Integer, Optional ByVal IdItem As Long = 0, Optional ByVal bUseInstance As Boolean = 1, Optional ByVal bUsePC As Boolean = 1, Optional ByVal bUseItem As Boolean = 1)
#Else
Public Sub UnLockAction(Db As Connection, ByVal idAction As Integer, Optional ByVal IdItem As Long = 0, Optional ByVal bUseInstance As Boolean = 1, Optional ByVal bUsePC As Boolean = 1, Optional ByVal bUseItem As Boolean = 1)
#End If
   Dim PcName As String, Q1 As String
   Dim Rc As Long
   
   Q1 = " WHERE idAction=" & idAction
   
   If bUseItem Then
      Q1 = Q1 & " AND idItem=" & IdItem
   End If
   
   If bUsePC Then
      PcName = "'" & GetComputerName() & "'"
      Q1 = Q1 & " AND PcName=" & PcName
   End If
   
   If bUseInstance Then
      Q1 = Q1 & " AND hInstance=" & App.hInstance
   End If
   
   Rc = DeleteSQL(Db, "LockAction", Q1)

End Sub
' Si está bloqueado retorna el nombre del PC, si no retorna ""
#If DATACON = DAO_CONN Then
Public Function IsLockedAction(Db As Database, ByVal idAction As Integer, Optional ByVal IdItem As Long = 0) As String
#Else
Public Function IsLockedAction(Db As Connection, ByVal idAction As Integer, Optional ByVal IdItem As Long = 0) As String
#End If

   Dim Q1 As String
   Dim Rs As Recordset
      
   Q1 = "SELECT idLock, PcName FROM LockAction"
   Q1 = Q1 & " WHERE idAction=" & idAction
   Q1 = Q1 & " AND idItem=" & IdItem
   Set Rs = OpenRs(Db, Q1)

   If Rs.EOF = True Then
      IsLockedAction = ""
   Else
      IsLockedAction = vFld(Rs("PcName"))
   End If

   Call CloseRs(Rs)

End Function
Public Function FldSize(Fld As Field, Optional ByVal DefSize As Integer = 0) As Integer

#If DATACON = DAO_CONN Then
   FldSize = Fld.Size
#Else
   FldSize = Fld.DefinedSize
#End If

   If FldSize <= 0 Then
      FldSize = DefSize
   End If

End Function
#If DATACON = DAO_CONN Then
Public Function CreateTableLockAction(Db As Database, Optional ByVal bMsg As Boolean = 1) As Boolean
#Else
Public Function CreateTableLockAction(Db As Connection, Optional ByVal bMsg As Boolean = 1) As Boolean
#End If
   Dim Q1 As String, ConnStr As String
   Dim Rc As Long
#If DATACON = DAO_CONN Then
   Dim Tbl As TableDef, Fld As Field
#End If
   
   CreateTableLockAction = True
   
   If SqlType(Db, ConnStr) = SQL_SERVER Then
      Q1 = "CREATE TABLE LockAction (idLock int IDENTITY (1, 1) NOT NULL"
      Q1 = Q1 & ", Fecha datetime NULL, PcName char(30) NULL"
      Q1 = Q1 & ", hInstance int NULL, idAction int NULL, idItem int NULL"
      Q1 = Q1 & " ) ON [PRIMARY]"
      Rc = ExecSQL(Db, Q1, bMsg)
   Else
#If DATACON = DAO_CONN Then
      Q1 = "CREATE TABLE LockAction ( "
      Q1 = Q1 & " Fecha datetime NULL, PcName char(30) NULL"
      Q1 = Q1 & ", hInstance int NULL, idAction int NULL, idItem int NULL"
      Q1 = Q1 & " )"
      Rc = ExecSQL(Db, Q1, bMsg)
   
      Db.TableDefs.Refresh
   
      Set Tbl = Db.TableDefs("LockAction")
   
      ERR.Clear
      Set Fld = Tbl.CreateField("idLock", dbLong)
      Fld.Attributes = dbAutoIncrField
      Fld.OrdinalPosition = 0
   
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 And bMsg <> 0 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LockAction.idLock", vbExclamation
      End If
#End If
   End If

   Q1 = "CREATE UNIQUE INDEX idLock ON LockAction (idLock )"
   Rc = ExecSQL(Db, Q1, bMsg)

   Q1 = "CREATE UNIQUE INDEX ActionItem ON LockAction (idAction, idItem ) WITH PRIMARY"
   Rc = ExecSQL(Db, Q1, bMsg)

End Function



'Sub test(Optional ByVal un As Boolean = 0)
'   Dim i As Integer
'   Dim Ac1 As Integer, Ac2 As Integer
'   Const n = 5
'
'   If un Then
'      For i = 1 To n
'         Debug.Print "* DEL Acc=" & i
'         Call UnLockAction(DbMain, i, , False)
'      Next i
'   End If
'
'   For i = 1 To 50
'      Ac1 = Int(1 + Rnd * n)
'
'      Debug.Print "ADD Acc=" & Ac1 & ", ID=" & LockAction(DbMain, Ac1)
'
'      Ac2 = Int(1 + Rnd * 5)
'
'      Debug.Print "DEL Acc=" & Ac2
'      Call UnLockAction(DbMain, Ac2)
'
'      Debug.Print "IS Acc=" & Ac1 & " - " & IsLockedAction(DbMain, Ac1)
'
'   Next i
'
'End Sub

' para que ponga todos los campos
#If DATACON = DAO_CONN Then
Public Function GenInsSel(Db As Database, ByVal TbFrom As String, ByVal TbTo As String)
   Dim tBF As TableDef, TbT As TableDef
   Dim Q1 As String, F As Integer, T As Integer
   Dim b As Boolean
   
   Set tBF = Db.TableDefs(TbFrom)
   Set TbT = Db.TableDefs(TbTo)

   Q1 = ""
   For T = 0 To TbT.Fields.Count - 1
   
      b = 0
      For F = 0 To tBF.Fields.Count - 1
         If StrComp(TbT.Fields(T).Name, tBF.Fields(F).Name, vbTextCompare) = 0 Then
            b = 1
            Exit For
         End If
      Next F
      
      If b Then
         Q1 = Q1 & ", " & TbT.Fields(T).Name
      Else
         Q1 = Q1 & ", NULL as " & TbT.Fields(T).Name
      End If
      
   Next T
   
   Q1 = "INSERT INTO " & TbTo & " SELECT " & Mid(Q1, 2) & " FROM " & TbFrom

   GenInsSel = Q1

End Function
#End If

Public Function DbGenTmpName2(ByVal DbType As Byte, ByVal Extra As String, Optional bTime As Boolean = 0) As String
   Dim Buf As String
   
   Buf = DbGenTmpName(Extra, bTime)
   
   If DbType = SQL_SERVER Then
      DbGenTmpName2 = "#" & Buf  ' En SQL Server las temporales parten con #
   Else
      DbGenTmpName2 = Buf
   End If

End Function
Public Function DbGenTmpName(ByVal Extra As String, Optional bTime As Boolean = 0) As String
   Dim Buf As String, i As Integer

   Buf = W.PcName
   Buf = ReplaceStr(Buf, " ", "")
   Buf = ReplaceStr(Buf, "'", "")
   Buf = ReplaceStr(Buf, """", "")
   Buf = ReplaceStr(Buf, "-", "_")
   Buf = ReplaceStr(Buf, ".", "_")
   Buf = ReplaceStr(Buf, ",", "_")
   
   Buf = "tmp_" & Buf
   
   Extra = Trim(Extra)
   If Extra <> "" Then
      Extra = ReplaceStr(Extra, " ", "")
      Extra = ReplaceStr(Extra, "'", "")
      Extra = ReplaceStr(Extra, """", "")
      Extra = ReplaceStr(Extra, "-", "_")
      Extra = ReplaceStr(Extra, ".", "_")
      Extra = ReplaceStr(Extra, ",", "_")
      Buf = Buf & "_" & Extra
   End If
   
   If bTime Then
      Buf = Buf & "_" & Format(Now, "dhnnss")
   End If
   
   Buf = Buf & "_" & (App.hInstance Mod 10000) ' 5 abr 2019: Para que no se repita en el mismo equipo
   
   DbGenTmpName = Buf

End Function

Public Function FmtFld(Fld As Field, Optional ByVal Fmt As String = "", Optional ByVal bPositive As Boolean = 0, Optional ByVal Fact As Integer = 1) As String
   Dim v As Variant

   If IsNull(Fld) Then
      FmtFld = ""
   Else
      v = vFld(Fld)
      If bPositive Then
         If Val(v) <= 0 Then
            FmtFld = ""
            Exit Function
         End If
      End If
   
      If Fmt <> "" Then
         FmtFld = Format(v * Fact, Fmt)
      Else
         FmtFld = Format(v * Fact)
      End If
   End If

End Function
' 9 sep 2020: por si viene un campo null
Public Function FmtFldNum(Fld As Field, Optional ByVal Fmt As String = "0") As String
   Dim v As Double

   If IsNull(Fld) Then
      v = 0
   Else
      v = Val(vFld(Fld))
   End If

   FmtFldNum = Format(v, Fmt)

End Function
' Para Horas, si es mayor que cero asume que viene como hhnn sino asume hora en double
Public Function FmtHrFld(Fld As Field, Optional ByVal Fmt As String = "hh:nn") As String
   Dim H As Double

   If IsNull(Fld) Then
      FmtHrFld = ""
   Else
      H = vFld(Fld)
      
      If H < 0 Then
         FmtHrFld = ""
         Exit Function
      End If
      
      If H >= 1 Then
         FmtHrFld = Format(vFmtHour2(H), Fmt)
      Else
         FmtHrFld = Format(H, Fmt)
      End If
      
   End If

End Function

Public Function FmtFldSiNo(Fld As Field) As String
   Dim v As Variant

   If IsNull(Fld) Then
      FmtFldSiNo = ""
   Else
      FmtFldSiNo = FmtSiNo(vFld(Fld))
   End If

End Function

#If DATACON = DAO_CONN Then
Public Function SqlType(Db As Database, Optional ByVal ConnStr As String = "") As Byte
#Else
Public Function SqlType(Db As Connection, Optional ByVal ConnStr As String = "") As Byte
#End If

#If DATACON = DAO_CONN Then
   ConnStr = Db.Connect
#Else
   ConnStr = Db.ConnectionString
#End If

'   30 oct 2018: se quita la version en Provider=SQLNCLI y Provider=MSDASQL y Provider=SQLOLEDB
'   If Left(ConnStr, 5) = "ODBC;" Or InStr(1, ConnStr, "Sql Server", vbTextCompare) <> 0 Or InStr(1, ConnStr, "MSDASQL.1", vbTextCompare) <> 0 Or InStr(1, ConnStr, "Provider=SQLNCLI10.1", vbTextCompare) <> 0 Or InStr(1, ConnStr, "SQLOLEDB.1", vbTextCompare) <> 0 Then
   If InStr(1, ConnStr, "Sql Server", vbTextCompare) <> 0 Or InStr(1, ConnStr, "Provider=MSDASQL", vbTextCompare) <> 0 Or InStr(1, ConnStr, "Provider=SQLNCLI", vbTextCompare) <> 0 Or InStr(1, ConnStr, "Provider=SQLOLEDB", vbTextCompare) <> 0 Then
      SqlType = SQL_SERVER
   ElseIf InStr(1, ConnStr, "MySQL", vbTextCompare) > 0 Then
      SqlType = SQL_MYSQL
   Else
      SqlType = SQL_ACCESS
   End If

End Function

Function SqlCase(ByVal DbType As Integer, ByVal Cond As String, ByVal StmTrue As String, ByVal StmFalse As String) As String

   If DbType = SQL_ACCESS Then ' desde SQL Server 2012 tambien se puede usar IIF, ver https://www.mytecbits.com/microsoft/sql-server/iif-vs-case
      SqlCase = " IIF( " & Cond & " , " & StmTrue & " , " & StmFalse & " ) "
   Else
      SqlCase = " CASE WHEN " & Cond & " THEN " & StmTrue & " ELSE " & StmFalse & " END "
   End If

End Function

Public Function SqlNum(ByVal Num As String) As String

   If Trim(Num) = "" Then
      SqlNum = " NULL "
   Else
      SqlNum = Str0(vFmt(Num))
   End If

End Function
Public Function SqlSiNo(ByVal SiNo As String) As String

   If Trim(SiNo) = "" Then
      SqlSiNo = " NULL "
   Else
      SqlSiNo = ValSiNo(SiNo)
   End If

End Function
Public Function SqliNum(ByVal Num As String) As String

   If Trim(Num) = "" Then
      SqliNum = " NULL "
   Else
      SqliNum = Val(Num)
   End If

End Function

' Para poner números en una sentencia SQL
Function vFmtNull(ByVal Buf As String, Optional ByVal Div As Integer = 1) As String
   Dim Num As Double

   If Trim(Buf) = "" Then
      vFmtNull = " NULL "
   
   Else
      Num = 0
      On Error Resume Next
      Num = Format(Buf) / Div
      vFmtNull = Str0(Num)
   End If
   
End Function
' Para poner horas en una sentencia SQL
Function vFmtHrNull(ByVal Buf As String) As String
   Dim Num As Double

   If Trim(Buf) = "" Then
      vFmtHrNull = " NULL "
   
   Else
      Num = 0
      On Error Resume Next
      Num = vFmtHour(Buf)
      vFmtHrNull = Format(Num, "hhnn")
   End If
   
End Function


Public Sub SetFldValue(Tx As TextBox, Fld As Field)

#If DATACON = DAO_CONN Then
   If Fld.Type = dbText Then
      Tx.MaxLength = Fld.Size
   End If
#Else
   Tx.MaxLength = Fld.DefinedSize
#End If
   
   If IsNull(Fld) Then
      Tx.Text = ""
   Else
      Tx.Text = vFld(Fld)
   End If
   
End Sub
#If DATACON = ADO_CONN Then

Public Function DbLoadFile(Db As Connection, ByVal Qry As String, ByVal Filename As String) As Integer
   Dim Rs As ADODB.Recordset
   
   On Error Resume Next
   DbLoadFile = 0
   
   Set Rs = New ADODB.Recordset
   Rs.Open Qry, Db, adOpenKeyset, adLockOptimistic
   If Rs Is Nothing Then
      DbLoadFile = ERR.Number
      Call AddLog("DbLoadFile: Err=" & ERR.Number & ", " & ERR.Description)
      Exit Function
   End If
   
   If Rs.EOF Then
      DbLoadFile = 3179  ' Encountered unexpected end of file.

      Call AddLog("DbLoadFile: no se encontró el registro. Err=3179")

      If W.InDesign Then
         MsgBox1 "La consulta no retorna ningún registro." & Qry, vbExclamation
      End If
   
   Else
   
      Dim mstream As ADODB.Stream
      Set mstream = New ADODB.Stream
      mstream.Type = adTypeBinary
      mstream.Open
      mstream.LoadFromFile Filename
      Rs.Fields(0).Value = mstream.Read
      Rs.Update
      If ERR Then
         DbLoadFile = ERR.Number
         Call AddLog("DbLoadFile: Update, Err=" & ERR.Number & ", " & ERR.Description)
      End If
      
      Rs.MoveNext
      
      If Rs.EOF = False And W.InDesign Then
         MsgBox1 "La consulta retorna más de un registro, se asignó al primero." & Qry, vbExclamation
      End If
   End If
   
   Rs.Close
   Set Rs = Nothing

End Function
' Debe venir el campo en la primera posición
Public Function DbSaveFile(Db As Connection, ByVal Qry As String, ByVal Filename As String) As Integer
   Dim Rs As ADODB.Recordset, FName As String
   
   On Error Resume Next
   DbSaveFile = 0
   
   Set Rs = New ADODB.Recordset
   Rs.Open Qry, Db, adOpenKeyset, adLockOptimistic
   
   If Rs Is Nothing Then
      DbSaveFile = ERR.Number
      Exit Function
   End If
   
   If Rs.EOF Then
      DbSaveFile = 3179  ' Encountered unexpected end of file.

      If W.InDesign Then
         MsgBox1 "La consulta no retorna ningún registro." & Qry, vbExclamation
      End If
   
   Else
      Dim mstream As ADODB.Stream
      Set mstream = New ADODB.Stream
      mstream.Type = adTypeBinary
      mstream.Open
      mstream.Write Rs.Fields(0).Value
           
      Debug.Print mstream.Size
      mstream.SaveToFile Filename, adSaveCreateOverWrite
      If ERR Then
         DbSaveFile = ERR.Number
      End If
      
      Rs.MoveNext
      If Rs.EOF = False And W.InDesign Then
         MsgBox1 "La consulta retorna más de un registro." & Qry, vbExclamation
      End If
   End If
   
   Rs.Close
   Set Rs = Nothing

End Function


#End If
#If DATACON = DAO_CONN Then 'ADO
' Para soportar tablas linkeadas. Ojo, no DbOpenTable porque hay una constante con ese nombre
Public Function DbOpenTable1(Db As Database, TableName As String) As Recordset
' Assume MS-ACCESS table
   Dim ConnStr As String, DbName As String

   On Error Resume Next

   If Db.TableDefs(TableName).Connect <> "" Then   ' está linkeada ?
      ConnStr = Db.TableDefs(TableName).Connect

      DbName = GetTxConnectInfo(ConnStr, "DATABASE")
      ConnStr = GetTxConnectInfo(ConnStr, "PWD")
      If ConnStr <> "" Then
         ConnStr = ";PWD=" & ConnStr & ";"
      End If

      Set DbOpenTable1 = DBEngine.Workspaces(0).OpenDatabase _
                    (DbName, False, False, ConnStr).OpenRecordset(TableName, _
                    dbOpenTable)
                    
'      Set DbOpenTable1 = DBEngine.Workspaces(0).OpenDatabase _
'                    (Mid(Db.TableDefs(TableName).Connect, 11), _
'                    False, False, "").OpenRecordset(TableName, _
'                    dbOpenTable)
                    
                    
   Else
      Set DbOpenTable1 = Db.OpenRecordset(TableName, dbOpenTable)
   End If
      
End Function
#End If

#If DATACON = DAO_CONN Then 'ADO
Public Sub ChkDbSize(Db As Database, ByVal MAXKBSIZE As Long)
   Dim FSize As Long
   
   If SqlType(Db, Db.Connect) <> SQL_ACCESS Then
      Exit Sub
   End If

   FSize = FileLen(Db.Name)

   If FSize / 1024 > MAXKBSIZE Then
      MsgBox1 "ATENCIÓN" & vbCrLf & vbCrLf & "El tamaño de la base de datos supera los " & Format(MAXKBSIZE / 1024, NUMFMT) & " MBytes," & vbCrLf & "es necesario que utilice la opción para compactarla.", vbExclamation
   End If

End Sub
#End If

' Para generar Scripts
Public Function ParaMySQL(ByVal Fld As String) As String

   Fld = ReplaceStr(Fld, "\", "\\") ' primero que los otros
   Fld = ReplaceStr(Fld, vbTab, "\t")
   Fld = ReplaceStr(Fld, vbCr, "\r")
   Fld = ReplaceStr(Fld, vbLf, "\n")
   Fld = ReplaceStr(Fld, "'", "\'")
   Fld = ReplaceStr(Fld, """", "\""")

   ParaMySQL = Fld
End Function

' Solo para MySQL
#If DATACON = DAO_CONN Then 'ADO
Public Function DbGetLastID(Db As Database) As Long
#Else
Public Function DbGetLastID(Db As Connection) As Long
#End If
   Dim Q1 As String, Rs As Recordset

   Q1 = "SELECT LAST_INSERT_ID()"
   Set Rs = OpenRs(DbMain, Q1)
   DbGetLastID = vFld(Rs(0))
   Call CloseRs(Rs)

End Function
' Para SQL Server, cuando el dato viene desde SAP el signo viene al final
' Viene como num- y debe quedar como -num
#If DATACON = DAO_CONN Then 'ADO
Public Function DbCorrigeSigno(Db As Database, ByVal TblName As String, ByVal FldName As String) As Long
#Else
Public Function DbCorrigeSigno(Db As Connection, ByVal TblName As String, ByVal FldName As String) As Long
#End If
   Dim Q1 As String, Rc As Long

   Q1 = "UPDATE " & TblName & " SET " & FldName & "= '-' + LTrim(Substring(" & FldName & ", 1, Len(" & FldName & ")-1))"
   Q1 = Q1 & " WHERE Right( " & FldName & " , 1 ) = '-'"
   Rc = ExecSQL(Db, Q1)

   DbCorrigeSigno = Rc
End Function

Public Function GetAccessVersion(ByVal DbFile As String) As Integer
   Dim intFormat As Integer, objAccess As Object

   On Error Resume Next

   Set objAccess = CreateObject("Access.Application")
   If objAccess Is Nothing Then
      GetAccessVersion = -1
      Exit Function
   End If

   objAccess.OpenCurrentDatabase DbFile
   
   intFormat = objAccess.CurrentProject.FileFormat
   
   Select Case intFormat
       Case 2: Debug.Print "Microsoft Access 2"
       Case 7: Debug.Print "Microsoft Access 95"
       Case 8: Debug.Print "Microsoft Access 97"
       Case 9: Debug.Print "Microsoft Access 2000"
       Case 10: Debug.Print "Microsoft Access 2003"
       Case 12: Debug.Print "Microsoft Access 2007"
       Case Else: Debug.Print "Versión desconocida"
   End Select

   GetAccessVersion = intFormat
End Function

#If DATACON = DAO_CONN Then 'dao
' OJO: ConnString no debe tener un ; en el primer caracter
Public Function DbTablePath(Db As Database, ByVal TableName As String) As String
   Dim i As Integer, j As String
   Dim ConnStr As String
   
   On Error Resume Next
   
   DbTablePath = ""
   
   ConnStr = Db.TableDefs(TableName).Connect
   
   i = InStr(1, ConnStr, "DATABASE=", vbTextCompare)
   If i <= 0 Then
      Exit Function
   End If
   
   i = i + 9
   j = InStr(i, ConnStr, ";", vbBinaryCompare)
   
   If j > 0 Then
      DbTablePath = Mid(ConnStr, i, j - i)
   Else
      DbTablePath = Mid(ConnStr, i)
   End If

End Function

#End If

' 21 ago 2018: se agrega esta función
Public Function SqlConcat(ByVal DbType As Integer, ByVal Item1 As String, ByVal Item2 As String, Optional ByVal Item3 As String = "", Optional ByVal Item4 As String = "", Optional ByVal Item5 As String = "", Optional ByVal Item6 As String = "") As String

   If DbType = SQL_ACCESS Then
      SqlConcat = Item1 & " & " & Item2
      
      If Item3 <> "" Then
         SqlConcat = SqlConcat & " & " & Item3
      End If
      
      If Item4 <> "" Then
         SqlConcat = SqlConcat & " & " & Item4
      End If
      
      If Item5 <> "" Then
         SqlConcat = SqlConcat & " & " & Item5
      End If
      
      If Item6 <> "" Then
         SqlConcat = SqlConcat & " & " & Item6
      End If
      
   Else
      SqlConcat = "Concat( " & Item1 & ", " & Item2
      
      If Item3 <> "" Then
         SqlConcat = SqlConcat & ", " & Item3
      End If
      
      If Item4 <> "" Then
         SqlConcat = SqlConcat & ", " & Item4
      End If
      
      If Item5 <> "" Then
         SqlConcat = SqlConcat & ", " & Item5
      End If
      
      If Item6 <> "" Then
         SqlConcat = SqlConcat & ", " & Item6
      End If
     
      SqlConcat = SqlConcat & " )"
      
   End If

End Function

Public Function SqlInt(ByVal Expresion As String, Optional ByVal DbType As Integer = 0) As String

   If DbType = 0 Then
      DbType = gDbType
   End If
   
   If DbType = SQL_ACCESS Then
      SqlInt = "Int(" & Expresion & ")"
   Else
      SqlInt = "Floor(" & Expresion & ")"
   End If

End Function
'Para SQLServer y Access
#If DATACON = DAO_CONN Then 'ADO
Public Function AdvTbAddNew(Db As Database, ByVal Tbl As String, ByVal FldId As String, ByVal FldName As String, Optional ByVal Value2FldName As String = "") As Long
#Else
Public Function AdvTbAddNew(Db As Connection, ByVal Tbl As String, ByVal FldId As String, ByVal FldName As String, Optional ByVal Value2FldName As String = "") As Long
#End If

   Dim InsQry As String

#If DATACON = DAO_CONN Then
   AdvTbAddNew = TbAddNew(Db, Tbl, FldId, FldName, Value2FldName)
#Else
   InsQry = "INSERT INTO " & Tbl & "(" & FldName & ") VALUES ('" & Value2FldName & "')"
   AdvTbAddNew = TbAddNew4(Db, InsQry, FldId)
#End If

End Function
'Para SQLServer y Access

#If DATACON = DAO_CONN Then 'ADO
Public Function AdvTbAddNewMult(Db As Database, ByVal Tbl As String, ByVal FldId As String, FldArray() As AdvTbAddNew_t, Optional ByVal Msg As Boolean = False) As Long
#Else
Public Function AdvTbAddNewMult(Db As Connection, ByVal Tbl As String, ByVal FldId As String, FldArray() As AdvTbAddNew_t, Optional ByVal Msg As Boolean = False) As Long
#End If

   Dim InsQry As String
   Dim InsQryF As String, InsQryV As String
   Dim Rs As Recordset
   Dim i As Integer

   On Error Resume Next

   AdvTbAddNewMult = -1

#If DATACON = DAO_CONN Then

      Set Rs = DbOpenTable1(Db, Tbl)
      
      If Rs Is Nothing Then
         SqlErr = ERR
         SqlError = ERR.Description
         Call AddLog("AdvTbAddNewMult: error: " & SqlErr & ", " & SqlError & ", en OpenRecorset/Execute con " & Tbl)
         If Msg Then
            MsgBox1 "Error " & SqlErr & ", " & SqlError, vbExclamation
         End If
         AdvTbAddNewMult = -1
         Exit Function
      End If
      
      Rs.AddNew
      AdvTbAddNewMult = vFld(Rs(FldId))
      
      For i = 0 To UBound(FldArray)
         If FldArray(i).FldName <> "" And FldArray(i).FldValue <> "" Then ' por si tiene un índice que no permite null en FldName
            Rs(FldArray(i).FldName) = FldArray(i).FldValue
         End If
      Next i
      
      Rs.Update
      
      If ERR Then
         SqlErr = ERR
         SqlError = ERR.Description
         Call AddLog("AdvTbAddNewMult: " & Tbl & ", error: " & SqlErr & ", " & SqlError)
         AdvTbAddNewMult = -1
      End If
      
      Call CloseRs(Rs)
   
'      InsQryF = "INSERT INTO " & Tbl & "("
'      InsQryV = ")VALUES("
'
'      For i = 0 To UBound(FldArray)
'
'         If FldArray(i).FldName <> "" And FldArray(i).FldValue <> "" Then ' por si tiene un índice que no permite null en FldName
'            InsQryF = InsQryF & FldArray(i).FldName & ","
'            InsQryV = InsQryV & "'" & FldArray(i).FldValue & "',"
'         End If
'
'      Next i
'
'      InsQry = Left(InsQryF, Len(InsQryF) - 1) & " " & Left(InsQryV, Len(InsQryV) - 1) & ")"


#Else
'      InsQry = "INSERT INTO " & Tbl & "(" & FldName & ")VALUES('" & Value2FldName & "')"
      InsQryF = "INSERT INTO " & Tbl & " ("
      InsQryV = ") VALUES ("
      
      For i = 0 To UBound(FldArray)
      
         If FldArray(i).FldName <> "" And FldArray(i).FldValue <> "" Then ' por si tiene un índice que no permite null en FldName
            InsQryF = InsQryF & FldArray(i).FldName & ","
            InsQryV = InsQryV & "'" & FldArray(i).FldValue & "',"
         End If
         
      Next i
      
      InsQry = Left(InsQryF, Len(InsQryF) - 1) & " " & Left(InsQryV, Len(InsQryV) - 1) & ")"
      
      AdvTbAddNewMult = TbAddNew4(Db, InsQry, FldId, Msg)
#End If

End Function

Public Function SqlTrim(ByVal Expr As String, Optional DbType As Byte = 0) As String

   If DbType = 0 Then
      DbType = gDbType
   End If

   If DbType = SQL_SERVER Then
      SqlTrim = "LTrim(Rtrim(" & Expr & "))"
   Else
      SqlTrim = "Trim(" & Expr & ")"
   End If

End Function
' 4 jun 2019: Para saber si son la misma base aunque apunten a rutas diferentes
Public Function SameMdb(ByVal Fn1 As String, ByVal Fn2 As String, ByVal bOpened As Boolean) As Boolean
   
   If StrComp(Fn1, Fn2, vbTextCompare) = 0 Then
      SameMdb = True
      Exit Function
   End If
   
   If FileSize(Fn1) <> FileSize(Fn2) Then
      Exit Function
   End If

   If FileDate(Fn1) <> FileDate(Fn2) Then
      Exit Function
   End If

   If bOpened Then
      Fn1 = Left(Fn1, Len(Fn1) - 4) & ".ldb"
      Fn2 = Left(Fn2, Len(Fn2) - 4) & ".ldb"

      If FileSize(Fn1) <> FileSize(Fn2) Then
         Exit Function
      End If

      If FileDate(Fn1) <> FileDate(Fn2) Then
         Exit Function
      End If

   End If

   SameMdb = True ' es la misma

End Function


Public Function SqlInStr(ByVal StrWhere As String, ByVal StrWhat As String, Optional Position As Integer = 1, Optional ByVal DbType As Integer = 0) As String

   If DbType = 0 Then
      DbType = gDbType
   End If
   
   If StrWhere = "''" Then
      SqlInStr = 0
      Exit Function
   End If
   
   If DbType = SQL_ACCESS Then
'      SqlInStr = "InStr(" & StrWhere & "," & StrWhat & "," & Position & ")" 3 ene 2020: no se usa la posición
      SqlInStr = "InStr(" & StrWhere & "," & StrWhat & ")"
   Else
      SqlInStr = "CharIndex(" & StrWhat & "," & StrWhere & "," & Position & ")"
   End If

End Function


#If DATACON = DAO_CONN Then 'ADO
Public Function DbGetName(Db As Database)
#Else
Public Function DbGetName(Db As Connection)
#End If
   
#If DATACON = DAO_CONN Then 'ADO
   If SqlType(Db) = SQL_ACCESS Then
      DbGetName = Db.Name
   Else
#End If
      DbGetName = GetConnectInfo(Db, "DATABASE")
#If DATACON = DAO_CONN Then 'ADO
   End If
#End If

End Function

Public Function SqlIdInsert(ByVal DbType As Integer, ByVal Tbl As String, ByVal bOn As Boolean) As String
   Dim Q1 As String
   
   If DbType = SQL_SERVER Then
      Q1 = "SET IDENTITY_INSERT " & Tbl & IIf(bOn, " ON;", " OFF;")
   End If

   ' en MySQL es siempre ON

   SqlIdInsert = Q1
End Function

#If DATACON = DAO_CONN Then 'ADO
Public Function SqlVersion(Db As Database) As String
#Else
Public Function SqlVersion(Db As Connection) As String
#End If
Dim DbType As Integer, Q1 As String, Rs As Recordset
   
   DbType = SqlType(Db)
   
   If DbType = SQL_SERVER Then
      Q1 = "SELECT @@Version"
      Set Rs = OpenRs(Db, Q1)
      If Not Rs Is Nothing Then
         SqlVersion = vFld(Rs(0))
      End If
      Call CloseRs(Rs)
   ElseIf DbType = SQL_MYSQL Then
      Q1 = "SELECT @@Version"
      Set Rs = OpenRs(Db, Q1)
      If Not Rs Is Nothing Then
         SqlVersion = "MySQL " & vFld(Rs(0))
      End If
      Call CloseRs(Rs)
#If DATACON = DAO_CONN Then
   ElseIf DbType = SQL_ACCESS Then
      SqlVersion = "MS Access " & Db.Connect & " - " & Db.Version
#End If
   End If

End Function

' Para hacer Paginamiento. El resultado de esta función se debe poner después del ORDER BY
' Para el primer registro, SkipRows debe ser cero
Public Function SqlPaging(ByVal DbType As Integer, ByVal SkipRows As Long, ByVal GetRows As Long) As String
   Dim Q1 As String

   If SkipRows < 0 Then
      SkipRows = 0
   End If

   If DbType = SQL_SERVER Then
      Q1 = " OFFSET " & SkipRows & " ROWS FETCH NEXT " & GetRows & " ROWS ONLY"
   ElseIf DbType = SQL_MYSQL Then
      Q1 = " LIMIT " & SkipRows & ", " & GetRows
#If DATACON = DAO_CONN Then
   ElseIf DbType = SQL_ACCESS Then
      Q1 = "" ' No tiene para paginamiento
#End If
   End If

   SqlPaging = Q1

End Function

Public Function SqlTxt(ByVal Txt As String) As String

   If Txt = "" Then
      SqlTxt = "NULL"
   Else
      SqlTxt = "'" & ParaSQL(Txt) & "'"
   End If

End Function


'feña
Public Sub SetDbSecurityCambio(ByVal DbPath As String, ByVal Passw As String, ByVal CfgFile As String, ByVal SegCfg As String, ConnStr As String, Tipo As String, Rut As String)
   Dim Db As Database
   Dim Cfg As String, ConnStr1 As String
   Dim Seg As Boolean
   Dim passOld As String
   Dim Q1 As String
   Dim Rs As Recordset
   
   If Tipo = "EMP" Then
        passOld = PASSW_PREFIX & Rut
   Else
        passOld = PASSW_LEXCONT
   End If
   
   
   On Error Resume Next
   
   Cfg = GetIniString(CfgFile, "Config", "Secur", "")
   'MsgBox "ruta : " & CfgFile & "clave CFG: " & Cfg & " CLAVE SegCfg: " & SegCfg
   If Cfg <> SegCfg Then
      Seg = True
      ConnStr1 = ";PWD=" & Passw & ";"
      ConnStr = ";PWD=" & Passw & ";"
   Else
      Seg = False
      ConnStr1 = ";PWD=" & Passw & ";"
      ConnStr = ";PWD=" & Passw & ";"
   End If

  ERR.Clear

  ' Probamos a abrir la base con lo contrario de la seguridad esperada
   Set Db = OpenDatabase(DbPath, True, False, ConnStr1)
   If Not Db Is Nothing Then ' si pudo abrir, entonces hay que cambiar
      
      If Seg Then
         'Db.NewPassword "", passOld   ' le pone clave
         'If Passw = PASSW_PREFIX_NEW & rut Or Passw = PASSW_LEXCONT_NEW Or Passw = PASSW_LEXCONT_NEW2 Then
            Db.NewPassword Passw, passOld
          '  Call AddLog("pass que tenia " & Passw & "  La vieja que se le puso " & passOld)
         'End If
      Else
         'Db.NewPassword passOld, ""   ' le quita la clave
         'If Passw = PASSW_PREFIX_NEW & rut Or Passw = PASSW_LEXCONT_NEW Or Passw = PASSW_LEXCONT_NEW2 Then
            Db.NewPassword Passw, passOld
            Call AddLog("pass que tenia " & Passw & "  La vieja que se le puso " & passOld)
          '  Else
           ' Db.NewPassword Passw, passOld
         'End If
         
      ' Db.NewPassword Passw, passOld
      End If
     
      If ERR Then
         Call AddLog("SetDbSecurity: Seg=" & Seg & ", " & DbPath & ", Error " & ERR & ", " & ERR.Description)
      End If
      
       Q1 = "SELECT TIPO FROM Param WHERE Tipo = 'CCLAVES'"
         Q1 = Q1 & " AND Valor = '1'"
         Set Rs = OpenRs(Db, Q1)
         
         If Rs.EOF = False Then
      
        Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('CCLAVES',0,'1')"
        Call ExecSQL(Db, Q1)
      End If
      
      Call CloseDb(Db)
   
   Else
   ConnStr1 = ";PWD=" & PASSW_PREFIX & Rut & ";"
   
   ERR.Clear
   
    Set Db = OpenDatabase(DbPath, True, False, ConnStr1)
   If Not Db Is Nothing Then ' si pudo abrir, entonces hay que cambiar
      
      If Seg Then
         'Db.NewPassword "", passOld   ' le pone clave
         'If Passw = PASSW_PREFIX_NEW & rut Or Passw = PASSW_LEXCONT_NEW Or Passw = PASSW_LEXCONT_NEW2 Then
            Db.NewPassword Passw, passOld
            Call AddLog("pass que tenia " & Passw & "  La vieja que se le puso " & passOld)
         'End If
      Else
         'Db.NewPassword passOld, ""   ' le quita la clave
         'If Passw = PASSW_PREFIX_NEW & rut Or Passw = PASSW_LEXCONT_NEW Or Passw = PASSW_LEXCONT_NEW2 Then
            Db.NewPassword passOld, passOld
          '  Call AddLog("pass que tenia " & Passw & "  La vieja que se le puso " & passOld)
         'End If
         
        ' Db.NewPassword Passw, passOld
      End If
      
      If ERR Then
         Call AddLog("SetDbSecurity: Seg=" & Seg & ", " & DbPath & ", Error " & ERR & ", " & ERR.Description)
      End If
      
      Call CloseDb(Db)
      Else
      
   
      Call AddLog("SetDbSecurity: No se pudo quitar/poner clave a la base de datos, " & DbPath & ", Error " & ERR & ", " & ERR.Description)
      Debug.Print "*** No se pudo quitar/poner clave a la base de datos."
     End If
    End If
End Sub
'#End If

Public Sub CambiarClaveOLD()
Dim Ano As Long
Dim sArchivo As String
Dim i As Integer
Dim DbName As String
Dim Q1 As String
Dim Rs As Recordset


DbName = gDbPath & "\" & BD_COMUN


'Si se utilizara este metodo, descomentar estas 3 lineas mas declarar las variables PASSW_LEXCONT..
'Call SetDbSecurityCambio(DbName, PASSW_LEXCONT_NEW, gCfgFile, SG_SEGCFG, gComunConnStr, "", "")
'Call SetDbSecurityCambio(DbName, PASSW_LEXCONT_NEW2, gCfgFile, SG_SEGCFG, gComunConnStr, "", "")
'Call SetDbSecurityCambio(DbName, PASSW_LEXCONT, gCfgFile, SG_SEGCFG, gComunConnStr, "", "")



'       Q1 = "SELECT TIPO FROM Param WHERE Tipo = 'CCLAVES'"
'         Q1 = Q1 & " AND Valor = '1'"
'         Set Rs = OpenRs(DbName, Q1)
'
'If Rs.EOF = False Then

    Ano = 2014
    For i = 0 To 10
        Ano = Ano + 1
        If ExistFile(gDbPath & "\Empresas\" & Ano) Then
          
        sArchivo = Dir(gDbPath & "\Empresas\" & Ano & "\")
        Do While sArchivo <> ""
        'List1.AddItem sArchivo
        DbName = gDbPath & "\Empresas\" & Ano & "\" & sArchivo
        'Call OpenDbEmp(Replace(sArchivo, ".mdb", ""), ano)
        'Descomentar estas 2 lineas si se utiliza este metodo
        'Call SetDbSecurityCambio(DbName, PASSW_PREFIX_NEW & Replace(sArchivo, ".mdb", ""), gCfgFile, SG_SEGCFG, gEmpresa.ConnStr, "EMP", Replace(sArchivo, ".mdb", ""))
        'Call SetDbSecurityCambio(DbName, PASSW_PREFIX & Replace(sArchivo, ".mdb", ""), gCfgFile, SG_SEGCFG, gEmpresa.ConnStr, "EMP", Replace(sArchivo, ".mdb", ""))
        'CloseDb (DbMain)
        sArchivo = Dir
        Loop
        
        End If
          
    Next i

'End If
End Sub
'fin feña
