Attribute VB_Name = "MCorrigeBaseAdm"
Option Explicit
Const MDBV21 = "Actualizav21.mdb"

Public Sub CorrigeBaseAdm()
   
   If Not CorrigeBaseAdm_2005_01 Then
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V21 Then    ' 30 Sept. 2005
      Exit Sub
   End If
    
   If Not CorrigeBaseAdm_V22 Then    ' 23 Enero 2006
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V23 Then    ' 17 Marzo 2006 'PS
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V24 Then    ' 30 Marzo 2006
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V25 Then    ' 24 Abril 2006
      Exit Sub
   End If

'   If Not CorrigeBaseAdm_V26 Then    ' 25 Agosto 2006
'      Exit Sub
'   End If


   'esto aún no va en una versión del corrige base hasta que lo liberemos
   If gFunciones.RazFinancieras Then
      Call CrearTblRazFin
      Call DefTiposRazFin
      Call DefRazonesFin
   End If
   
End Sub
Public Function CorrigeBaseAdm_V26() As Boolean   '25 Ago 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim Id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim MaxTipoDoc As Integer

   On Error Resume Next

   UpdOK = True

   Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      DbVer = 0
   Else
      DbVer = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)
   If DbVer = 26 And UpdOK = True Then   ' 25 Ago 2006
     
      'agregamos tipo doc "Venta sin Documento" en Libro Ventas
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'VSD'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Venta sin Documento', 'VSD', 'ACTIVO', -1)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
     
     If UpdOK Then
         DbVer = 27
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
     
   CorrigeBaseAdm_V26 = True
   
End Function
Public Function CorrigeBaseAdm_V25() As Boolean   '24 Abr 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim Id As Long
   Dim i As Integer
   Dim Rc As Long

   On Error Resume Next

   UpdOK = True

   Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      DbVer = 0
   Else
      DbVer = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)
   If DbVer = 25 And UpdOK = True Then   ' 24 Abr 2006
      Call CorrigeTipoCapPropio_1("PlanAvanzado")
      Call CorrigeTipoCapPropio_1("PlanIntermedio")
      Call CorrigeTipoCapPropio_1("PlanBasico")
     
     If UpdOK Then
         DbVer = 26
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
     
   CorrigeBaseAdm_V25 = True
   
End Function
Public Function CorrigeBaseAdm_V24() As Boolean   '30 Mar 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim Id As Long
   Dim i As Integer
   Dim Rc As Long

   On Error Resume Next

   UpdOK = True

   Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      DbVer = 0
   Else
      DbVer = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)
   If DbVer = 24 And UpdOK = True Then   ' 30 Marzo 2006
      Call CorrigeCodF22_1("PlanAvanzado")
      Call CorrigeCodF22_1("PlanIntermedio")
      Call CorrigeCodF22_1("PlanBasico")
     
     If UpdOK Then
         DbVer = 25
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
     
   CorrigeBaseAdm_V24 = True
   
End Function
Public Function CorrigeBaseAdm_V23() As Boolean   '17 Mar 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim Id As Long
   Dim i As Integer
   Dim Rc As Long

   On Error Resume Next

   UpdOK = True

   Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      DbVer = 0
   Else
      DbVer = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)
   If DbVer = 23 And UpdOK = True Then   ' 17 Marzo 2006
     Call AlterField(DbMain, "CodActiv", "Codigo", dbText, 8)
     
     If UpdOK Then
         DbVer = 24
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
     
   CorrigeBaseAdm_V23 = True
   
End Function

Public Function CorrigeBaseAdm_V22() As Boolean   '23 Enero 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim Id As Long
   Dim i As Integer
   Dim Rc As Long

   On Error Resume Next

   UpdOK = True

   Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      DbVer = 0
   Else
      DbVer = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)

      '--------------------- Versión 22 -----------------------------------
   
   If DbVer = 22 And UpdOK = True Then   ' 23 Ene 2006
   
   
      '--------------------- Codigos F22 en Tabla Cuentas -----------------
                  
      Err.Clear
      'eliminamos algunos códigos de exportación a Form 22, en los planes predefinidos:
      '  628 porque es un campo que se ingresa con detalles, no ingreso directo
      '  366 y 384 porque son campos calculados por Form 22
      
      Q1 = "UPDATE PlanAvanzado SET CodF22 = 0 WHERE CodF22 IN( 628, 366, 384)"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanIntermedio SET CodF22 = 0 WHERE CodF22 IN( 628, 366, 384)"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanBasico SET CodF22 = 0 WHERE CodF22 IN( 628, 366, 384)"
      Call ExecSQL(DbMain, Q1)
      
      
      If UpdOK Then
         DbVer = 23
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
         
     
   CorrigeBaseAdm_V22 = True
   
End Function

Public Function CorrigeBaseAdm_V21() As Boolean   '20 o 21
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim Id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim DbActMDB As String, ConnStr As String

   On Error Resume Next

   UpdOK = True

   Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      DbVer = 0
   Else
      DbVer = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)

      '--------------------- Versión 21 -----------------------------------
   
   If (DbVer = 20 Or DbVer = 21) And UpdOK = True Then   ' 30 Sept. 2005
   
   
         '--------------------- TipoDocs -----------------
                  
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
                  
      'agregamos campo CodF29ExCount    (Cod29 para Count de Exento de Docs que tienen Exento y Afecto)
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29ExCount", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29ExCount", vbExclamation
         UpdOK = False
      End If
      
      'agregamos campo CodF29Exento    (Cod29 para Total Exento de Docs que tienen Exento y Afecto)
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29Exento", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29Exento", vbExclamation
         UpdOK = False
      End If
      
      'actualizamos códigos que faltan
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & " CodF29Count = 586"
      Q1 = Q1 & ", CodF29IVA = 142"
      Q1 = Q1 & " WHERE Diminutivo = 'BOE' "
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & " CodF29Count = 110"
      Q1 = Q1 & ", CodF29IVA = 111"
      Q1 = Q1 & ", CodF29ExCount = 586"
      Q1 = Q1 & ", CodF29Exento = 142"
      Q1 = Q1 & " WHERE Diminutivo IN ('VEM', 'OTV') "
      Call ExecSQL(DbMain, Q1)
           
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & " CodF29ExCount = 586"
      Q1 = Q1 & ", CodF29Exento = 142"
      Q1 = Q1 & " WHERE Diminutivo IN ('LFV', 'NDV', 'FAV') "
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & " CodF29ExCount = 586"
      Q1 = Q1 & ", CodF29Exento = -142"
      Q1 = Q1 & " WHERE Diminutivo = 'NCV' "
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & " CodF29Neto = 142"
      Q1 = Q1 & ", CodF29IVA = 0"
      Q1 = Q1 & " WHERE Diminutivo = 'BOE' "
      Call ExecSQL(DbMain, Q1)
      
      
      '*** PS Actualización de tabla CODIGO DE ACTIVIDAD
      DbActMDB = gDbPath & "\" & MDBV21
      If ExistFile(DbActMDB) Then
         'ConnStr = "PWD=" & PASSW_LEXCONT & ";"
         
         Call ExecSQL(DbMain, "Drop Table " & "CodActiv")
         
         'La base de datos ActualizaV21.mdb no tiene Password
         'Call LinkMdbTable(DbMain, DbActMDB, "CodActiv", "CodActivNew", , , ConnStr)
         Call LinkMdbTable(DbMain, DbActMDB, "CodActiv", "CodActivNew")
         
         Q1 = "SELECT Codigo,Descripcion as Descrip,Version INTO CodActiv FROM CodActivNew "
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "CREATE UNIQUE INDEX Codigo ON CodActiv (Codigo) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Call ExecSQL(DbMain, "Drop Table " & "CodActivNew")
         
      Else
         UpdOK = False
         MsgBox1 "No se encontró archivo " & DbActMDB & "." & vbNewLine & vbNewLine & "No se actualizaron los nuevos Códigos de Actividad Económica.", vbExclamation + vbOKOnly
      End If
      
      '***
      
      If UpdOK Then
         DbVer = 22
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
         
     
   CorrigeBaseAdm_V21 = True
   
End Function
   
' Para hacer manteciones a ciertas tablas de la db LexContab, con manejo de versión
Public Function CorrigeBaseAdm_2005_01() As Boolean
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim Id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim MaxTipoDoc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next

   UpdOK = True

   Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      DbVer = 0
   Else
      DbVer = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)

   '--------------------- Versión 0 -----------------------------------

   If DbVer = 0 And UpdOK = True Then
   
      '--------------------- EmpresasAno -----------------------------------

      Set Tbl = DbMain.TableDefs("EmpresasAno")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("IdCompAper", dbLong)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.IdCompAper", vbExclamation
         UpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("NumLastCompUnico", dbLong)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.NumLastCompUnico", vbExclamation
         UpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("NumLastCompA", dbLong)  'número último comp. de apertura

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.NumLastA", vbExclamation
         UpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("NumLastCompE", dbLong)  'número último comp. de egreso

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.NumLastCompE", vbExclamation
         UpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("NumLastCompI", dbLong)  'número último comp. de ingreso

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.NumLastCompI", vbExclamation
         UpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("NumLastCompT", dbLong)  'número último comp. de traspaso

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.NumLastCompT", vbExclamation
         UpdOK = False
      End If

      '--------------------- Actualización Versión -------------------------

      If UpdOK Then
         DbVer = 1
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 1 -----------------------------------

   If DbVer = 1 And UpdOK = True Then
   
      Q1 = "CREATE TABLE LParam (Codigo SMALLINT, Valor char(255) NULL )"
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX Codigo ON LParam (Codigo) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "INSERT INTO LParam (Codigo, Valor) VALUES (1, NULL )"
      Rc = ExecSQL(DbMain, Q1, False)


      If UpdOK Then
         DbVer = 2
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 2 -----------------------------------

   If DbVer = 2 And UpdOK = True Then

      Q1 = "DROP TABLE IPC"
      Rc = ExecSQL(DbMain, Q1)

      Q1 = "CREATE TABLE IPC ( AnoMes long, pIPC float, vIPC float, fCM float)"
      Rc = ExecSQL(DbMain, Q1)

      Q1 = "CREATE UNIQUE INDEX AnoMes ON IPC (AnoMes) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1)

      If UpdOK Then
         DbVer = 3
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 3 -----------------------------------

   If DbVer = 3 And UpdOK = True Then ' 13 DIC 2004
   
      Q1 = "CREATE UNIQUE INDEX Perfil ON Perfiles (Nombre)"
      Call ExecSQL(DbMain, Q1)

      Q1 = "SELECT Max(idPerfil) as M FROM Perfiles"
      Set Rs = OpenRs(DbMain, Q1)
      Id = vFld(Rs("M")) + 1
      Call CloseRs(Rs)

      Q1 = "INSERT INTO Perfiles (idPerfil, Nombre, Privilegios, idApp)"
      Q1 = Q1 & " VALUES (" & Id & ",'(todo)', 65535, 0)"
      Rc = ExecSQL(DbMain, Q1)

      Q1 = "UPDATE UsuarioEmpresa INNER JOIN Usuarios ON UsuarioEmpresa.idUsuario = Usuarios.IdUsuario"
      Q1 = Q1 & " SET UsuarioEmpresa.idPerfil=" & Id
      Q1 = Q1 & " WHERE Usuarios.Usuario IN ('Usuario1', 'Usuario2', 'Usuario3')"
      Rc = ExecSQL(DbMain, Q1)

      If UpdOK Then
         DbVer = 4
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 4 -----------------------------------

   If DbVer = 4 And UpdOK = True Then ' 21 Ene 2005
   
      '--------------------- TipoDocs -----------------------------------
      
      'EmpresaVacia.mdb no tiene password
      Call LinkMdbTable(DbMain, gDbPath & "\EmpresaVacia.mdb", "TipoDocs", "TipoDocsEV")
      
      Q1 = "SELECT * INTO TipoDocs FROM TipoDocsEV"
      Call ExecSQL(DbMain, Q1, False)
      
      Call UnLinkTable(DbMain, "TipoDocsEV")
      
      Q1 = "CREATE UNIQUE INDEX TipoDoc ON TipoDocs (TipoLib, TipoDoc) WITH PRIMARY"
      Call ExecSQL(DbMain, Q1, False)
      
      DbMain.TableDefs.Refresh ' para que vea la nueva tabla
      
         '--------------------- TipoDocFijo -----------------------------------

      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("TipoDocFijo", dbBoolean)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.TipoDocFijo", vbExclamation
         UpdOK = False
      
      End If
                                 
      'ponemos en -1 este campo para todos, aún los que ya lo tenían
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TipoDocFijo=-1")
      
           '--------------------- Nuevos TiposDocs -----------------------------------
      
      'eliminamos tipo doc Otros de Libro de Compras
      Q1 = "DELETE * FROM TipoDocs WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'OTR'"
      Call ExecSQL(DbMain, Q1)
   
      'eliminamos tipo doc Liquidación Factura de Libro de Compras
      Q1 = "DELETE * FROM TipoDocs WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'LFC'"
      Call ExecSQL(DbMain, Q1)
            
      'agregamos tipo doc Factura de Compras en Libro Ventas
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FCV'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Factura de Compras', 'FCV', 'ACTIVO', -1)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)

      'agregamos tipo doc Liquidación Factura en Libro Ventas
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'LFV'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Liquidación Factura', 'LFV', 'ACTIVO', -1)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
  
      'agregamos tipo doc Boleta de Venta en Libro Ventas
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = '" & TDOC_BOLVENTA & "'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc) FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Boleta de Venta', '" & TDOC_BOLVENTA & "', 'ACTIVO', -1)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
      
      'agregamos tipo doc Devolución Venta Boleta en Libro Ventas
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = '" & TDOC_DEVVENTABOL & "'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Devolución Venta con Boleta', '" & TDOC_DEVVENTABOL & "', 'ACTIVO', -1)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
                
      'agregamos tipo doc Factura de Compras en Libro Compras
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'FCC'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & MaxTipoDoc & ", 'Factura de Compra', 'FCC', 'ACTIVO', -1)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
                 
      'agregamos tipo doc Form. Importaciones en Libro Compras
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'IMP'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & MaxTipoDoc & ", 'Form. Importaciones', 'IMP', 'ACTIVO', -1)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
   
      'agregamos tipo doc Form. Importaciones Exento en Libro Compras
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'IEX'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & MaxTipoDoc & ", 'Form. Import. Exenta', 'IEX', 'ACTIVO', -1)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)

      'agregamos campos de F29 en TipoDocs
      
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29Count", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29Count", vbExclamation
         UpdOK = False
      End If

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29Neto", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29Neto", vbExclamation
         UpdOK = False
      End If

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29IVA", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29IVA", vbExclamation
         UpdOK = False
      End If

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29IVADTE", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29IVADTE", vbExclamation
         UpdOK = False
      End If

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29AFCount", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29AFCount", vbExclamation
         UpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29AFIVA", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29AFIVA", vbExclamation
         UpdOK = False
      End If
         
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29RetHon", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29RetHon", vbExclamation
         UpdOK = False
      End If
  
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29RetDieta", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29RetDieta", vbExclamation
         UpdOK = False
      End If
   
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29IVARet3ro", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29IVARet3ro", vbExclamation
         UpdOK = False
      End If
      
      'actualizamos los códigos del formulario 29
      
         'primero limpiamos todos los códigos
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 0"
      Q1 = Q1 & ", CodF29Neto =0"
      Q1 = Q1 & ", CodF29IVA = 0"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 519"
      Q1 = Q1 & ", CodF29Neto =0"
      Q1 = Q1 & ", CodF29IVA = 520"
      Q1 = Q1 & ", CodF29IVADTE = 511"
      Q1 = Q1 & ", CodF29AFCount = 524"
      Q1 = Q1 & ", CodF29AFIVA = 525"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'FAC'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 584"
      Q1 = Q1 & ", CodF29Neto = 562"
      Q1 = Q1 & ", CodF29IVA = 0"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'FCE'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 527"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = -528"
      Q1 = Q1 & ", CodF29IVADTE = 511"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'NCC'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 531"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = 532"
      Q1 = Q1 & ", CodF29IVADTE = 511"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'NDC'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 519"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = 520"
      Q1 = Q1 & ", CodF29IVADTE = 511"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 39"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'FCC'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 534"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = 535"
      Q1 = Q1 & ", CodF29IVADTE = 511"
      Q1 = Q1 & ", CodF29AFCount = 536"
      Q1 = Q1 & ", CodF29AFIVA = 553"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'IMP'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 566"
      Q1 = Q1 & ", CodF29Neto = 560"
      Q1 = Q1 & ", CodF29IVA = 0"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'IEX'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 503"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = 502"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FAV'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 586"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = 142"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FVE'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 509"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = -510"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'NCV'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 512"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = 513"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'NDV'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 515"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = 587"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FCV'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 500"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = 501"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'LFV'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 110"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = 111"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'BOV'"
 
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 110"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = -111"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'DVB'"
 
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 585"
      Q1 = Q1 & ", CodF29Neto = 20"
      Q1 = Q1 & ", CodF29IVA = 0"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'EXP'"
 
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 585"
      Q1 = Q1 & ", CodF29Neto = -20"
      Q1 = Q1 & ", CodF29IVA = 0"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'NCE'"
 
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 585"
      Q1 = Q1 & ", CodF29Neto = 20"
      Q1 = Q1 & ", CodF29IVA = 0"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'NDE'"
      
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 0"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = 0"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 151"
      Q1 = Q1 & ", CodF29RetDieta = 153"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_RETEN & " AND Diminutivo = 'BOH'"
      
      Call ExecSQL(DbMain, Q1)

      If UpdOK Then
         DbVer = 5
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If
   
      '--------------------- Versión 5 -----------------------------------

   If DbVer = 5 And UpdOK = True Then ' 25 Ene 2005
   
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 515"
      Q1 = Q1 & ", CodF29Neto = 587"
      Q1 = Q1 & ", CodF29IVA = 0"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FCV'"
      
      Call ExecSQL(DbMain, Q1)

      If UpdOK Then
         DbVer = 6
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If
      
      '--------------------- Versión 6 -----------------------------------
      
   If DbVer = 6 And UpdOK = True Then ' 31 Ene 2005
   
      '--------------------- TipoDocs -----------------------------------
      
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("TieneAfecto", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.TieneAfecto", vbExclamation
         UpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("TieneExento", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.TieneExento", vbExclamation
         UpdOK = False
      End If
   
      Q1 = "UPDATE TipoDocs SET TieneAfecto = -1"
      Q1 = Q1 & " WHERE (TipoLib = " & LIB_COMPRAS & " AND Diminutivo NOT IN( 'FCE', 'IEX') )"
      Q1 = Q1 & " OR (TipoLib = " & LIB_VENTAS & " AND Diminutivo NOT IN( 'FVE', 'EXP', 'NCE', 'NDE', 'BOV', 'DVB') )"   'boletas de venta y devolución con boleta, no se ingresa afecto y exento, sólo el total y se calcula afecto e Iva automáticamente
      Q1 = Q1 & " OR (TipoLib = " & LIB_RETEN & ")"
               
      Call ExecSQL(DbMain, Q1)
               
      Q1 = "UPDATE TipoDocs SET TieneExento = -1"
      Q1 = Q1 & " WHERE (TipoLib = " & LIB_COMPRAS & ")"
      Q1 = Q1 & " OR (TipoLib = " & LIB_VENTAS & " AND Diminutivo NOT IN( 'BOV', 'DVB') )"   'boletas de venta y devolución con boleta, no se ingresa afecto y exento, sólo el total y se calcula afecto e Iva automáticamente
      Q1 = Q1 & " OR (TipoLib = " & LIB_RETEN & ")"
      
      Call ExecSQL(DbMain, Q1)
   
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("ExigeRUT", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.ExigeRut", vbExclamation
         UpdOK = False
      End If
   
      Q1 = "UPDATE TipoDocs SET ExigeRUT = -1"
      Q1 = Q1 & " WHERE (TipoLib = " & LIB_COMPRAS & " AND Diminutivo NOT IN( 'IMP', 'IEX') )"
      Q1 = Q1 & " OR (TipoLib = " & LIB_VENTAS & " AND Diminutivo NOT IN( 'BOV', 'DVB', 'EXP', 'NCE', 'NDE') )"
      Q1 = Q1 & " OR (TipoLib = " & LIB_RETEN & ")"
      
      Call ExecSQL(DbMain, Q1)
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("EsRebaja", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.EsRebaja", vbExclamation
         UpdOK = False
      End If

      Q1 = "UPDATE TipoDocs SET EsRebaja = -1"
      Q1 = Q1 & " WHERE (TipoLib = " & LIB_COMPRAS & " AND Diminutivo IN( 'NCC') )"
      Q1 = Q1 & " OR (TipoLib = " & LIB_VENTAS & " AND Diminutivo IN( 'NCV', 'NCE', 'DVB' ) )"
      
      Call ExecSQL(DbMain, Q1)
            
            
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 586"
      Q1 = Q1 & ", CodF29Neto = 142"
      Q1 = Q1 & ", CodF29IVA = 0"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FVE'"
      
      Call ExecSQL(DbMain, Q1)

      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = -110"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = -111"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'DVB'"
 
      Call ExecSQL(DbMain, Q1)


      If UpdOK Then
         DbVer = 7
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If
      
      '--------------------- Versión 7 -----------------------------------

   If DbVer = 7 And UpdOK = True Then ' 9 Mar 2005
         
      'actualizamos el idPadre de una cuenta del plan intermedio
      Q1 = "SELECT IdCuenta FROM PlanIntermedio WHERE Codigo = '3030100'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         Q1 = "UPDATE PlanIntermedio SET IdPadre = " & vFld(Rs("IdCuenta")) & " WHERE Codigo = '3030101'"
         Call ExecSQL(DbMain, Q1)
      End If
      
      Call CloseRs(Rs)
      
      Q1 = "UPDATE PlanAvanzado SET Nombre = 'DIETA' WHERE Codigo = '3010607'"
      Call ExecSQL(DbMain, Q1)
     
      If UpdOK Then
         DbVer = 8
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If
   
   
      '--------------------- Versión 8 -----------------------------------

   If DbVer = 8 And UpdOK = True Then ' 11 Mar 2005
   
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoValor")
      
      Err.Clear
      Set Fld = Tbl.CreateField("Multiple", dbBoolean)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.Multiple", vbExclamation
         UpdOK = False
      
      End If
                                 
      'ponemos en True este campo para los tipos de docs que corresponde
      Call ExecSQL(DbMain, "UPDATE TipoValor SET Multiple=-1 WHERE Valor = 'Afecto' OR Valor = 'Exento' OR Valor = 'Bruto' OR Valor = 'Honorarios sin Retención'")
      
      If UpdOK Then
         DbVer = 9
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
    
      '--------------------- Versión 9 -----------------------------------

   If DbVer = 9 And UpdOK = True Then ' 24 Mar 2005
   
      'agregamos tipo doc Boletas exentas a libro de ventas
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'BOE'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Boleta Exenta', 'BOE', 'ACTIVO', -1, 0, 0, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
      
      'agregamos tipo doc Ventas menores
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'VEM'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Venta Menor', 'VEM', 'ACTIVO', -1, 0, 0, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
      
      If UpdOK Then
         DbVer = 10
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
      '--------------------- Versión 10 -----------------------------------

   If DbVer = 10 And UpdOK = True Then ' 29 Mar 2005
   
      'agregamos tipo doc Otros a Libro de Compras
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'OTC'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & MaxTipoDoc & ", 'Otros', 'OTC', 'ACTIVO', -1, -1, -1, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
      
      'agregamos tipo doc Otros a Libro de Ventas
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'OTV'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Otros', 'OTV', 'ACTIVO', -1, -1, -1, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
      
      'agregamos campo DocImpExp a tabla TipoDocs, para marcas los tipos de docs que son de importación o exportación y por tanto llevan RUT extranjero
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("DocImpExp", dbBoolean)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.DocImpExp", vbExclamation
         UpdOK = False
      
      End If
      
      'seteamos el campo para los docs de importación/exportación
      
      Q1 = "UPDATE TipoDocs SET DocImpExp=1 WHERE Diminutivo IN('IMP', 'IEX', 'EXP', 'NCE', 'NDE')"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos el campo para los docs de boletas
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("DocBoletas", dbBoolean)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.DocBoletas", vbExclamation
         UpdOK = False
      
      End If
      
      'seteamos el campo para los docs de boletas
      Q1 = "UPDATE TipoDocs SET DocBoletas=1 WHERE Diminutivo IN('BOV', 'DVB', 'BOE', 'VEM')"
      Call ExecSQL(DbMain, Q1)
      
      If UpdOK Then
         DbVer = 11
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
      '--------------------- Versión 11 -----------------------------------

   If DbVer = 11 And UpdOK = True Then ' 15 Abr 2005
   
      'actualizamos códigos F29 de Boleta Exenta
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & "  CodF29Count = 110"
      Q1 = Q1 & ", CodF29Neto = 0"
      Q1 = Q1 & ", CodF29IVA = 111"
      Q1 = Q1 & ", CodF29IVADTE = 0"
      Q1 = Q1 & ", CodF29AFCount = 0"
      Q1 = Q1 & ", CodF29AFIVA = 0"
      Q1 = Q1 & ", CodF29RetHon = 0"
      Q1 = Q1 & ", CodF29RetDieta = 0"
      Q1 = Q1 & ", CodF29IVARet3ro = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'BOE'"
      Call ExecSQL(DbMain, Q1)
      
      'agregamos cuentas de IVA Irrecuperable a planes de cuenta predefinidos
            
      'PlanAvanzado - Costo Directo
      Q1 = "SELECT IdCuenta FROM PlanAvanzado WHERE Codigo='3010400'"
      Set Rs = OpenRs(DbMain, Q1)
      
      '*** 16 MAY 2005 PAM - Se cambia "Rs.EOF = True" por "Rs.EOF = False"
      If Rs.EOF = False Then ' existe ?
         IdCtaPadre = vFld(Rs("IdCuenta"))
         
         Q1 = "INSERT INTO PlanAvanzado "
         Q1 = Q1 & " (idPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22)"
         Q1 = Q1 & " VALUES(" & IdCtaPadre & ", '3010442', 'IVAIRRECU', 'I.V.A. No Recuperable',"
         Q1 = Q1 & " ' ', 4, " & ECTA_ACTIVA & ", " & CLASCTA_RESULTADO & ", 0, 0, 0, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
      
      'PlanAvanzado - Costo No del Giro
      Q1 = "SELECT IdCuenta FROM PlanAvanzado WHERE Codigo='3020700'"
      Set Rs = OpenRs(DbMain, Q1)
      
      '*** 16 MAY 2005 PAM - Se cambia "Rs.EOF = True" por "Rs.EOF = False"
      If Rs.EOF = False Then  ' existe ?
         IdCtaPadre = vFld(Rs("IdCuenta"))
         
         Q1 = "INSERT INTO PlanAvanzado "
         Q1 = Q1 & " (idPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22)"
         Q1 = Q1 & " VALUES(" & IdCtaPadre & ", '3020705', ' ', 'I.V.A. No Recuperable',"
         Q1 = Q1 & " ' ', 4, " & ECTA_ACTIVA & ", " & CLASCTA_RESULTADO & ", 0, 0, 0, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
   
      'PlanIntermedio - Costo Directo
      Q1 = "SELECT IdCuenta FROM PlanIntermedio WHERE Codigo='3010400'"
      Set Rs = OpenRs(DbMain, Q1)
      
      '*** 16 MAY 2005 PAM - Se cambia "Rs.EOF = True" por "Rs.EOF = False"
      If Rs.EOF = False Then ' existe
         IdCtaPadre = vFld(Rs("IdCuenta"))
         
         Q1 = "INSERT INTO PlanIntermedio "
         Q1 = Q1 & " (idPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22)"
         Q1 = Q1 & " VALUES(" & IdCtaPadre & ", '3010442', 'IVAIRRECU', 'I.V.A. No Recuperable',"
         Q1 = Q1 & " ' ', 4, " & ECTA_ACTIVA & ", " & CLASCTA_RESULTADO & ", 0, 0, 0, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
      
      'PlanIntermedio - Costo No del Giro
      Q1 = "SELECT IdCuenta FROM PlanIntermedio WHERE Codigo='3020700'"
      Set Rs = OpenRs(DbMain, Q1)
      
      '*** 16 MAY 2005 PAM - Se cambia "Rs.EOF = True" por "Rs.EOF = False"
      If Rs.EOF = False Then  ' existe ?
         IdCtaPadre = vFld(Rs("IdCuenta"))
         
         Q1 = "INSERT INTO PlanIntermedio "
         Q1 = Q1 & " (idPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22)"
         Q1 = Q1 & " VALUES(" & IdCtaPadre & ", '3020705', ' ', 'I.V.A. No Recuperable',"
         Q1 = Q1 & " ' ', 4, " & ECTA_ACTIVA & ", " & CLASCTA_RESULTADO & ", 0, 0, 0, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
      
      'PlanBasico - Costo Directo
      Q1 = "SELECT IdCuenta FROM PlanBasico WHERE Codigo='3010400'"
      Set Rs = OpenRs(DbMain, Q1)
      
      '*** 16 MAY 2005 PAM - Se cambia "Rs.EOF = True" por "Rs.EOF = False"
      If Rs.EOF = False Then  ' existe
         IdCtaPadre = vFld(Rs("IdCuenta"))
         
         Q1 = "INSERT INTO PlanBasico "
         Q1 = Q1 & " (idPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22)"
         Q1 = Q1 & " VALUES(" & IdCtaPadre & ", '3010442', 'IVAIRRECU', 'I.V.A. No Recuperable',"
         Q1 = Q1 & " ' ', 4, " & ECTA_ACTIVA & ", " & CLASCTA_RESULTADO & ", 0, 0, 0, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
      
      'PlanBasico - Costo No del Giro
      Q1 = "SELECT IdCuenta FROM PlanBasico WHERE Codigo='3020700'"
      Set Rs = OpenRs(DbMain, Q1)
      
      '*** 16 MAY 2005 PAM - Se cambia "Rs.EOF = True" por "Rs.EOF = False"
      If Rs.EOF = False Then  ' existe
         IdCtaPadre = vFld(Rs("IdCuenta"))
         
         Q1 = "INSERT INTO PlanBasico "
         Q1 = Q1 & " (idPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22)"
         Q1 = Q1 & " VALUES(" & IdCtaPadre & ", '3020705', ' ', 'I.V.A. No Recuperable',"
         Q1 = Q1 & " ' ', 4, " & ECTA_ACTIVA & ", " & CLASCTA_RESULTADO & ", 0, 0, 0, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
            
      'ponemos en True este campo para los tipos de docs que corresponde
      Call ExecSQL(DbMain, "UPDATE TipoValor SET Multiple=-1")
      Call ExecSQL(DbMain, "UPDATE TipoValor SET Multiple=0 WHERE (TipoLib IN (" & LIB_COMPRAS & "," & LIB_VENTAS & ") AND Valor = 'Total Documento') OR (TipoLib = " & LIB_RETEN & " AND (Valor = 'Neto' OR Valor = 'Impuesto'))")

      
      'seteamos el campo DocImpExp=1 para doc 'OTC' que se nos había quedado afuera
      
      Q1 = "UPDATE TipoDocs SET DocImpExp=1 WHERE Diminutivo IN('OTC')"
      Call ExecSQL(DbMain, Q1)
      
      
      If UpdOK Then
         DbVer = 12
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
  
   End If
      
    '--------------------- Versión 12 -----------------------------------
   
   If DbVer = 12 And UpdOK = True Then   '28 Abril 2005
   
      Q1 = "UPDATE PlanAvanzado SET Atrib" & ATRIB_CAJA & "=1 WHERE " & GenLike(DbMain, "Caja", "Descripcion", GL_RWILD)
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanIntermedio SET Atrib" & ATRIB_CAJA & "=1 WHERE " & GenLike(DbMain, "Caja", "Descripcion", GL_RWILD)
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanBasico SET Atrib" & ATRIB_CAJA & "=1 WHERE " & GenLike(DbMain, "Caja", "Descripcion", GL_RWILD)
      Call ExecSQL(DbMain, Q1)
      
      '--------------------- Actualización Versión -------------------------
      If UpdOK Then
         DbVer = 13
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
    '--------------------- Versión 13 -----------------------------------
   
   If DbVer = 13 And UpdOK = True Then   '29 Abril 2005
   
      Q1 = "UPDATE PlanAvanzado SET TipoCapPropio=" & CAPPROPIO_PASIVO_NOEXIGIBLE & " WHERE Codigo IN('2030101','2030201', '2030301', '2030401','2030501','2031101','2031201','2031301')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanIntermedio SET TipoCapPropio=" & CAPPROPIO_PASIVO_NOEXIGIBLE & " WHERE Codigo IN('2030101','2030201', '2030301', '2030401','2030501','2031101','2031201','2031301')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanBasico SET TipoCapPropio=" & CAPPROPIO_PASIVO_NOEXIGIBLE & " WHERE Codigo IN('2030101','2030201', '2030301', '2030401','2030501','2031101','2031201','2031301')"
      Call ExecSQL(DbMain, Q1)
      
      '--------------------- Actualización Versión -------------------------
      If UpdOK Then
         DbVer = 14
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
      
    '--------------------- Versión 14 -----------------------------------
   
   If DbVer = 14 And UpdOK = True Then   ' 13 mayo 2005
   
      'actualizamos el idPadre de una cuenta del plan intermedio
      Q1 = "SELECT IdCuenta FROM PlanIntermedio WHERE Codigo = '3050100'"
      Set Rs = OpenRs(DbMain, Q1)
      
      Q1 = ""
      If Rs.EOF = False Then
         Q1 = "UPDATE PlanIntermedio SET IdPadre = " & vFld(Rs("IdCuenta")) & " WHERE Codigo = '3050101'"
      End If
      Call CloseRs(Rs)
           
      If Q1 <> "" Then
         Call ExecSQL(DbMain, Q1) ' se actualiza después de cerrar el Recordset porque podría estar bloqueado
      End If
           
      If UpdOK Then
         DbVer = 15
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
    '--------------------- Versión 15 -----------------------------------
    
   If DbVer = 15 And UpdOK = True Then   ' 14 julio 2005
   
      Q1 = "CREATE TABLE Equipos (PC char(50), MAC char(18), CodPC char(15), Aut BIT)"
      Rc = ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE UNIQUE INDEX PC_MAC ON Equipos (PC, MAC ) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1)
              
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor ) "
      Q1 = Q1 & " VALUES ('VER', 1, '" & VER_DEMO & "' )"
      Rc = ExecSQL(DbMain, Q1)   ' Version
              
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor ) "
      Q1 = Q1 & " VALUES ('VER', 2, '0' )"  ' codigo red
      Rc = ExecSQL(DbMain, Q1)
              
      If UpdOK Then
         DbVer = 16
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   '--------------------- Versión 16 -----------------------------------
   
   If DbVer = 16 And UpdOK = True Then   ' 14 julio 2005
   
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor ) "
      Q1 = Q1 & " VALUES ('VER', 3, '0' )" ' RUT
      Rc = ExecSQL(DbMain, Q1)
                            
      Q1 = "DROP INDEX PC_MAC ON Equipos"
      Rc = ExecSQL(DbMain, Q1)
                            
      Q1 = "CREATE UNIQUE INDEX PC_MAC_COD ON Equipos (PC, MAC, CodPC ) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1)

      If UpdOK Then
         DbVer = 17
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   '--------------------- Versión 17 -----------------------------------
   
   If DbVer = 17 And UpdOK = True Then   ' 21 julio 2005
      
      Call CrearTblContEmpresa
      
      If UpdOK Then
         DbVer = 18
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
     
   '--------------------- Versión 18 -----------------------------------
   
   If DbVer = 18 And UpdOK = True Then   ' 22 julio 2005
      
      'agregamos tipo doc "Boleta de Retención a Terceros" en Libro Retenciones
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_RETEN & " AND Diminutivo = 'BRT'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc)  FROM TipoDocs WHERE TipoLib=" & LIB_RETEN
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo)"
         Q1 = Q1 & " VALUES(" & LIB_RETEN & "," & MaxTipoDoc & ", 'Boleta de Ret. a Terceros', 'BRT', 'ACTIVO', -1)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
           
      If UpdOK Then
         DbVer = 19
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
      '--------------------- Versión 19 -----------------------------------
   
   If DbVer = 19 And UpdOK = True Then   ' 18 Agosto 2005

      'agregamos tipo valor IVA Activo Fijo a Libro de Compras
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA Activo Fijo'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)

         'obtenemos el máximo
         Q1 = "SELECT Max(Codigo) FROM TipoValor WHERE TipoLib=" & LIB_COMPRAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If

         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & MaxTipoDoc & ", 'IVA Activo Fijo', ' ', ' ', -1)"
         Call ExecSQL(DbMain, Q1)

      End If

      Call CloseRs(Rs)
      
      'eliminamos Count de Devoluciones de venta con Boleta (Laura Cabrera)
      Q1 = "UPDATE TipoDocs SET "
      Q1 = Q1 & " CodF29Count = 0"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'DVB'"
 
      Call ExecSQL(DbMain, Q1)

      If UpdOK Then
         DbVer = 20
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   
   CorrigeBaseAdm_2005_01 = UpdOK
      
End Function

Private Function CrearTblContEmpresa() As Boolean
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   CrearTblContEmpresa = True
      
   '--------------------- Crear tabla ControlEmpresa -----------------
   
   Set Tbl = New TableDef
   Tbl.Name = "ControlEmpresa"
   
   Err.Clear
   Set Fld = Tbl.CreateField("IdEmpresa", dbLong)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.IdEmpresa", vbExclamation
      CrearTblContEmpresa = False
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("Ano", dbInteger)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.Ano", vbExclamation
      CrearTblContEmpresa = False
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("RazonSocial", dbText, 200)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.RazonSocial", vbExclamation
      CrearTblContEmpresa = False
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("RUT", dbText, 12)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.RUT", vbExclamation
      CrearTblContEmpresa = False
   End If
   
   'creamos los meses
   For i = 1 To 12
      
      Err.Clear
      Set Fld = Tbl.CreateField("Mes" & i, dbByte)
      Tbl.Fields.Append Fld
      
      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.Mes" & i, vbExclamation
         CrearTblContEmpresa = False
      End If
   
   Next i
   
   'Activo Fijo
   Err.Clear
   Set Fld = Tbl.CreateField("AF_Depreciacion", dbByte)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.AF_Depreciacion", vbExclamation
      CrearTblContEmpresa = False
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("AF_CM", dbByte)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.AF_CM", vbExclamation
      CrearTblContEmpresa = False
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("AF_33BisLir", dbByte)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.AF_33BisLir", vbExclamation
      CrearTblContEmpresa = False
   End If

   'Corrección Monetaria
   
   Err.Clear
   Set Fld = Tbl.CreateField("CM_Activos", dbByte)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.CM_Activos", vbExclamation
      CrearTblContEmpresa = False
   End If
  
   Err.Clear
   Set Fld = Tbl.CreateField("CM_Pasivos", dbByte)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.CM_Pasivos", vbExclamation
      CrearTblContEmpresa = False
   End If
   
   'Balance definitivo
   
   Err.Clear
   Set Fld = Tbl.CreateField("BalDefinitivo", dbByte)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.BalDefinitivo", vbExclamation
      CrearTblContEmpresa = False
   End If
   
   'CPT Municipalidad
   
   Err.Clear
   Set Fld = Tbl.CreateField("CPT_Municip", dbByte)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.CPT_Municip", vbExclamation
      CrearTblContEmpresa = False
   End If
   
   'F22 Renta
   
   Err.Clear
   Set Fld = Tbl.CreateField("F22Renta", dbByte)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.F22Renta", vbExclamation
      CrearTblContEmpresa = False
   End If
   
   DbMain.TableDefs.Append Tbl
   If Err = 0 Then
      DbMain.TableDefs.Refresh
      
      Q1 = "CREATE UNIQUE INDEX IdEmpresa ON ControlEmpresa (Ano, IdEmpresa)"
      Rc = ExecSQL(DbMain, Q1)
      
   ElseIf Err <> 3010 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla ControlEmpresa", vbExclamation
      CrearTblContEmpresa = False
      
   End If
   
   Set Tbl = Nothing
      
End Function
'Esta función debe ser invocada desde CorrigeBaseAdm y
'Se debe agregar en link de la tabla en las db empresa-año
Private Function CrearTblRazFin() As Boolean
   Dim Q1 As String
   Dim Rc As Long
   Dim Tbl As TableDef, Fld As Field
   
   On Error Resume Next
   
   'Creamos tabla RazonesFin
   
   CrearTblRazFin = True
   
   Set Tbl = New TableDef
   Tbl.Name = "RazonesFin"
   
   Err.Clear
   Set Fld = Tbl.CreateField("IdRazon", dbLong)
   Fld.Attributes = dbAutoIncrField
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      CrearTblRazFin = False
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "RazonesFin.IdRazon", vbExclamation
   End If
      
   Err.Clear
   Set Fld = Tbl.CreateField("Tipo", dbInteger)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      CrearTblRazFin = False
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "RazonesFin.Tipo", vbExclamation
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("Nombre", dbText, 50)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      CrearTblRazFin = False
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "RazonesFin.Nombre", vbExclamation
   End If
      
   Err.Clear
   Set Fld = Tbl.CreateField("UnidadRes", dbText, 10)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      CrearTblRazFin = False
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "RazonesFin.UnidadRes", vbExclamation
   End If
      
   Err.Clear
   Set Fld = Tbl.CreateField("TxtNumerador", dbText, 50)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      CrearTblRazFin = False
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "RazonesFin.TxtNumerador", vbExclamation
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("TxtDenominador", dbText, 50)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      CrearTblRazFin = False
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "RazonesFin.TxtDenominador", vbExclamation
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("Operador", dbText, 1)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      CrearTblRazFin = False
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "RazonesFin.Operador", vbExclamation
   End If
      
   DbMain.TableDefs.Append Tbl
   Set Tbl = Nothing

   DbMain.TableDefs.Refresh

   Q1 = "CREATE UNIQUE INDEX IdRazon ON RazonesFin (IdRazon) WITH PRIMARY"
   Rc = ExecSQL(DbMain, Q1)
   
   Q1 = "CREATE UNIQUE INDEX Nombre ON RazonesFin (Nombre)"
   Rc = ExecSQL(DbMain, Q1)

   'Creamos tabla CuentasRazon

   CrearTblRazFin = True
   
   Set Tbl = New TableDef
   Tbl.Name = "CuentasRazon"
   
   Err.Clear
   Set Fld = Tbl.CreateField("IdRazon", dbLong)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      CrearTblRazFin = False
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "CuentasRazon.IdRazon", vbExclamation
   End If
      
   Err.Clear
   Set Fld = Tbl.CreateField("NumDenom", dbInteger)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      CrearTblRazFin = False
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "CuentasRazon.NumDenom", vbExclamation
   End If
      
   Err.Clear
   Set Fld = Tbl.CreateField("IdCuenta", dbLong)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      CrearTblRazFin = False
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "CuentasRazon.IdCuenta", vbExclamation
   End If
            
   Err.Clear
   Set Fld = Tbl.CreateField("Operador", dbText, 1)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      CrearTblRazFin = False
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "CuentasRazon.Operador", vbExclamation
   End If
   
   DbMain.TableDefs.Append Tbl
   Set Tbl = Nothing

   DbMain.TableDefs.Refresh

   Q1 = "CREATE UNIQUE INDEX IdCtaRazon ON CuentasRazon (IdRazon, NumDenom, IdCuenta) WITH PRIMARY"
   Rc = ExecSQL(DbMain, Q1)
   

End Function
'Define tipos de razones financieras
Private Sub DefTiposRazFin()
   Dim Q1 As String
   
   On Error Resume Next
   
   'agregamos tipos de Razones Financieras
   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', 1, 'Endeudamiento')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', 2, 'Liquidez')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', 3, 'Rentabilidad')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', 4, 'Rotaciones')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', 5, 'Consolidación')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', 6, 'Obsolescencia')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', 7, 'Otros')"
   Call ExecSQL(DbMain, Q1)

End Sub

Private Sub DefRazonesFin()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tipo As Integer
   
   On Error Resume Next
   
   '1.- Endeudamiento
   
   Tipo = 1
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Deuda/Patrimonio E', 'Veces', 'Deuda Total', 'Patrimonio', '/')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", '" & ParaSQL("Deuda/Patrimonio E'") & "', 'Veces', 'Deuda Total', 'Patrimonio', '/')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Total Activos/Total Deuda', 'Veces', 'Total Activos', 'Total Deuda', '/')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Cobertura de Intereses', 'Veces', 'Utilidad Operacional', 'Gastos Financieros', '/')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Endeudamiento Corto Plazo', 'Veces', 'Deuda Corto Plazo', 'Patrimonio', '/')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Endeudamiento Largo Plazo', 'Veces', 'Deuda Largo Plazo', 'Patrimonio', '/')"
   Call ExecSQL(DbMain, Q1)

   '2.- Liquidez
   
   Tipo = 2
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Razón Circulante', 'Veces', 'Activos Circulantes', 'Pasivos Circulantes', '/')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Razón Ácida', 'Veces', 'Activos Circulantes - Existencias', 'Pasivos Circulantes', '/')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Razón de Tesorería', 'Veces', 'Efectivo + Efectivo Equivalente', 'Pasivos Circulantes', '/')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Márgen de Maniobra', '$', 'Activos Circulantes', 'Pasivos Circulantes', '-')"
   Call ExecSQL(DbMain, Q1)

   '3.- Rentabilidad
   
   Tipo = 3
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Utilidades sobre Venta', '%', 'Utilidades', 'Ventas', '/')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Utilidades sobre Activos', '%', 'Utilidades', 'Activos', '/')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Utilidades sobre Patrimonio', '%', 'Utilidades', 'Patrimonio', '/')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'ROA', '%', 'Utilidades antes de Impuestos', 'Activo Total Neto', '/')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'ROE', '%', 'Utilidad Líquida', 'Patrimonio', '/')"
   Call ExecSQL(DbMain, Q1)
    
   '4.- Rotaciones
   
   Tipo = 4
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Rotación de Cuentas por Cobrar', 'Veces', 'Ventas Brutas', 'Total Cuentas por Cobrar', '/')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Rotación de Cuentas por Pagar', 'Veces', 'Costo de Venta', 'Total Cuentas por Pagar', '/')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Rotación de Inventario', 'Veces', 'Costo de Venta', 'Inventario', '/')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Periodo Medio de Cobro', 'Días', 'Total Cuentas por Cobrar', 'Ventas Brutas/365', '/')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Periodo Medio de Pago', 'Días', 'Total Cuentas por Pagar', 'Costo de Ventas Bruto/365', '/')"
   Call ExecSQL(DbMain, Q1)
  
   '5.- Consolidación
   
   Tipo = 5
  
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Activos Consolidados/Activo Neto', '%', 'Valor Activos Consolidados', 'Total Activo Neto', '/')"
   Call ExecSQL(DbMain, Q1)
  
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Activos Consolidados/Activo Neto - Activo Circ.', '%', 'Valor Activos Consolidados', 'Total Activo Neto - Activo Circulante', '/')"
   Call ExecSQL(DbMain, Q1)
  
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Activos Consolidados/Activo Neto - Activo Circ.', '%', 'Valor Activos Consolidados', 'Total Activo Neto - Activo Circulante', '/')"
   Call ExecSQL(DbMain, Q1)
  
   '6.- Obsolescencia
   
   Tipo = 6

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Obsolescencia', '%', 'Depreciación Acumulada', 'Total Activo Fijo Bruto', '/')"
   Call ExecSQL(DbMain, Q1)

End Sub
'define las cuentas que intervienen en cada razón financiera
Private Sub DefCuentasRazFin()
   On Error Resume Next


End Sub
