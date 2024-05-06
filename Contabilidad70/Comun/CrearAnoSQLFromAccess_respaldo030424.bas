Attribute VB_Name = "CrearAnoSQLFromAccess"
Option Explicit
'Crea nuevo año, si no existe, copiando del año anterior si existe, o con DB vacía.
'Supone que está abierta la DB LexContab
'Genera Saldos de Apertura
'No genera Comprobante de Apertura
Public Function CrearNuevoAno(ByVal IdEmpresa As Long, ByVal Ano As Integer, ByVal Rut As String, ByVal NombreEmpresa As String) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim RutMdb As String
   Dim Msg As String
   Dim EmpVacia As Boolean
#If DATACON = 1 Then
   Dim DbActual As Database
#End If
   Dim PathDbActual As String
   Dim CopyErr As Boolean
   Dim FCierre As Long
   Dim DbPath As String
   Dim Frm As Form
   Dim Rc As Integer
   Dim ConnStr As String
   Dim IdCompAperTrib As Long
   Dim NuevoAnoVacio As Boolean
   Dim Wh As String, Fld As String, Fld2 As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
         
   'Chequeo si está creada la base de datos para nuevo año, si no, la creo
   RutMdb = Rut & ".mdb"
   
   EmpVacia = False
   
   Call AddLog("CrearNuevoAno: 1", 1)

   'veamos si existe el año
   If gEmprSeparadas Then
      If ExistFile(gDbPath & "\Empresas\" & Ano & "\" & RutMdb) = True Then   'existe archivo nuevo año en empresas separadas
         CrearNuevoAno = True
         Call AddLog("CrearNuevoAno: 2", 1)
         
        
         'PS 19/04/2006, Parche por posible PATO que exista, donde CREA el .MDB pero no lo agregó en la tabla EMPRESASANO
         Call CheckRcEmpAno(Ano, IdEmpresa)
         Exit Function
      End If
      
      Call AddLog("CrearNuevoAno: 3", 1)
      
   Else
      Q1 = "SELECT Ano FROM EmpresasAno WHERE Ano=" & Ano & " AND idEmpresa=" & IdEmpresa
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then   'existe nuevo año en empresas juntas
         Call CloseRs(Rs)
         CrearNuevoAno = True
         Exit Function
      End If
      Call CloseRs(Rs)
   
   End If
   
   CrearNuevoAno = False
   
  'No existe, lo creamos
  
   If MsgBox1("¡ADVERTENCIA!" & vbCrLf & "No existe información de la empresa " & NombreEmpresa & " para el año " & Ano & "." & vbCrLf & "¿ Desea crearla a partir del año anterior (si existe)?", vbYesNo Or vbDefaultButton1 Or vbQuestion) <> vbYes Then
      
      If MsgBox1("¡ADVERTENCIA!" & vbCrLf & "Si no desea partir de la información del año anterior, desea crear un nuevo año vacío para la empresa " & NombreEmpresa & "?", vbYesNo Or vbDefaultButton1 Or vbQuestion) <> vbYes Then
         Exit Function
          gEmpresa.NuevoAno = False
      Else
         NuevoAnoVacio = True
         'SF 14691055
         gEmpresa.NuevoAno = True
         'SF 14691055
      End If
   
   End If

   Msg = "ATENCIÓN! Para hacer esta operación nadie debe estar trabajando con la empresa " & NombreEmpresa & "." & vbNewLine & "¿Desea continuar?"
   'Msg = Msg & " Verifique que haya sido creada la cuenta para almacenar el resultado del ejercicio. Esta será utilizada para realizar el proceso de apertura del año siguiente."
   If MsgBox1(Msg, vbQuestion Or vbDefaultButton1 Or vbYesNo) <> vbYes Then
      Exit Function
   End If
   
   Call AddLog("CrearNuevoAno: 4", 1)
   
   On Error Resume Next
   
   If gEmprSeparadas Then
   
   
#If DATACON = 1 Then       'Access
   
      'Creo el directorio año, por si no existe
      Call CreateDir(Ano)
      Call CreateDir(Ano)
   
      Call AddLog("CrearNuevoAno: 5", 1)
   
      ERR.Clear
      
      'Copio la base de datos del año anterior (si existe) para el año siguiente
      If ExistFile(gDbPath & "\Empresas\" & Ano - 1 & "\" & RutMdb) And Not NuevoAnoVacio Then
      
         Call AddLog("CrearNuevoAno: 5.1", 1)
        
      
         'vemos si el año anterior está cerrado
         Q1 = "SELECT FCierre FROM EmpresasAno WHERE IdEmpresa=" & IdEmpresa & " AND Ano=" & Ano - 1
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then
            FCierre = vFld(Rs("FCierre"))
         End If
         
         Call CloseRs(Rs)
         
         If FCierre = 0 Then
            MsgBox1 "El año anterior aún no ha sido cerrado. No es posible abrir el nuevo año y generar saldos de apertura.", vbExclamation + vbOKOnly
            Exit Function
         End If
                        
         PathDbActual = DbMain.Name
         Call CloseDb(DbMain)
      
         ' pam: se hace el corrige base en la base del año anterior, por si acaso
         Rc = OpenDbEmp(Rut, Ano - 1)
         Call CorrigeBase
         Call CloseDb(DbMain)
         
         ERR.Clear
      
         Call FileCopy(gDbPath & "\Empresas\" & Ano - 1 & "\" & RutMdb, gDbPath & "\Empresas\" & Ano & "\" & RutMdb)
      
         If ERR = 70 Then
            'Acceso denegado. Todavia existe alguien trabajando
            MsgErr "No se puede copiar la base de datos año " & Ano - 1 & " al año " & Ano & ". Es posible que esté siendo utilizada por algún usuario, verifique esto antes de continuar."
            CopyErr = True   'no hacemos exit para reabrir la db actual antes de salir
            
         End If
         
         'If OpenDb(DbMain, PathDbActual) = False Then
         If gEmpresa.Rut <> "" Then
         
            Rc = OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
            
         Else
            Rc = OpenDbAdm()
         End If
         
         If Rc = False Then
            MsgBox1 "No se pudo volver a abrir la base de datos " & PathDbActual & ".", vbExclamation
            Exit Function
         End If
      
         If CopyErr Then
            Exit Function
         End If

      Else  'no existe año anterior o desea crearlo vacio (no a partir de año anterior), creamos a partir de la DB vacía
         If CrearMdbVacia(Ano, RutMdb) = False Then
            Exit Function
         End If
         
        
         DoEvents
         
         EmpVacia = True
         
      End If
#End If
      
   Else     'empresas juntas (SQL Server)
   
      'copiamos el contenido de las tablas que se duplican en el año siguiente, si es que hay datos
      
      If Not NuevoAnoVacio Then
          
         'vemos si el año anterior está cerrado
         Q1 = "SELECT FCierre FROM EmpresasAno WHERE IdEmpresa=" & IdEmpresa & " AND Ano=" & Ano - 1
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then
            FCierre = vFld(Rs("FCierre"))
            Call CloseRs(Rs)
         
         Else
            Call CloseRs(Rs)
            
            Rc = MsgBox1("No existe año anterior para esta empresa en la base SQL Server." & vbCrLf & vbCrLf & "¿Desea obtener los datos de esta empresa desde la base de datos Access del año anterior, si existe?", vbExclamation + vbYesNo + vbDefaultButton2)
            If Rc = vbYes Then
               CrearNuevoAno = CrearNuevoAnoSQLFromAccess(IdEmpresa, Ano, Rut, NombreEmpresa)
               Exit Function
            ElseIf Rc = vbCancel Then
               Exit Function
            ElseIf MsgBox1("Dado que no existe año anterior para esta empresa en la base SQL Server, debe partir de una empresa vacía." & vbCrLf & vbCrLf & "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
'            If MsgBox1("No existe año anterior para esta empresa en la base SQL Server, debe partir de una empresa vacía." & vbCrLf & vbCrLf & "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
               Exit Function
            Else
               NuevoAnoVacio = True
            End If
            
         End If
                  
         If Not NuevoAnoVacio Then   'preguntamos de nuevo porque el usuario puede cambiar de idea en las preguntas, no se acuerda si existe año anterior
         
            If FCierre = 0 Then
               MsgBox1 "El año anterior aún no ha sido cerrado. No es posible abrir el nuevo año y generar saldos de apertura.", vbExclamation + vbOKOnly
               Exit Function
            End If
         
            'Copiamos las Cuentas del año anterior
            Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, IdPadre, IdCuenta As IdCuentaOld, IdPadre as IdPadreOld, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22, Atrib1, Atrib2, Atrib3, Atrib4, Atrib5, Atrib6, Atrib7, Atrib8, Atrib9, Atrib10, CodF29, CorrelativoCheque, CodIFRS_EstRes, CodIFRS_EstFin, DebeTrib, HaberTrib, CodIFRS, CodF22_14Ter, TipoPartida, CodCtaPlanSII"
            Fld2 = " IdEmpresa, Ano, IdPadre, IdCuentaOld, IdPadreOld, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22, Atrib1, Atrib2, Atrib3, Atrib4, Atrib5, Atrib6, Atrib7, Atrib8, Atrib9, Atrib10, CodF29, CorrelativoCheque, CodIFRS_EstRes, CodIFRS_EstFin, DebeTrib, HaberTrib, CodIFRS, CodF22_14Ter, TipoPartida, CodCtaPlanSII"
            Q1 = "INSERT INTO Cuentas (" & Fld2 & ") SELECT " & Fld & " FROM Cuentas as Cuentas1 WHERE Cuentas1.IdEmpresa = " & IdEmpresa & " AND Cuentas1.Ano = " & Ano - 1
            Q1 = Q1 & " ORDER BY Cuentas1.IdCuenta"
            Call ExecSQL(DbMain, Q1)
                        
            'actualizamos los padres
            Tbl = " Cuentas "
            sFrom = " Cuentas "
            sFrom = sFrom & " INNER JOIN Cuentas As Cuentas1 ON Cuentas.IdPadreOld = Cuentas1.IdCuentaOld "
            sFrom = sFrom & " AND Cuentas.IdEmpresa = Cuentas1.IdEmpresa AND Cuentas.Ano = Cuentas1.Ano "
            sSet = " Cuentas.IdPadre = Cuentas1.IdCuenta "
            sWhere = " WHERE Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
            Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
            
            'Copiamos CuentasBásicas
            Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, IdCuenta as IdCuentaOld, Tipo, TipoLib, TipoValor, IdCuenta"
            Fld2 = " IdEmpresa, Ano, IdCuentaOld, Tipo, TipoLib, TipoValor, IdCuenta "
            Q1 = "INSERT INTO CuentasBasicas ( " & Fld2 & " ) SELECT " & Fld & " FROM CuentasBasicas WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
            Q1 = Q1 & " ORDER BY Id"
            Call ExecSQL(DbMain, Q1)
                        
            'actualizamos IdCuenta de Cuentas Básicas
            Tbl = " CuentasBasicas "
            sFrom = " CuentasBasicas "
            sFrom = sFrom & " INNER JOIN Cuentas ON CuentasBasicas.IdCuentaOld = Cuentas.IdCuentaOld "
            sFrom = sFrom & " AND Cuentas.IdEmpresa = CuentasBasicas.IdEmpresa AND Cuentas.Ano = CuentasBasicas.Ano "
            sSet = " CuentasBasicas.IdCuenta = Cuentas.IdCuenta "
            sWhere = " WHERE CuentasBasicas.IdEmpresa = " & IdEmpresa & " AND CuentasBasicas.Ano = " & Ano
            Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
           
            'Copiamos ImpAdic
            Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, TipoLib, TipoValor, IdCuenta, Tasa, EsRecuperable, CodCuenta"
            Fld2 = " IdEmpresa, Ano, TipoLib, TipoValor, IdCuenta, Tasa, EsRecuperable, CodCuenta "
            Q1 = "INSERT INTO ImpAdic ( " & Fld2 & " ) SELECT " & Fld & " FROM ImpAdic WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
            Q1 = Q1 & " ORDER BY IdImpAdic "
            Call ExecSQL(DbMain, Q1)
                        
            'actualizamos IdCuenta de los ImpAdic
            Tbl = " ImpAdic "
            sFrom = " ImpAdic "
            sFrom = sFrom & " INNER JOIN Cuentas ON ImpAdic.CodCuenta = Cuentas.Codigo "
            sFrom = sFrom & " AND Cuentas.IdEmpresa = ImpAdic.IdEmpresa AND Cuentas.Ano = ImpAdic.Ano "
            sSet = " ImpAdic.IdCuenta = Cuentas.IdCuenta "
            sWhere = " WHERE ImpAdic.IdEmpresa = " & IdEmpresa & " AND ImpAdic.Ano = " & Ano
            Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
           
            'Copiamos CtasAjustesExCont
            Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, TipoAjuste, IdItem, IdCuenta, CodCuenta"
            Fld2 = " IdEmpresa, Ano, TipoAjuste, IdItem, IdCuenta, CodCuenta "
            Q1 = "INSERT INTO CtasAjustesExCont ( " & Fld2 & " ) SELECT " & Fld & " FROM CtasAjustesExCont WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
            Q1 = Q1 & " ORDER BY IdCtaAjustes "
            Call ExecSQL(DbMain, Q1)
                        
            'actualizamos IdCuenta de los CtasAjustesExCont
            Tbl = " CtasAjustesExCont "
            sFrom = " CtasAjustesExCont "
            sFrom = sFrom & " INNER JOIN Cuentas ON CtasAjustesExCont.CodCuenta = Cuentas.Codigo "
            sFrom = sFrom & " AND Cuentas.IdEmpresa = CtasAjustesExCont.IdEmpresa AND Cuentas.Ano = CtasAjustesExCont.Ano "
            sSet = " CtasAjustesExCont.IdCuenta = Cuentas.IdCuenta "
            sWhere = " WHERE CtasAjustesExCont.IdEmpresa = " & IdEmpresa & " AND CtasAjustesExCont.Ano = " & Ano
            Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
           
            'Copiamos CtasAjustesExContRLI
            Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, TipoAjuste, IdGrupo, IdItem, IdCuenta, CodCuenta"
            Fld2 = " IdEmpresa, Ano, TipoAjuste, IdGrupo, IdItem, IdCuenta, CodCuenta "
            Q1 = "INSERT INTO CtasAjustesExContRLI ( " & Fld2 & " ) SELECT " & Fld & " FROM CtasAjustesExContRLI WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
            Q1 = Q1 & " ORDER BY IdCtaAjustesRLI "
            Call ExecSQL(DbMain, Q1)
                        
            'actualizamos IdCuenta de los CtasAjustesExContRLI
            Tbl = " CtasAjustesExContRLI "
            sFrom = " CtasAjustesExContRLI "
            sFrom = sFrom & " INNER JOIN Cuentas ON CtasAjustesExContRLI.CodCuenta = Cuentas.Codigo "
            sFrom = sFrom & " AND Cuentas.IdEmpresa = CtasAjustesExContRLI.IdEmpresa AND Cuentas.Ano = CtasAjustesExContRLI.Ano "
            sSet = " CtasAjustesExContRLI.IdCuenta = Cuentas.IdCuenta "
            sWhere = " WHERE CtasAjustesExContRLI.IdEmpresa = " & IdEmpresa & " AND CtasAjustesExContRLI.Ano = " & Ano
            Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
           
                      
            'Socios
            Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, RUT, Nombre, PjePart, MontoSuscrito, MontoPagado, IdCuentaAportes, IdCuentaRetiros, IdTipoSocio, Vigente"
            Fld2 = " IdEmpresa, Ano, RUT, Nombre, PjePart, MontoSuscrito, MontoPagado, IdCuentaAportes, IdCuentaRetiros, IdTipoSocio, Vigente"
            Q1 = "INSERT INTO Socios ( " & Fld2 & " ) SELECT " & Fld & " FROM Socios WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
            Q1 = Q1 & " ORDER BY IdSocio "
            Call ExecSQL(DbMain, Q1)
            
             EmpVacia = False
             
             'SF 14691055
          gEmpresa.NuevoAno = False
          'SF 14691055
             
            'Empresa
            Fld = Ano & " As Ano, id, Rut, NombreCorto, RazonSocial, ApPaterno, ApMaterno, Nombre, Calle, Numero, Dpto, Telefonos, Fax, Region, Comuna, Ciudad, Giro, ActEconom, CodActEconom, DomPostal, ComunaPostal, Email, Web, FechaConstitucion, FechaInicioAct, RepConjunta, RutRepLegal1, RepLegal1, RutRepLegal2, RepLegal2, Contador, RutContador, TipoContrib, TransaBolsa, Franq14bis, FranqLey18392, FranqDL600, FranqDL701, FranqDS341, Opciones, TContribFUT, Franq14ter, Franq14quater, ObligaLibComprasVentas, FranqRentaAtribuida, FranqSemiIntegrado, Franq14ASemiIntegrado, FranqProPymeTransp, FranqProPymeGeneral "
            Fld2 = " Ano, id, Rut, NombreCorto, RazonSocial, ApPaterno, ApMaterno, Nombre, Calle, Numero, Dpto, Telefonos, Fax, Region, Comuna, Ciudad, Giro, ActEconom, CodActEconom, DomPostal, ComunaPostal, Email, Web, FechaConstitucion, FechaInicioAct, RepConjunta, RutRepLegal1, RepLegal1, RutRepLegal2, RepLegal2, Contador, RutContador, TipoContrib, TransaBolsa, Franq14bis, FranqLey18392, FranqDL600, FranqDL701, FranqDS341, Opciones, TContribFUT, Franq14ter, Franq14quater, ObligaLibComprasVentas, FranqRentaAtribuida, FranqSemiIntegrado, Franq14ASemiIntegrado, FranqProPymeTransp, FranqProPymeGeneral "
            Q1 = "INSERT INTO Empresa ( " & Fld2 & " ) SELECT " & Fld & " FROM Empresa WHERE Id = " & IdEmpresa & " AND Ano = " & Ano - 1
            Call ExecSQL(DbMain, Q1)
            
            'ParamEmpresa
            Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, Tipo, Codigo, Valor, Valor as ValorOld "
            Fld2 = " IdEmpresa, Ano, Tipo, Codigo, Valor, ValorOld "
            Q1 = "INSERT INTO ParamEmpresa ( " & Fld2 & " ) SELECT " & Fld & " FROM ParamEmpresa WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
            Call ExecSQL(DbMain, Q1)
           
            'ParamEmpresa
            Tbl = " ParamEmpresa "
            sFrom = " ParamEmpresa "
            sFrom = sFrom & " INNER JOIN Cuentas ON " & SqlVal("ParamEmpresa.ValorOld") & " = Cuentas.IdCuentaOld "
            sFrom = sFrom & " AND Cuentas.IdEmpresa = ParamEmpresa.IdEmpresa AND Cuentas.Ano = ParamEmpresa.Ano "
            sSet = " ParamEmpresa.Valor = Cuentas.IdCuenta "
            sWhere = " WHERE ParamEmpresa.IdEmpresa = " & IdEmpresa & " AND ParamEmpresa.Ano = " & Ano & " AND Left(Tipo,3) = 'CTA'"
            Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
            
            
                     
         Else
            EmpVacia = True

         End If
         
      Else
         EmpVacia = True
       
      End If
        
      
      Q1 = "SELECT Count(*) as N FROM CT_Comprobante WHERE IdEmpresa = " & IdEmpresa
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
      
         If vFld(Rs("N")) = 0 Then    'no tiene comp tipo para esta emnpresa, agregamos los base
         
            Call ResetCompTipoEmpJuntas(IdEmpresa)
         
         End If
      End If
      
      Call CloseRs(Rs)
            
   End If
            
   Call AddLog("CrearNuevoAno: 6", 1)
      
   'Actualizo la Tabla que tiene los años que tiene datos la empresa
   Q1 = "INSERT INTO EmpresasAno (IdEmpresa, Ano, FApertura) VALUES ("
   Q1 = Q1 & IdEmpresa & "," & Ano & "," & CLng(Int(Now)) & ")"
   Call ExecSQL(DbMain, Q1)
   
   On Error GoTo 0
      
#If DATACON = 1 Then
      
   If gEmprSeparadas Then
   
      'Guardo la base de datos actual y abro la DB del nuevo año, para borrar los registros de las tablas que corresponden
      Set DbActual = DbMain
      Set DbMain = Nothing
      
      If OpenDbEmp(Rut, Ano) = False Then
         Exit Function
      End If
         
      If Not EmpVacia Then
      
         'borramos registros nuevo año
         Q1 = "Comprobante"
         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "MovComprobante"
         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "Documento"
         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "MovDocumento"
         Call DeleteSQL(DbMain, Q1, "")
         
'         Q1 = "ComprobanteFull"
'         Call DeleteSQL(DbMain, Q1, "")
'
'         Q1 = "MovComprobanteFull"
'         Call DeleteSQL(DbMain, Q1, "")
'
'         Q1 = "DocumentoFull"
'         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "LibroCaja"
         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "EstadoMes"
         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "Cartola"
         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "DetCartola"
         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "LogImpreso"
         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "MovActivoFijo"
         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "PropIVA_TotMensual"
         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "LogComprobantes"
         Call DeleteSQL(DbMain, Q1, "")
              
         Q1 = "ActFijoCompsFicha"
         Call DeleteSQL(DbMain, Q1, "")
         
         Q1 = "ActFijoFicha"
         Call DeleteSQL(DbMain, Q1, "")
                         
         Q1 = "AjustesExtLibCaja"
         Call DeleteSQL(DbMain, Q1, "")
                         
         Q1 = "AsistImpPrimCat"
         Call DeleteSQL(DbMain, Q1, "")
                         
         Q1 = "BaseImponible14Ter"
         Call DeleteSQL(DbMain, Q1, "")
                  
         'actualizamos año en algunas tablas
         Q1 = "UPDATE Cuentas SET Ano = " & Ano
         Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "UPDATE CuentasBasicas SET Ano = " & Ano
         Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "UPDATE ImpAdic SET Ano = " & Ano
         Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "UPDATE CtasAjustesExCont SET Ano = " & Ano
         Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "UPDATE CtasAjustesExContRLI SET Ano = " & Ano
         Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
         Call ExecSQL(DbMain, Q1)
 
         Q1 = "UPDATE Socios SET Ano = " & Ano
         Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "UPDATE ParamEmpresa SET Ano = " & Ano
         Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "UPDATE Empresa SET Ano = " & Ano
         Q1 = Q1 & " WHERE Id = " & IdEmpresa
         Call ExecSQL(DbMain, Q1)
         
         'asignamos IdEmpresa a Comprobante Tipo
         Q1 = "UPDATE CT_Comprobante SET IdEmpresa = " & IdEmpresa
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "UPDATE CT_MovComprobante SET IdEmpresa = " & IdEmpresa
         Call ExecSQL(DbMain, Q1)
         
      End If
      
   End If
#End If
      
   'Marcamos inicio año, indicando que se originó a partir de año anterior
   'Esto se usa para que en ReadEmpresa se genere el comprobante de apertura,
   'la primera vez que entra, sólo si Tipo =  INITAÑO, Codigo = 1 y Valor = EMPHISTORIA
      
   If Not EmpVacia Then
      Wh = " WHERE Tipo = 'INITAÑO'"
      Wh = Wh & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
      
      Q1 = "SELECT * FROM ParamEmpresa " & Wh
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF Then
         Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano)"
         Q1 = Q1 & " VALUES( 'INITAÑO', 1, 'EMPHISTORIA'," & IdEmpresa & "," & Ano & ")"
      Else
         Q1 = "UPDATE ParamEmpresa SET Valor = 'EMPHISTORIA', Codigo = 1 " & Wh
      End If
      
      Call ExecSQL(DbMain, Q1)
      Call CloseRs(Rs)
   
   Else 'empresa vacía
   
      Call AddLog("CrearNuevoAno: 6.1", 1)
   
      'ConnStr = "PWD=" & PASSW_LEXCONT & ";"
      'Call LinkMdbTable(DbMain, gDbPath & "\" & BD_COMUN, "EmpresasAno", , , , ConnStr)
      
#If DATACON = 1 Then
      If gEmprSeparadas Then
         Call LinkMdbTable(DbMain, gDbPath & "\" & BD_COMUN, "EmpresasAno", , , , gComunConnStr)
      End If
#End If
     
'      Call AddLog("CrearNuevoAno: 6.2", 1)
      
'      Call GenCompAperSinMovs(1, IdEmpresa, Ano, IdCompAperTrib)   'se movierona a IniEmpresa por CorrigeBase
'
'      Call InsertParamEmpBas(IdEmpresa, Ano)

'      Call AddLog("CrearNuevoAno: 6.3", 1)

   End If
      
#If DATACON = 1 Then       'Access
      
   If gEmprSeparadas Then
   '   Q1 = "DELETE * FROM ParamEmpresa WHERE Tipo = " & TPE_DBINFO
      Q1 = " WHERE Tipo = " & TPE_DBINFO
'      Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano   'no corre empresa año en empresas separadas
      Call DeleteSQL(DbMain, "ParamEmpresa", Q1)
   
      Call ChkDbInfo(DbMain, Rut, Ano, IdEmpresa)
         
      'cierro la nueva DB
      Call CloseDb(DbMain)
      
      If Not EmpVacia Then
         'compactamos la base de datos
         Call CompactDb(gDbPath & "\Empresas\" & Ano, RutMdb, False, True)
         DoEvents
      End If
      
      Set DbMain = DbActual
      
   End If
   
#End If

   CrearNuevoAno = True
   
   Call AddLog("CrearNuevoAno: FIN ", 1)
   
End Function
'Crea nuevo año en SQL Server desde Access, si no existe, copiando del año anterior si existe, o con DB vacía.
Public Function CrearNuevoAnoSQLFromAccess(ByVal IdEmpresa As Long, ByVal Ano As Integer, ByVal Rut As String, ByVal NombreEmpresa As String) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim RsDao As dao.Recordset
   Dim RsDaoAux As dao.Recordset
   Dim RutMdb As String
   Dim Msg As String
   Dim EmpVacia As Boolean
   Dim DbAnoAnt As Database
   Dim PathDbAnoAnt As String
   Dim CopyErr As Boolean
   Dim FCierre As Long
   Dim DbPath As String
   Dim Frm As Form
   Dim Rc As Integer
   Dim ConnStr As String
   Dim IdCompAperTrib As Long
   Dim NuevoAnoVacio As Boolean
   Dim Wh As String, Fld As String, Fld2 As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   Dim i As Integer
   Dim IdCta As Long
   Dim FldArray(11) As AdvTbAddNew_t
   Dim idcomp As Long, IdCompNew As Long
   Dim IdEmpAnoAnt As Long
   Dim IdCtaAportes As Long, IdCtaRetiros As Long
   Dim TblName As String, Where As String, OrderBy As String
   Dim IdGrupo As Long, IdGrupoNew As Long
   Dim DbVer As Integer
   Dim FrmSelBase As FrmSelRuta
   Dim FNBaseAccess As String
   Dim Idx As Integer
   Dim LastAccessPath As String
         

   
   EmpVacia = False

   CrearNuevoAnoSQLFromAccess = False

   If gDbType = SQL_ACCESS Then
      Exit Function
   End If

   Call AddLog("CrearNuevoAnoSQLFromAccess: 1", 1)

   'Veamos si ya existe nuevo año en base actual SQL Server
   Q1 = "SELECT Ano FROM EmpresasAno WHERE Ano=" & Ano & " AND idEmpresa=" & IdEmpresa
   Set Rs = OpenRs(DbMain, Q1)

   If Not Rs.EOF Then   'existe nuevo año en SQL Server
      Call CloseRs(Rs)
      CrearNuevoAnoSQLFromAccess = True
      Exit Function
   End If
   Call CloseRs(Rs)

   Call AddLog("CrearNuevoAnoSQLFromAccess: 2", 1)
      
   NuevoAnoVacio = False

   'Vemos si ya se hizo una creación de año desde Access antes. Si es así, el path debería estar en el IniFile
   LastAccessPath = GetIniString(gIniFile, "SQLFromAccess", "PathLPContabAccess", "")
   
   'veamos si existe archivo del año anterior en Access
   RutMdb = Rut & ".mdb"
   
   If LastAccessPath <> "" Then
      PathDbAnoAnt = LastAccessPath & "\Empresas\" & Ano - 1 & "\" & RutMdb
   Else
      PathDbAnoAnt = gDbPath & "\Empresas\" & Ano - 1 & "\" & RutMdb
   End If
   If Not ExistFile(PathDbAnoAnt) Then 'no existe archivo año anterior en Access

      'querrá seleccionar la Ruta del archivo?
      If MsgBox1("No existe año anterior en Access para esta empresa en la siguiente Ruta " & vbCrLf & vbCrLf & PathDbAnoAnt & vbCrLf & vbCrLf & "Desea seleccionar otra Ruta para el archivo Access?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
         
         If MsgBox1("No es posible abrir el nuevo año y generar saldos de apertura. Debe partir de una empresa vacía." & vbCrLf & vbCrLf & "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
            Exit Function

         Else
            NuevoAnoVacio = True
         
         End If
         
      Else
         'permitimos seleccionar la Ruta del archivo
         Set FrmSelBase = New FrmSelRuta
         
         If FrmSelBase.FSelFile("Seleccionar base de datos Access LPContab.mdb", "Archivos MDB (*.mdb)|*.mdb", "LPContab.mdb", FNBaseAccess) = vbCancel Then
            
            If MsgBox1("No es posible abrir el nuevo año y generar saldos de apertura. Debe partir de una empresa vacía." & vbCrLf & vbCrLf & "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
               Exit Function
   
            Else
               NuevoAnoVacio = True
            
            End If
            
         Else
            'seleccionada la ruta, veamos si existe
            LastAccessPath = FNBaseAccess
            Idx = InStrRev(FNBaseAccess, "\")
            If Idx > 0 Then
               PathDbAnoAnt = Left(FNBaseAccess, Idx) & "Empresas\" & Ano - 1 & "\" & RutMdb
            Else
               PathDbAnoAnt = ""
            End If
            
            If Not ExistFile(PathDbAnoAnt) = True Then 'no existe archivo año anterior en Access
               If MsgBox1("No es posible abrir el nuevo año y generar saldos de apertura. Debe partir de una empresa vacía." & vbCrLf & vbCrLf & "¿Desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
                  Exit Function
      
               Else
                  NuevoAnoVacio = True
               
               End If
            Else     'si existe, guardamos la ruta de la carpeta LPContab
               LastAccessPath = Left(FNBaseAccess, Idx - 1)
               Call SetIniString(gIniFile, "SQLFromAccess", "PathLPContabAccess", LastAccessPath)
            End If
            
         End If
         
      End If
         
   End If

   Call AddLog("CrearNuevoAnoSQLFromAccess: 3", 1)


  'No existe el nuevo año, lo creamos

   Msg = "ATENCIÓN! Para hacer esta operación nadie debe estar trabajando con la empresa " & NombreEmpresa & "." & vbNewLine & vbNewLine & "¿Desea continuar?"
   'Msg = Msg & " Verifique que haya sido creada la cuenta para almacenar el resultado del ejercicio. Esta será utilizada para realizar el proceso de apertura del año siguiente."
   If MsgBox1(Msg, vbQuestion Or vbDefaultButton1 Or vbYesNo) <> vbYes Then
      Exit Function
   End If

   Call AddLog("CrearNuevoAnoSQLFromAccess: 4", 1)

   On Error Resume Next

   'abrimos año anterior en Access

   If Not NuevoAnoVacio Then
     '2868088
      ConnStr = ";PWD=" & PASSW_PREFIX & Rut & ";"
      
      'ConnStr = ";PWD=" & PASSW_PREFIX_NEW & Rut & ";"
     '2868088
      Set DbAnoAnt = OpenDatabase(PathDbAnoAnt, False, False, ConnStr)

      If ERR <> 0 Or DbAnoAnt Is Nothing Then
         MsgBox1 "No fue posible abrir el año anterior en Access para esta empresa. (" & Error & ")" & vbCrLf & vbCrLf & "No es posible abrir el nuevo año generando los saldos de apertura. Debe partir de una empresa vacía.", vbExclamation
         Exit Function
      End If

      Call AddLog("CrearNuevoAno: 5.1", 1)

      'vemos si el año anterior está cerrado
      Q1 = "SELECT Id FROM Empresa "      'esta tabla tiene un solo registro
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)

      If RsDao.EOF = False Then
         IdEmpAnoAnt = vFldDao(RsDao("Id"))
      End If

      Call CloseRs(RsDao)
     
      'vemos si el año anterior está actaulizado con la última versión
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)

      If RsDao.EOF = False Then
         DbVer = Val(vFldDao(RsDao("Valor")))
      End If

      Call CloseRs(RsDao)
   
      If DbVer < 705 Then
         MsgBox1 "El año anterior en Access no está actualizado con la úlrtima versión del sistema LPContabilidad Access." & vbCrLf & vbCrLf & "Debe abrirlo con la versión actualizada de LPContabilidad Access 7.0, cerrarlo y luego, volver a intentar a crear el nuevo año en SQL Server.", vbExclamation + vbOKOnly
         CloseDb (DbAnoAnt)
         Exit Function
      End If
    
      'vemos si el año anterior está cerrado
      Q1 = "SELECT FCierre FROM EmpresasAno WHERE IdEmpresa=" & IdEmpAnoAnt & " AND Ano=" & Ano - 1
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)

      If RsDao.EOF = False Then
         FCierre = vFldDao(RsDao("FCierre"))
      End If

      Call CloseRs(RsDao)

      If FCierre = 0 Then
         MsgBox1 "El año anterior en Access aún no ha sido cerrado. No es posible abrir el nuevo año y generar saldos de apertura.", vbExclamation + vbOKOnly
         CloseDb (DbAnoAnt)
         Exit Function
      End If
      
      'veamos si existe el año que se desea crear en SQL Server. Si no existe, lo crea.
      Call CheckRcEmpAno(Ano, IdEmpresa)


      'copiamos las tablas que se duplican en el año siguiente, si es que hay datos


      'Copiamos las Cuentas del año anterior
      Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, IdCuenta, IdPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22, Atrib1, Atrib2, Atrib3, Atrib4, Atrib5, Atrib6, Atrib7, Atrib8, Atrib9, Atrib10, CodF29, CorrelativoCheque, CodIFRS_EstRes, CodIFRS_EstFin, DebeTrib, HaberTrib, CodIFRS, CodF22_14Ter, TipoPartida, CodCtaPlanSII "
      Fld2 = " IdEmpresa, Ano, IdCuentaOld, IdPadreOld, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22, Atrib1, Atrib2, Atrib3, Atrib4, Atrib5, Atrib6, Atrib7, Atrib8, Atrib9, Atrib10, CodF29, CorrelativoCheque, CodIFRS_EstRes, CodIFRS_EstFin, DebeTrib, HaberTrib, CodIFRS, CodF22_14Ter, TipoPartida, CodCtaPlanSII"
      Q1 = "SELECT " & Fld & " FROM Cuentas "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano - 1 ' dado que el año anterior es Access y empresas separadas
      Q1 = Q1 & " ORDER BY Cuentas.IdCuenta"
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)
      
      Do While Not RsDao.EOF
         
         Q1 = "INSERT INTO Cuentas (" & Fld2 & ") VALUES("
         For i = 0 To RsDao.Fields.Count - 1
            If RsDao(i).Type = dbText Or RsDao(i).Type = dbMemo Or RsDao(i).Type = dbChar Then
               Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDao(i))) & "',"
            Else
               Q1 = Q1 & vFldDao(RsDao(i)) & ","
            End If
            
         Next i
         
         Q1 = Left(Q1, Len(Q1) - 1) & ")"
         Call ExecSQL(DbMain, Q1)
         
         RsDao.MoveNext
     
      Loop
     
      Call CloseRs(RsDao)
      
      'hilamos las cuentas con sus padres ya que se nos movieron los ids por el Identity
      
      Tbl = " Cuentas "
      sFrom = " Cuentas INNER JOIN Cuentas As Cuentas1 ON Cuentas.IdPadreOld = Cuentas1.IdCuentaOld  "
      sFrom = sFrom & " AND Cuentas.IdEmpresa = Cuentas1.IdEmpresa "
      sFrom = sFrom & " AND Cuentas.Ano = Cuentas1.Ano "
      sSet = " Cuentas.IdPadre = Cuentas1.IdCuenta "
      sWhere = " WHERE Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
      Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)

      Tbl = " Cuentas "
      sFrom = " Cuentas "
      sSet = " Cuentas.IdPadre = 0 "
      sWhere = " WHERE Cuentas.Nivel = 1 AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
      Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)

      

      'Copiamos CuentasBásicas
      Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, Tipo, TipoLib, TipoValor, CuentasBasicas.IdCuenta, Codigo "
      Fld2 = " IdEmpresa, Ano, Tipo, TipoLib, TipoValor, IdCuenta "
      Q1 = "SELECT " & Fld & " FROM CuentasBasicas "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano - 1 ' dado que el año anterior es Access y empresas separadas
      Q1 = Q1 & " INNER JOIN Cuentas ON CuentasBasicas.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & " ORDER BY Tipo, TipoLib, TipoValor "
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)
      
      Do While Not RsDao.EOF
      
         IdCta = 0
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & RsDao("Codigo") & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCta = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
         
         Q1 = "INSERT INTO CuentasBasicas (" & Fld2 & ") VALUES("
         
         For i = 0 To RsDao.Fields.Count - 1
            If RsDao(i).Name = "IdCuenta" Then
               Q1 = Q1 & IdCta
               Exit For
            ElseIf RsDao(i).Type = dbText Or RsDao(i).Type = dbMemo Or RsDao(i).Type = dbChar Then
               Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDao(i))) & "',"
            Else
               Q1 = Q1 & vFldDao(RsDao(i)) & ","
            End If
            
         Next i
         
         Q1 = Q1 & ")"
         Call ExecSQL(DbMain, Q1)
         
         RsDao.MoveNext
     
      Loop
      
      Call CloseRs(RsDao)
      
      'Copiamos ImpAdic
      Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, TipoLib, TipoValor, Tasa, EsRecuperable, CodCuenta, IdCuenta"
      Fld2 = " IdEmpresa, Ano, TipoLib, TipoValor, Tasa, EsRecuperable, CodCuenta, IdCuenta "

      Q1 = "SELECT " & Fld & " FROM ImpAdic "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano - 1 ' dado que el año anterior es Access y empresas separadas
      Q1 = Q1 & " ORDER BY IdImpAdic "
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)
      
      Do While Not RsDao.EOF
      
         IdCta = 0
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & RsDao("CodCuenta") & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCta = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
         
         Q1 = "INSERT INTO ImpAdic (" & Fld2 & ") VALUES("
         
         For i = 0 To RsDao.Fields.Count - 1
            If RsDao(i).Name = "IdCuenta" Then
               Q1 = Q1 & IdCta
               Exit For
            ElseIf RsDao(i).Type = dbText Or RsDao(i).Type = dbMemo Or RsDao(i).Type = dbChar Then
               Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDao(i))) & "',"
            Else
               Q1 = Q1 & vFldDao(RsDao(i)) & ","
            End If
           
         Next i
         
         Q1 = Q1 & ")"
         Call ExecSQL(DbMain, Q1)
         
         RsDao.MoveNext
     
      Loop
      
      Call CloseRs(RsDao)

      'Copiamos CtasAjustesExCont
      Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, TipoAjuste, IdItem, CodCuenta, IdCuenta"
      Fld2 = " IdEmpresa, Ano, TipoAjuste, IdItem, CodCuenta, IdCuenta "

      Q1 = "SELECT " & Fld & " FROM CtasAjustesExCont "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano - 1 ' dado que el año anterior es Access y empresas separadas
      Q1 = Q1 & " ORDER BY IdCtaAjustes "
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)
      
      Do While Not RsDao.EOF
      
         IdCta = 0
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & RsDao("CodCuenta") & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCta = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
         
         Q1 = "INSERT INTO CtasAjustesExCont (" & Fld2 & ") VALUES("
         
         For i = 0 To RsDao.Fields.Count - 1
            If RsDao(i).Name = "IdCuenta" Then
               Q1 = Q1 & IdCta
               Exit For
            ElseIf RsDao(i).Type = dbText Or RsDao(i).Type = dbMemo Or RsDao(i).Type = dbChar Then
               Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDao(i))) & "',"
            Else
               Q1 = Q1 & vFldDao(RsDao(i)) & ","
            End If
           
         Next i
         
         Q1 = Q1 & ")"
         Call ExecSQL(DbMain, Q1)
         
         RsDao.MoveNext
     
      Loop
      
      Call CloseRs(RsDao)

      'Copiamos CtasAjustesExContRLI
      Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, TipoAjuste, IdGrupo, IdItem, CodCuenta, IdCuenta"
      Fld2 = " IdEmpresa, Ano, TipoAjuste, IdGrupo, IdItem, CodCuenta, IdCuenta "

      Q1 = "SELECT " & Fld & " FROM CtasAjustesExContRLI "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano - 1 ' dado que el año anterior es Access y empresas separadas
      Q1 = Q1 & " ORDER BY IdCtaAjustesRLI "
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)
      
      Do While Not RsDao.EOF
      
         IdCta = 0
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & RsDao("CodCuenta") & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCta = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
         
         Q1 = "INSERT INTO CtasAjustesExContRLI (" & Fld2 & ") VALUES("
         
         For i = 0 To RsDao.Fields.Count - 1
            If RsDao(i).Name = "IdCuenta" Then
               Q1 = Q1 & IdCta
               Exit For
            ElseIf RsDao(i).Type = dbText Or RsDao(i).Type = dbMemo Or RsDao(i).Type = dbChar Then
               Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDao(i))) & "',"
            Else
               Q1 = Q1 & vFldDao(RsDao(i)) & ","
            End If
           
         Next i
         
         Q1 = Q1 & ")"
         Call ExecSQL(DbMain, Q1)
         
         RsDao.MoveNext
     
      Loop
      
      Call CloseRs(RsDao)

      'Socios
      Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, RUT, Socios.Nombre, PjePart, MontoSuscrito, MontoPagado, IdTipoSocio, Vigente, IdCuentaAportes, IdCuentaRetiros"
      Fld2 = " IdEmpresa, Ano, RUT, Nombre, PjePart, MontoSuscrito, MontoPagado, IdTipoSocio, Vigente, IdCuentaAportes, IdCuentaRetiros "

      Q1 = "SELECT " & Fld & ", Cuentas.Codigo as CodCtaAportes, Cuentas1.Codigo as CodCtaRetiros FROM (Socios "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano - 1 ' dado que el año anterior es Access y empresas separadas
      Q1 = Q1 & " INNER JOIN Cuentas ON Socios.IdCuentaAportes = Cuentas.IdCuenta) "
      Q1 = Q1 & " INNER JOIN Cuentas as Cuentas1 ON Socios.IdCuentaRetiros = Cuentas1.IdCuenta "
      Q1 = Q1 & " ORDER BY IdSocio "
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)
      
      Do While Not RsDao.EOF
      
         IdCtaAportes = 0
         IdCtaRetiros = 0
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & RsDao("CodCtaAportes") & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCtaAportes = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
         
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & RsDao("CodCtaRetiros") & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCtaRetiros = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
         
         Q1 = "INSERT INTO Socios (" & Fld2 & ") VALUES("
         
         For i = 0 To RsDao.Fields.Count - 1
            If RsDao(i).Name = "IdCuentaAportes" Then
               Q1 = Q1 & IdCtaAportes & ","
            ElseIf RsDao(i).Name = "IdCuentaRetiros" Then
               Q1 = Q1 & IdCtaRetiros & ","
               Exit For
            ElseIf RsDao(i).Type = dbText Or RsDao(i).Type = dbMemo Or RsDao(i).Type = dbChar Then
               Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDao(i))) & "',"
            Else
               Q1 = Q1 & vFldDao(RsDao(i)) & ","
            End If
           
         Next i
         
         Q1 = Left(Q1, Len(Q1) - 1) & ")"
         Call ExecSQL(DbMain, Q1)
        
         RsDao.MoveNext
     
      Loop
      
      Call CloseRs(RsDao)


      'Empresa
      Fld = Ano & " As Ano, " & IdEmpresa & " As id, Rut, NombreCorto, RazonSocial, ApPaterno, ApMaterno, Nombre, Calle, Numero, Dpto, Telefonos, Fax, Region, Comuna, Ciudad, Giro, ActEconom, CodActEconom, DomPostal, ComunaPostal, Email, Web, FechaConstitucion, FechaInicioAct, RepConjunta, RutRepLegal1, RepLegal1, RutRepLegal2, RepLegal2, Contador, RutContador, TipoContrib, TransaBolsa, Franq14bis, FranqLey18392, FranqDL600, FranqDL701, FranqDS341, Opciones, TContribFUT, Franq14ter, Franq14quater, ObligaLibComprasVentas, FranqRentaAtribuida, FranqSemiIntegrado, Franq14ASemiIntegrado, FranqProPymeTransp, FranqProPymeGeneral "
      Fld2 = " Ano, id, Rut, NombreCorto, RazonSocial, ApPaterno, ApMaterno, Nombre, Calle, Numero, Dpto, Telefonos, Fax, Region, Comuna, Ciudad, Giro, ActEconom, CodActEconom, DomPostal, ComunaPostal, Email, Web, FechaConstitucion, FechaInicioAct, RepConjunta, RutRepLegal1, RepLegal1, RutRepLegal2, RepLegal2, Contador, RutContador, TipoContrib, TransaBolsa, Franq14bis, FranqLey18392, FranqDL600, FranqDL701, FranqDS341, Opciones, TContribFUT, Franq14ter, Franq14quater, ObligaLibComprasVentas, FranqRentaAtribuida, FranqSemiIntegrado, Franq14ASemiIntegrado, FranqProPymeTransp, FranqProPymeGeneral  "

      Q1 = "SELECT " & Fld & " FROM Empresa "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano - 1 ' dado que el año anterior es Access y empresas separadas
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)
      
      If Not RsDao.EOF Then
      
         Q1 = "INSERT INTO Empresa (" & Fld2 & ") VALUES("
         
         For i = 0 To RsDao.Fields.Count - 1
            If RsDao(i).Type = dbText Or RsDao(i).Type = dbMemo Or RsDao(i).Type = dbChar Then
               Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDao(i))) & "',"
            Else
               Q1 = Q1 & vFldDao(RsDao(i)) & ","
            End If
            
         Next i
         
         Q1 = Left(Q1, Len(Q1) - 1) & ")"
         Call ExecSQL(DbMain, Q1, False)
              
      End If
      
      Call CloseRs(RsDao)
      
      
      'ParamEmpresa
      'eliminamos los default que pueden haber
      Q1 = "DELETE FROM ParamEmpresa WHERE IdEmpresa =" & IdEmpresa & " AND Ano = " & Ano
      Call ExecSQL(DbMain, Q1)
      
      'Ahora copiamos los registros
      Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, Tipo, ParamEmpresa.Codigo, Valor "
      Fld2 = " IdEmpresa, Ano, Tipo, Codigo, Valor "

      Q1 = "SELECT " & Fld & ", Cuentas.Codigo As CodCuenta FROM ParamEmpresa "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano - 1 ' dado que el año anterior es Access y empresas separadas
      Q1 = Q1 & " LEFT JOIN Cuentas ON Cuentas.IdCuenta = Val(ParamEmpresa.Valor)"
      Q1 = Q1 & " ORDER BY Tipo, ParamEmpresa.Codigo "
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)
      
      Do While Not RsDao.EOF
      
         IdCta = 0
         
         If Left(vFldDao(RsDao("Tipo")), 3) = "CTA" And vFldDao(RsDao("CodCuenta")) <> "" Then
         
            Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & RsDao("CodCuenta") & "'"
            Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               IdCta = vFld(Rs("IdCuenta"))
            End If
            Call CloseRs(Rs)
         
         End If
         
         Q1 = "INSERT INTO ParamEmpresa (" & Fld2 & ") VALUES("
         
         For i = 0 To RsDao.Fields.Count - 1
            If RsDao(i).Name = "Valor" Then
               If IdCta > 0 Then
                  Q1 = Q1 & "'" & IdCta & "',"
               Else
                  Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDao(i))) & "',"
               End If
               
               Exit For
            Else
               If RsDao(i).Type = dbText Or RsDao(i).Type = dbMemo Or RsDao(i).Type = dbChar Then
                  Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDao(i))) & "',"
               Else
                  Q1 = Q1 & vFldDao(RsDao(i)) & ","
               End If
            End If
            
         Next i
         
         Q1 = Left(Q1, Len(Q1) - 1) & ")"
         Call ExecSQL(DbMain, Q1)
         
         RsDao.MoveNext
     
      Loop
      
      Call CloseRs(RsDao)
      
      
      
      'comprobantes tipo
      
      Fld = IdEmpresa & " As IdEmpresa, Correlativo, Nombre, Descrip, Fecha, Tipo, Estado, Glosa, TotalDebe, TotalHaber, IdUsuario"
      Fld2 = " IdEmpresa, Correlativo, Nombre, Descrip, Fecha, Tipo, Estado, Glosa, TotalDebe, TotalHaber, IdUsuario "
   
      Q1 = "SELECT IdComp, " & Fld & " FROM CT_Comprobante "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa  ' dado que el año anterior es Access y empresas separadas
      Q1 = Q1 & " ORDER BY Tipo, Nombre "
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)
      
      Do While Not RsDao.EOF
      
         idcomp = vFldDao(RsDao("IdComp"))
                     
         FldArray(0).FldName = "IdEmpresa"
         FldArray(0).FldValue = IdEmpresa
         FldArray(0).FldIsNum = True
                     
         FldArray(1).FldName = "Correlativo"
         FldArray(1).FldValue = vFldDao(RsDao("Correlativo"))
         FldArray(1).FldIsNum = True
         
         FldArray(2).FldName = "Nombre"
         FldArray(2).FldValue = vFldDao(RsDao("Nombre"))
         FldArray(2).FldIsNum = False
               
         FldArray(3).FldName = "Descrip"
         FldArray(3).FldValue = vFldDao(RsDao("Descrip"))
         FldArray(3).FldIsNum = False
               
         FldArray(4).FldName = "Fecha"
         FldArray(4).FldValue = vFldDao(RsDao("Fecha"))
         FldArray(4).FldIsNum = True
         
         FldArray(5).FldName = "Tipo"
         FldArray(5).FldValue = vFldDao(RsDao("Tipo"))
         FldArray(5).FldIsNum = True
         
         FldArray(6).FldName = "Estado"
         FldArray(6).FldValue = vFldDao(RsDao("Estado"))
         FldArray(6).FldIsNum = True
         
         FldArray(7).FldName = "Glosa"
         FldArray(7).FldValue = vFldDao(RsDao("Glosa"))
         FldArray(7).FldIsNum = False
                  
         FldArray(8).FldName = "TotalDebe"
         FldArray(8).FldValue = 0
         FldArray(8).FldIsNum = True
         
         FldArray(9).FldName = "TotalHaber"
         FldArray(9).FldValue = 0
         FldArray(9).FldIsNum = True
         
         FldArray(10).FldName = "IdUsuario"
         FldArray(10).FldValue = gUsuario.IdUsuario
         FldArray(10).FldIsNum = True
         
         IdCompNew = AdvTbAddNewMult(DbMain, "CT_Comprobante", "IdComp", FldArray)
      
         Fld = IdEmpresa & " As IdEmpresa, Orden, CodCuenta, Debe, Haber, Glosa, IdCCosto, IdAreaNeg, Conciliado, 0 as IdCuenta "
         Fld2 = " IdEmpresa, Orden, CodCuenta, Debe, Haber, Glosa, IdCCosto, IdAreaNeg, Conciliado, IdCuenta "
         
         Q1 = "SELECT " & Fld & " FROM CT_MovComprobante "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa  ' dado que el año anterior es Access y empresas separadas
         Q1 = Q1 & " WHERE IdComp = " & idcomp
         Q1 = Q1 & " ORDER BY IdMov "
         Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
         
         Do While Not RsDaoAux.EOF
                  
            IdCta = 0
            Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & RsDaoAux("CodCuenta") & "'"
            Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               IdCta = vFld(Rs("IdCuenta"))
            End If
            Call CloseRs(Rs)
                        
            Q1 = "INSERT INTO CT_MovComprobante ( IdComp, " & Fld2 & ") VALUES(" & IdCompNew & ","
            
            For i = 0 To RsDaoAux.Fields.Count - 1
               If RsDaoAux(i).Name = "IdCuenta" Then
                  Q1 = Q1 & "'" & IdCta & "',"
                  Exit For
               ElseIf RsDaoAux(i).Type = dbText Or RsDaoAux(i).Type = dbMemo Or RsDaoAux(i).Type = dbChar Then
                  Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDaoAux(i))) & "',"
               Else
                  Q1 = Q1 & vFldDao(RsDaoAux(i)) & ","
               End If
               
            Next i
            
            Q1 = Left(Q1, Len(Q1) - 1) & ")"
            Call ExecSQL(DbMain, Q1)
            
            RsDaoAux.MoveNext
         
         Loop
         
         Call CloseRs(RsDaoAux)
         
         RsDao.MoveNext
     
      Loop
      
      Call CloseRs(RsDao)
       
      Call UpdateComprobantesTipo
      
      
      'ahora copiamos los otros datos así como están
      
      'AreaNegocio
      TblName = "AreaNegocio"
      Fld = IdEmpresa & " As IdEmpresa, Codigo, Descripcion, Vigente "
      Fld2 = " IdEmpresa, Codigo, Descripcion, Vigente "
      
      Where = " WHERE Vigente <> 0 "
      OrderBy = " ORDER BY Codigo "
      
      Call CopyTblSimpleFromAccessToSQLServer(DbAnoAnt, IdEmpresa, Ano, TblName, Fld, Fld2, Where, OrderBy)
      
      'CentroCosto
      TblName = "CentroCosto"
      Fld = IdEmpresa & " As IdEmpresa, Codigo, Descripcion, Vigente"
      Fld2 = " IdEmpresa, Codigo, Descripcion, Vigente"
      
      Where = " WHERE Vigente <> 0 "
      OrderBy = " ORDER BY Codigo "
      
      Call CopyTblSimpleFromAccessToSQLServer(DbAnoAnt, IdEmpresa, Ano, TblName, Fld, Fld2, Where, OrderBy)
            
      'Entidades
      TblName = "Entidades"
      Fld = IdEmpresa & " As IdEmpresa, Rut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, ActEcon, CodActEcon, DomPostal, ComPostal, email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, Giro, NotValidRut, EsSupermercado, EntRelacionada, CodCtaAfecto, CodCtaExento, CodCtaTotal, PropIVA, CodCCostoAfecto, CodAreaNegAfecto, CodCCostoExento, CodAreaNegExento, CodCCostoTotal, CodAreaNegTotal, CodCtaAfectoVta, CodCtaExentoVta, CodCtaTotalVta, CodCCostoAfectoVta, CodAreaNegAfectoVta, CodCCostoExentoVta, CodAreaNegExentoVta, CodCCostoTotalVta, CodAreaNegTotalVta, EsDelGiro"
      Fld2 = " IdEmpresa, Rut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, ActEcon, CodActEcon, DomPostal, ComPostal, email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, Giro, NotValidRut, EsSupermercado, EntRelacionada, CodCtaAfecto, CodCtaExento, CodCtaTotal, PropIVA, CodCCostoAfecto, CodAreaNegAfecto, CodCCostoExento, CodAreaNegExento, CodCCostoTotal, CodAreaNegTotal, CodCtaAfectoVta, CodCtaExentoVta, CodCtaTotalVta, CodCCostoAfectoVta, CodAreaNegAfectoVta, CodCCostoExentoVta, CodAreaNegExentoVta, CodCCostoTotalVta, CodAreaNegTotalVta, EsDelGiro"
      
      Where = " WHERE Estado = " & EE_ACTIVO
      OrderBy = " ORDER BY Rut "
      
      Call CopyTblSimpleFromAccessToSQLServer(DbAnoAnt, IdEmpresa, Ano, TblName, Fld, Fld2, Where, OrderBy)
      
      
      
       'AFGrupos y AFComponentes
      
      Fld = IdEmpresa & " As IdEmpresa, NombGrupo "
      Fld2 = " IdEmpresa, NombGrupo "
   
      Q1 = "SELECT IdGrupo, " & Fld & " FROM AFGrupos "
      Q1 = Q1 & " ORDER BY NombGrupo "
      Set RsDao = OpenRsDao(DbAnoAnt, Q1)
      
      For i = 0 To UBound(FldArray)
         FldArray(i).FldName = ""
         FldArray(i).FldValue = 0
         FldArray(i).FldIsNum = False
      Next i
      
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'por si ya existe el grupo y entrega llave duplicada FCA 10 dec 2019

      Do While Not RsDao.EOF
      
         IdGrupo = vFldDao(RsDao("IdGrupo"))
                     
         FldArray(0).FldName = "IdEmpresa"
         FldArray(0).FldValue = IdEmpresa
         FldArray(0).FldIsNum = True
                              
         FldArray(1).FldName = "NombGrupo"
         FldArray(1).FldValue = vFldDao(RsDao("NombGrupo"))
         FldArray(1).FldIsNum = False
                        
         IdGrupoNew = AdvTbAddNewMult(DbMain, "AFGrupos", "IdGrupo", FldArray, False)
         
         Fld = IdEmpresa & " As IdEmpresa, NombComp"
         Fld2 = " IdEmpresa, NombComp "
         
         Q1 = "SELECT " & Fld & " FROM AFComponentes "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa  ' dado que el año anterior es Access y empresas separadas
         Q1 = Q1 & " WHERE IdGrupo = " & IdGrupo
         Q1 = Q1 & " ORDER BY IdComp "
         Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
         
         Do While Not RsDaoAux.EOF
                                          
            Q1 = "INSERT INTO AFComponentes ( IdGrupo, " & Fld2 & ") VALUES(" & IdGrupoNew & ","
            
            For i = 0 To RsDaoAux.Fields.Count - 1
               If RsDaoAux(i).Type = dbText Or RsDaoAux(i).Type = dbMemo Or RsDaoAux(i).Type = dbChar Then
                  Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDaoAux(i))) & "',"
               Else
                  Q1 = Q1 & vFldDao(RsDaoAux(i)) & ","
               End If
               
            Next i
            
            Q1 = Left(Q1, Len(Q1) - 1) & ")"
            Call ExecSQL(DbMain, Q1, False)
            
            RsDaoAux.MoveNext
         
         Loop
         
         Call CloseRs(RsDaoAux)
         
         RsDao.MoveNext
     
      Loop
      
      Call CloseRs(RsDao)
     
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")   'por si ya existe el grupo y entrega llave duplicada FCA 10 dec 2019
      
      'ahora obtenemos los documentos centralizados y pagados con saldo pendiente desde el año anterior
     
      Call CopyDocsFromAccessToSQLServer(DbAnoAnt, IdEmpresa, Ano)
      
      'Luego los activos fijos con valor libro mayor que cero o no depreciables del año anteriro
      
      Call CopyActFijoFromAccessToSQLServer(DbAnoAnt, IdEmpresa, Ano)
      
      'finalmente generamos los saldos de apertura en el plan de cuentas
      
      Call GenSaldosAperturaAccessFromSQLServer(DbAnoAnt, IdEmpresa, Ano)
            
      EmpVacia = False

   Else
      'veamos si existe el año que se desea crear en SQL Server. Si no existe, lo crea.
      Call CheckRcEmpAno(Ano, IdEmpresa)
      EmpVacia = True

   End If


   Q1 = "SELECT Count(*) as N FROM CT_Comprobante WHERE IdEmpresa = " & IdEmpresa
   Set Rs = OpenRs(DbMain, Q1)

   If Not Rs.EOF Then

      If vFld(Rs("N")) = 0 Then    'no tiene comp tipo para esta emnpresa, agregamos los base

         Call ResetCompTipoEmpJuntas(IdEmpresa)

      End If
   End If

   Call CloseRs(Rs)

   

   Call AddLog("CrearNuevoAno: 6", 1)

   'Actualizo la Tabla que tiene los años que tiene datos la empresa (ya está ingreesado más arriba
'   Q1 = "INSERT INTO EmpresasAno (IdEmpresa, Ano, FApertura) VALUES ("
'   Q1 = Q1 & IdEmpresa & "," & Ano & "," & CLng(Int(Now)) & ")"
'   Call ExecSQL(DbMain, Q1)

   'seteamos la indicacióon que tiene historia con Access
   If Not EmpVacia Then
      Wh = " WHERE Tipo = 'INITAÑO'"
      Wh = Wh & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
      
      Q1 = "SELECT * FROM ParamEmpresa " & Wh
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF Then
         Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano)"
         Q1 = Q1 & " VALUES( 'INITAÑO', 1, 'EMPHISTACC'," & IdEmpresa & "," & Ano & ")"   'Codigo = 1 indica que debe generar comp apertura
      Else
         Q1 = "UPDATE ParamEmpresa SET Codigo = 1, Valor = 'EMPHISTACC'" & Wh
      End If
      
      Call ExecSQL(DbMain, Q1)
      Call CloseRs(Rs)
   End If
   
   CloseDb (DbAnoAnt)

   CrearNuevoAnoSQLFromAccess = True

   Call AddLog("CrearNuevoAnoSQLFromAccess: FIN ", 1)

End Function
Private Function CopyTblSimpleFromAccessToSQLServer(ByVal DbAnoAnt As Database, ByVal IdEmpresa As Long, ByVal Ano As Long, ByVal TblName As String, ByVal Fld As String, ByVal Fld2 As String, Optional ByVal Where As String = "", Optional ByVal OrderBy As String = "")
   Dim Q1 As String
   Dim Rs As Recordset
   Dim RsDao As dao.Recordset
   Dim RsDaoAux As dao.Recordset
   Dim i As Integer

   'Copiamos TblName
   Q1 = "SELECT " & Fld & " FROM " & TblName
   Q1 = Q1 & Where
   Q1 = Q1 & OrderBy

   Set RsDao = OpenRsDao(DbAnoAnt, Q1)
    
   Do While Not RsDao.EOF
    
'       IdCta = 0
'       Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & RsDao("CodCuenta") & "'"
'       Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
'       Set Rs = OpenRs(DbMain, Q1)
'       If Not Rs.EOF Then
'          IdCta = vFld(Rs("IdCuenta"))
'       End If
'       Call CloseRs(Rs)
       
         Q1 = "INSERT INTO " & TblName & " (" & Fld2 & ") VALUES("
         
         For i = 0 To RsDao.Fields.Count - 1
'          If RsDao(i).Name = "IdCuenta" Then
'             Q1 = Q1 & IdCta
'             Exit For
'          Else
            If RsDao(i).Type = dbText Or RsDao(i).Type = dbMemo Or RsDao(i).Type = dbChar Then
               Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDao(i))) & "',"
            Else
               Q1 = Q1 & vFldDao(RsDao(i)) & ","
            End If
           
         Next i
       
         Q1 = Left(Q1, Len(Q1) - 1) & ")"
         Call ExecSQL(DbMain, Q1, False)
         
         RsDao.MoveNext
      
      Loop
      
      Call CloseRs(RsDao)


End Function
Public Function CopyDocsFromAccessToSQLServer(ByVal DbAnoAnt As Database, ByVal IdEmpresa As Long, ByVal Ano As Long)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim RsDao As dao.Recordset
   Dim RsDaoAux As dao.Recordset, RsDaoAux2 As dao.Recordset
   Dim i As Integer
   Dim Where As String
   Dim IdCta As Long, IdSuc As Long, IdEntidad As Long
   Dim FldArray(43) As AdvTbAddNew_t
   Dim IdDocNew As Long
   Dim CodCuenta As String, CodSuc As String, RutEntidad As String
   Dim IdDoc As Long
   Dim Fld As String, Fld2 As String
   Dim IdAreaNeg As Long, IdCCosto As Long
   Dim CodAreaNeg As String
   Dim CodCCosto As String
  
   'marcamos los que vamos a exportar con -1
   Q1 = "UPDATE Documento SET FExported = -1"
'   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) "    'para que siempre pueda volver a importar por si tiene que volver a hacerlo
   Q1 = Q1 & " WHERE SaldoDoc <> 0 AND Estado IN(" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
   Call ExecSQLDao(DbAnoAnt, Q1)   'IdEmpresa y Año no se requiere, dado que son empresas/Año separados

   Where = " WHERE FExported < 0"
   
   'Agregamos docs marcados para exportar
   'Calculamos Total Pagado Año Anterior para que tengamos los saldos OK
   Q1 = " SELECT IdDoc, IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Documento.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
   Q1 = Q1 & " iif(SaldoDoc IS NULL or SaldoDoc = 0, 0, iif(SaldoDoc > 0, Total - abs(SaldoDoc), -1 *(Total - abs(SaldoDoc))))  As TotPagadoAnoAnt, "
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, "
   Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld "
   Q1 = Q1 & " FROM ((Documento "
   Q1 = Q1 & " LEFT JOIN Cuentas ON Documento.IdCuentaAfecto = Cuentas.IdCuenta ) "
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta ) "
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaTotal = Cuentas2.IdCuenta "
   Q1 = Q1 & Where
   Set RsDao = OpenRsDao(DbAnoAnt, Q1)
   
   Do While Not RsDao.EOF
       
      IdDoc = vFldDao(RsDao("IdDoc"))
                  
      FldArray(0).FldName = "IdEmpresa"
      FldArray(0).FldValue = IdEmpresa
      FldArray(0).FldIsNum = True
                  
      FldArray(1).FldName = "Ano"
      FldArray(1).FldValue = Ano        ' hay que almacenar el año al que corresponde el documento en la DB no el año del documento mismo. Esto no: Year(vFldDao(RsDao("FEmision")))
      FldArray(1).FldIsNum = True
      
      FldArray(2).FldName = "TipoLib"
      FldArray(2).FldValue = vFldDao(RsDao("TipoLib"))
      FldArray(2).FldIsNum = True
      
      FldArray(3).FldName = "TipoDoc"
      FldArray(3).FldValue = vFldDao(RsDao("TipoDoc"))
      FldArray(3).FldIsNum = True
      
      FldArray(4).FldName = "NumDoc"
      FldArray(4).FldValue = vFldDao(RsDao("NumDoc"))
      FldArray(4).FldIsNum = False
            
      FldArray(5).FldName = "NumDocHasta"
      FldArray(5).FldValue = vFldDao(RsDao("NumDocHasta"))
      FldArray(5).FldIsNum = False
            
      FldArray(6).FldName = "Giro"
      FldArray(6).FldValue = Abs(vFldDao(RsDao("Giro")))
      FldArray(6).FldIsNum = True
            
         IdEntidad = 0
         RutEntidad = ""
         Q1 = "SELECT Rut FROM Entidades WHERE Entidades.IdEntidad = " & vFldDao(RsDao("IdEntidad"))
         Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
         If Not RsDaoAux.EOF Then
            RutEntidad = vFldDao(RsDaoAux("Rut"))
         End If
         Call CloseRs(RsDaoAux)
         
         Q1 = "SELECT IdEntidad FROM Entidades WHERE Entidades.Rut = '" & RutEntidad & "'"
         Q1 = Q1 & " AND Entidades.IdEmpresa = " & IdEmpresa
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdEntidad = vFld(Rs("IdEntidad"))
         End If
         Call CloseRs(Rs)
         
            
      FldArray(7).FldName = "IdEntidad"
      FldArray(7).FldValue = IdEntidad
      FldArray(7).FldIsNum = True
            
      FldArray(8).FldName = "TipoEntidad"
      FldArray(8).FldValue = vFldDao(RsDao("TipoEntidad"))
      FldArray(8).FldIsNum = True
            
      FldArray(9).FldName = "RutEntidad"
      FldArray(9).FldValue = vFldDao(RsDao("RutEntidad"))
      FldArray(9).FldIsNum = False
            
      FldArray(10).FldName = "NombreEntidad"
      FldArray(10).FldValue = vFldDao(RsDao("NombreEntidad"))
      FldArray(10).FldIsNum = False
            
      FldArray(11).FldName = "FEmision"
      FldArray(11).FldValue = vFldDao(RsDao("FEmision"))
      FldArray(11).FldIsNum = True
      
      FldArray(12).FldName = "FVenc"
      FldArray(12).FldValue = vFldDao(RsDao("FVenc"))
      FldArray(12).FldIsNum = True
      
      FldArray(13).FldName = "Descrip"
      FldArray(13).FldValue = vFldDao(RsDao("Descrip"))
      FldArray(13).FldIsNum = False
               
      FldArray(14).FldName = "Estado"
      FldArray(14).FldValue = vFldDao(RsDao("Estado"))
      FldArray(14).FldIsNum = True
      
      FldArray(15).FldName = "Exento"
      FldArray(15).FldValue = vFldDao(RsDao("Exento"))
      FldArray(15).FldIsNum = True
      
         IdCta = 0
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & vFldDao(RsDao("CodCtaExentoOld")) & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCta = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
      
      FldArray(16).FldName = "IdCuentaExento"
      FldArray(16).FldValue = IdCta
      FldArray(16).FldIsNum = True
      
      FldArray(17).FldName = "Afecto"
      FldArray(17).FldValue = vFldDao(RsDao("Afecto"))
      FldArray(17).FldIsNum = True
      
         IdCta = 0
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & vFldDao(RsDao("CodCtaAfectoOld")) & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCta = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
         
      FldArray(18).FldName = "IdCuentaAfecto"
      FldArray(18).FldValue = IdCta
      FldArray(18).FldIsNum = True
      
      FldArray(19).FldName = "IVA"
      FldArray(19).FldValue = vFldDao(RsDao("IVA"))
      FldArray(19).FldIsNum = True
      
      IdCta = 0
      If vFldDao(RsDao("IVA")) <> 0 Then
      
         CodCuenta = ""
         Q1 = "SELECT Codigo FROM Cuentas WHERE Cuentas.IdCuenta = " & vFldDao(RsDao("IdCuentaIVA"))
         Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
         If Not RsDaoAux.EOF Then
            CodCuenta = vFldDao(RsDaoAux("Codigo"))
         End If
         Call CloseRs(RsDaoAux)
         
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & CodCuenta & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCta = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
               
      End If
      FldArray(20).FldName = "IdCuentaIVA"
      FldArray(20).FldValue = IdCta
      FldArray(20).FldIsNum = True
      
      FldArray(21).FldName = "OtroImp"
      FldArray(21).FldValue = vFldDao(RsDao("OtroImp"))
      FldArray(21).FldIsNum = True
      
      IdCta = 0
      If vFldDao(RsDao("OtroImp")) <> 0 Then
      
         CodCuenta = ""
         If vFldDao(RsDao("IdCuentaOtroImp")) <> 0 Then
            Q1 = "SELECT Codigo FROM Cuentas WHERE Cuentas.IdCuenta = " & vFldDao(RsDao("IdCuentaOtroImp"))
            Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
            If Not RsDaoAux.EOF Then
               CodCuenta = vFldDao(RsDaoAux("Codigo"))
            End If
            Call CloseRs(RsDaoAux)
            
            Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & CodCuenta & "'"
            Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               IdCta = vFld(Rs("IdCuenta"))
            End If
            Call CloseRs(Rs)
         End If
                  
      End If
      FldArray(22).FldName = "IdCuentaOtroImp"
      FldArray(22).FldValue = IdCta
      FldArray(22).FldIsNum = True
      
      FldArray(23).FldName = "Total"
      FldArray(23).FldValue = vFldDao(RsDao("Total"))
      FldArray(23).FldIsNum = True
      
         IdCta = 0
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & vFldDao(RsDao("CodCtaTotalOld")) & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCta = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
         
      FldArray(24).FldName = "IdCuentaTotal"
      FldArray(24).FldValue = IdCta
      FldArray(24).FldIsNum = True
      
      FldArray(25).FldName = "IdUsuario"
      FldArray(25).FldValue = gUsuario.IdUsuario
      FldArray(25).FldIsNum = True
      
      FldArray(26).FldName = "FechaCreacion"
      FldArray(26).FldValue = vFldDao(RsDao("FechaCreacion"))
      FldArray(26).FldIsNum = True
      
      FldArray(27).FldName = "FEmisionOri"
      FldArray(27).FldValue = vFldDao(RsDao("FEmisionOri"))
      FldArray(27).FldIsNum = True
      
      FldArray(28).FldName = "CorrInterno"
      FldArray(28).FldValue = vFldDao(RsDao("CorrInterno"))
      FldArray(28).FldIsNum = True
      
      FldArray(29).FldName = "SaldoDoc"
      FldArray(29).FldValue = vFldDao(RsDao("SaldoDoc"))
      FldArray(29).FldIsNum = True
      
      FldArray(30).FldName = "DTE"
      FldArray(30).FldValue = Abs(vFldDao(RsDao("DTE")))
      FldArray(30).FldIsNum = True
      
      FldArray(31).FldName = "PorcentRetencion"
      FldArray(31).FldValue = vFldDao(RsDao("PorcentRetencion"))
      FldArray(31).FldIsNum = True
            
      FldArray(32).FldName = "TipoRetencion"
      FldArray(32).FldValue = vFldDao(RsDao("TipoRetencion"))
      FldArray(32).FldIsNum = True
      
      FldArray(33).FldName = "MovEdited"
      FldArray(33).FldValue = Abs(vFldDao(RsDao("MovEdited")))
      FldArray(33).FldIsNum = True
      
      FldArray(34).FldName = "OtrosVal"
      FldArray(34).FldValue = vFldDao(RsDao("OtrosVal"))
      FldArray(34).FldIsNum = True
      
      FldArray(35).FldName = "FImporF29"
      FldArray(35).FldValue = vFldDao(RsDao("FImporF29"))
      FldArray(35).FldIsNum = True
      
      FldArray(36).FldName = "NumDocRef"
      FldArray(36).FldValue = vFldDao(RsDao("NumDocRef"))
      FldArray(36).FldIsNum = False
      
         IdCta = 0
         CodCuenta = ""
         If vFldDao(RsDao("IdCtaBanco")) <> 0 Then
            Q1 = "SELECT Codigo FROM Cuentas WHERE Cuentas.IdCuenta = " & vFldDao(RsDao("IdCtaBanco"))
            Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
            If Not RsDaoAux.EOF Then
               CodCuenta = vFldDao(RsDaoAux("Codigo"))
            End If
            Call CloseRs(RsDaoAux)
            
            Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & CodCuenta & "'"
            Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               IdCta = vFld(Rs("IdCuenta"))
            End If
            Call CloseRs(Rs)
         End If
      
      FldArray(37).FldName = "IdCtaBanco"
      FldArray(37).FldValue = IdCta
      FldArray(37).FldIsNum = True
      
      FldArray(38).FldName = "TipoRelEnt"
      FldArray(38).FldValue = vFldDao(RsDao("TipoRelEnt"))
      FldArray(38).FldIsNum = True
      
         IdSuc = 0
         CodSuc = ""
         If vFldDao(RsDao("IdSucursal")) <> 0 Then
            Q1 = "SELECT Codigo FROM Sucursales WHERE Sucursales.IdSucursal = " & vFldDao(RsDao("IdSucursal"))
            Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
            If Not RsDaoAux.EOF Then
               CodSuc = vFldDao(RsDaoAux("Codigo"))
            End If
            Call CloseRs(RsDaoAux)
            
            Q1 = "SELECT IdSucursal FROM Sucursales WHERE Sucursales.Codigo = '" & CodSuc & "'"
            Q1 = Q1 & " AND Sucursales.IdEmpresa = " & IdEmpresa
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               IdSuc = vFld(Rs("IdSucursal"))
            End If
            Call CloseRs(Rs)
         End If
         
      FldArray(39).FldName = "IdSucursal"
      FldArray(39).FldValue = IdSuc
      FldArray(39).FldIsNum = True
      
      FldArray(40).FldName = "TotPagadoAnoAnt"
      FldArray(40).FldValue = vFldDao(RsDao("TotPagadoAnoAnt"))
      FldArray(40).FldIsNum = True
          
      FldArray(41).FldName = "CodCtaAfectoOld"
      FldArray(41).FldValue = vFldDao(RsDao("CodCtaAfectoOld"))
      FldArray(41).FldIsNum = False
          
      FldArray(42).FldName = "CodCtaExentoOld"
      FldArray(42).FldValue = vFldDao(RsDao("CodCtaExentoOld"))
      FldArray(42).FldIsNum = False
          
      FldArray(43).FldName = "CodCtaTotalOld"
      FldArray(43).FldValue = vFldDao(RsDao("CodCtaTotalOld"))
      FldArray(43).FldIsNum = False
          
      IdDocNew = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray)
   
      Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, 0 as IdCompCent, 0 as IdCompPago, Orden, MovDocumento.Debe, MovDocumento.Haber, Glosa, IdTipoValLib, EsTotalDoc, IdCCosto, IdAreaNeg, Codigo as CodCuenta, 0 as IdCuenta "
      Fld2 = " IdEmpresa, Ano, IdCompCent, IdCompPago, Orden, Debe, Haber, Glosa, IdTipoValLib, EsTotalDoc, IdCCosto, IdAreaNeg, CodCuentaOld, IdCuenta "
      
      Q1 = "SELECT " & Fld & " FROM MovDocumento "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa  ' dado que el año anterior es Access y empresas separadas
      Q1 = Q1 & " INNER JOIN Cuentas ON Cuentas.IdCuenta = MovDocumento.IdCuenta "
      Q1 = Q1 & " WHERE IdDoc = " & IdDoc
      Q1 = Q1 & " ORDER BY IdDoc, IdMovDoc "
      Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
      
      Do While Not RsDaoAux.EOF
               
               
         'obtenemos el IdCuenta en SQL Server e insertamos en MovDocumento
         IdCta = 0
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & RsDaoAux("CodCuenta") & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCta = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
                              
         'obtenemos el Área de Negocio y el Centro de Costo correspondientes en SQL Server
         IdAreaNeg = 0
         CodAreaNeg = ""
         If vFldDao(RsDaoAux("IdAreaNeg")) > 0 Then
            Q1 = "SELECT Codigo FROM AreaNegocio WHERE AreaNegocio.IdAreaNegocio = " & vFldDao(RsDaoAux("IdAreaNeg"))
            Set RsDaoAux2 = OpenRsDao(DbAnoAnt, Q1)
            If Not RsDaoAux2.EOF Then
               CodAreaNeg = vFldDao(RsDaoAux2("Codigo"))
            End If
            Call CloseRs(RsDaoAux2)
            
            Q1 = "SELECT IdAreaNegocio FROM AreaNegocio WHERE AreaNegocio.Codigo = '" & CodAreaNeg & "'"
            Q1 = Q1 & " AND AreaNegocio.IdEmpresa = " & IdEmpresa
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               IdAreaNeg = vFld(Rs("IdAreaNegocio"))
            End If
            Call CloseRs(Rs)
         End If
         
         IdCCosto = 0
         CodCCosto = ""
         If vFldDao(RsDaoAux("IdCCosto")) > 0 Then
            Q1 = "SELECT Codigo FROM CentroCosto WHERE CentroCosto.IdCCosto = " & vFldDao(RsDaoAux("IdCCosto"))
            Set RsDaoAux2 = OpenRsDao(DbAnoAnt, Q1)
            If Not RsDaoAux2.EOF Then
               CodCCosto = vFldDao(RsDaoAux2("Codigo"))
            End If
            Call CloseRs(RsDaoAux2)
            
            Q1 = "SELECT IdCCosto FROM CentroCosto WHERE CentroCosto.Codigo = '" & CodCCosto & "'"
            Q1 = Q1 & " AND CentroCosto.IdEmpresa = " & IdEmpresa
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               IdCCosto = vFld(Rs("IdCCosto"))
            End If
            Call CloseRs(Rs)
         End If
         
         'Ahora el Insert del Mov Documento
         Q1 = "INSERT INTO MovDocumento ( IdDoc, " & Fld2 & ") VALUES(" & IdDocNew & ","
         
         For i = 0 To RsDaoAux.Fields.Count - 1
            If RsDaoAux(i).Name = "IdCCosto" Then
               Q1 = Q1 & IdCCosto & ","
            ElseIf RsDaoAux(i).Name = "IdAreaNeg" Then
               Q1 = Q1 & IdAreaNeg & ","
            ElseIf RsDaoAux(i).Name = "IdCuenta" Then
               Q1 = Q1 & IdCta & ","
               Exit For
            ElseIf RsDaoAux(i).Type = dbText Or RsDaoAux(i).Type = dbMemo Or RsDaoAux(i).Type = dbChar Then
               Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDaoAux(i))) & "',"
            Else
               Q1 = Q1 & vFldDao(RsDaoAux(i)) & ","
            End If
            
         Next i
         
         Q1 = Left(Q1, Len(Q1) - 1) & ")"
         Call ExecSQL(DbMain, Q1)
        
         RsDaoAux.MoveNext
      
      Loop
      
      Call CloseRs(RsDaoAux)
      
      RsDao.MoveNext
   
   Loop
   
   Call CloseRs(RsDao)

   'Mensaje con cantidad
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   Set RsDao = OpenRsDao(DbMain, Q1)
   
   If RsDao.EOF = False Then
      MsgBox1 "Se encontraron " & vFld(RsDao("N")) & " documentos del año anterior, en estado Centralizado o Pagado, con saldo distinto de cero.", vbInformation
   End If
   
   Call CloseRs(RsDao)

   'limpiamos FExported en tabla nueva
   'limpiamos IdCompCent e IdCompPago que apuntan a comprobantes del año anterior
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET IdCompCent = 0, IdCompPago = 0, FExported = 0"
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
  

End Function

Public Function CopyDocsFromAccessToSQLServerNew(ByVal DbAnoAnt As Database, ByVal IdEmpresa As Long, ByVal Ano As Long)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Rs2 As Recordset
   Dim RsDao As dao.Recordset
   Dim RsDao2 As dao.Recordset
   Dim RsDaoAux As dao.Recordset, RsDaoAux2 As dao.Recordset
   Dim i As Integer
   Dim Where As String
   Dim IdCta As Long, IdSuc As Long, IdEntidad As Long
   Dim FldArray(43) As AdvTbAddNew_t
   Dim IdDocNew As Long
   Dim CodCuenta As String, CodSuc As String, RutEntidad As String
   Dim IdDoc As Long
   Dim Fld As String, Fld2 As String
   Dim IdAreaNeg As Long, IdCCosto As Long
   Dim CodAreaNeg As String
   Dim CodCCosto As String
  
   'marcamos los que vamos a exportar con -1
   Q1 = "UPDATE Documento SET FExported = -1"
'   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) "    'para que siempre pueda volver a importar por si tiene que volver a hacerlo
   Q1 = Q1 & " WHERE SaldoDoc <> 0 AND Estado IN(" & ED_CENTRALIZADO & "," & ED_PAGADO & ") AND TipoLib IN( " & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
   Call ExecSQLDao(DbAnoAnt, Q1)   'IdEmpresa y Año no se requiere, dado que son empresas/Año separados

   Where = " WHERE FExported < 0"
   
   'Agregamos docs marcados para exportar
   'Calculamos Total Pagado Año Anterior para que tengamos los saldos OK
   Q1 = " SELECT IdDoc, IdCompCent, IdCompPago, TipoLib, TipoDoc, NumDoc, NumDocHasta, Giro, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, Documento.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, IdCuentaOtroImp, Total, IdCuentaTotal, IdUsuario, FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDocTmp as OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal,"
   Q1 = Q1 & " iif(SaldoDoc IS NULL or SaldoDoc = 0, 0, iif(SaldoDoc > 0, Total - abs(SaldoDoc), -1 *(Total - abs(SaldoDoc))))  As TotPagadoAnoAnt, "
   Q1 = Q1 & IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, "
   Q1 = Q1 & " Cuentas.Codigo as CodCtaAfectoOld, Cuentas1.Codigo As CodCtaExentoOld, Cuentas2.Codigo As CodCtaTotalOld "
   Q1 = Q1 & " FROM ((Documento "
   Q1 = Q1 & " LEFT JOIN Cuentas ON Documento.IdCuentaAfecto = Cuentas.IdCuenta ) "
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas1 ON Documento.IdCuentaExento = Cuentas1.IdCuenta ) "
   Q1 = Q1 & " LEFT JOIN Cuentas As Cuentas2 ON Documento.IdCuentaTotal = Cuentas2.IdCuenta "
   Q1 = Q1 & Where
   Set RsDao = OpenRsDao(DbAnoAnt, Q1)
   
   Do While Not RsDao.EOF
   
      Q1 = " SELECT NUMDOC "
      Q1 = Q1 & " From Documento "
      Q1 = Q1 & " Where NumDoc = '" & RsDao("NUMDOC") & "' And IdEmpresa = " & gEmpresa.id & " And Ano = " & gEmpresa.Ano & " And TipoLib = '" & RsDao("TipoLib") & "' And TipoDoc = '" & RsDao("TipoDoc") & "' And Total = " & RsDao("Total") & ""
      Set Rs2 = OpenRs(DbMain, Q1)
       
       
      If Not Rs2.EOF Then
            Q1 = " UPDATE Documento SET SaldoDoc = " & RsDao("SaldoDoc") & " , TotPagadoAnoAnt = " & RsDao("TotPagadoAnoAnt") & " "
            Q1 = Q1 & " Where NumDoc = '" & RsDao("NUMDOC") & "' And IdEmpresa = " & gEmpresa.id & " And Ano = " & gEmpresa.Ano & " And TipoLib = '" & RsDao("TipoLib") & "' And TipoDoc = '" & RsDao("TipoDoc") & "' And Total = " & RsDao("Total") & ""
            Call ExecSQL(DbMain, Q1)
      Else
      
           IdDoc = vFldDao(RsDao("IdDoc"))
                       
           FldArray(0).FldName = "IdEmpresa"
           FldArray(0).FldValue = IdEmpresa
           FldArray(0).FldIsNum = True
                       
           FldArray(1).FldName = "Ano"
           FldArray(1).FldValue = Ano        ' hay que almacenar el año al que corresponde el documento en la DB no el año del documento mismo. Esto no: Year(vFldDao(RsDao("FEmision")))
           FldArray(1).FldIsNum = True
           
           FldArray(2).FldName = "TipoLib"
           FldArray(2).FldValue = vFldDao(RsDao("TipoLib"))
           FldArray(2).FldIsNum = True
           
           FldArray(3).FldName = "TipoDoc"
           FldArray(3).FldValue = vFldDao(RsDao("TipoDoc"))
           FldArray(3).FldIsNum = True
           
           FldArray(4).FldName = "NumDoc"
           FldArray(4).FldValue = vFldDao(RsDao("NumDoc"))
           FldArray(4).FldIsNum = False
                 
           FldArray(5).FldName = "NumDocHasta"
           FldArray(5).FldValue = vFldDao(RsDao("NumDocHasta"))
           FldArray(5).FldIsNum = False
                 
           FldArray(6).FldName = "Giro"
           FldArray(6).FldValue = Abs(vFldDao(RsDao("Giro")))
           FldArray(6).FldIsNum = True
                 
              IdEntidad = 0
              RutEntidad = ""
              Q1 = "SELECT Rut FROM Entidades WHERE Entidades.IdEntidad = " & vFldDao(RsDao("IdEntidad"))
              Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
              If Not RsDaoAux.EOF Then
                 RutEntidad = vFldDao(RsDaoAux("Rut"))
              End If
              Call CloseRs(RsDaoAux)
              
              Q1 = "SELECT IdEntidad FROM Entidades WHERE Entidades.Rut = '" & RutEntidad & "'"
              Q1 = Q1 & " AND Entidades.IdEmpresa = " & IdEmpresa
              Set Rs = OpenRs(DbMain, Q1)
              If Not Rs.EOF Then
                 IdEntidad = vFld(Rs("IdEntidad"))
              End If
              Call CloseRs(Rs)
              
                 
           FldArray(7).FldName = "IdEntidad"
           FldArray(7).FldValue = IdEntidad
           FldArray(7).FldIsNum = True
                 
           FldArray(8).FldName = "TipoEntidad"
           FldArray(8).FldValue = vFldDao(RsDao("TipoEntidad"))
           FldArray(8).FldIsNum = True
                 
           FldArray(9).FldName = "RutEntidad"
           FldArray(9).FldValue = vFldDao(RsDao("RutEntidad"))
           FldArray(9).FldIsNum = False
                 
           FldArray(10).FldName = "NombreEntidad"
           FldArray(10).FldValue = vFldDao(RsDao("NombreEntidad"))
           FldArray(10).FldIsNum = False
                 
           FldArray(11).FldName = "FEmision"
           FldArray(11).FldValue = vFldDao(RsDao("FEmision"))
           FldArray(11).FldIsNum = True
           
           FldArray(12).FldName = "FVenc"
           FldArray(12).FldValue = vFldDao(RsDao("FVenc"))
           FldArray(12).FldIsNum = True
           
           FldArray(13).FldName = "Descrip"
           FldArray(13).FldValue = vFldDao(RsDao("Descrip"))
           FldArray(13).FldIsNum = False
                    
           FldArray(14).FldName = "Estado"
           FldArray(14).FldValue = vFldDao(RsDao("Estado"))
           FldArray(14).FldIsNum = True
           
           FldArray(15).FldName = "Exento"
           FldArray(15).FldValue = vFldDao(RsDao("Exento"))
           FldArray(15).FldIsNum = True
           
              IdCta = 0
              Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & vFldDao(RsDao("CodCtaExentoOld")) & "'"
              Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
              Set Rs = OpenRs(DbMain, Q1)
              If Not Rs.EOF Then
                 IdCta = vFld(Rs("IdCuenta"))
              End If
              Call CloseRs(Rs)
           
           FldArray(16).FldName = "IdCuentaExento"
           FldArray(16).FldValue = IdCta
           FldArray(16).FldIsNum = True
           
           FldArray(17).FldName = "Afecto"
           FldArray(17).FldValue = vFldDao(RsDao("Afecto"))
           FldArray(17).FldIsNum = True
           
              IdCta = 0
              Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & vFldDao(RsDao("CodCtaAfectoOld")) & "'"
              Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
              Set Rs = OpenRs(DbMain, Q1)
              If Not Rs.EOF Then
                 IdCta = vFld(Rs("IdCuenta"))
              End If
              Call CloseRs(Rs)
              
           FldArray(18).FldName = "IdCuentaAfecto"
           FldArray(18).FldValue = IdCta
           FldArray(18).FldIsNum = True
           
           FldArray(19).FldName = "IVA"
           FldArray(19).FldValue = vFldDao(RsDao("IVA"))
           FldArray(19).FldIsNum = True
           
           IdCta = 0
           If vFldDao(RsDao("IVA")) <> 0 Then
           
              CodCuenta = ""
              Q1 = "SELECT Codigo FROM Cuentas WHERE Cuentas.IdCuenta = " & vFldDao(RsDao("IdCuentaIVA"))
              Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
              If Not RsDaoAux.EOF Then
                 CodCuenta = vFldDao(RsDaoAux("Codigo"))
              End If
              Call CloseRs(RsDaoAux)
              
              Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & CodCuenta & "'"
              Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
              Set Rs = OpenRs(DbMain, Q1)
              If Not Rs.EOF Then
                 IdCta = vFld(Rs("IdCuenta"))
              End If
              Call CloseRs(Rs)
                    
           End If
           FldArray(20).FldName = "IdCuentaIVA"
           FldArray(20).FldValue = IdCta
           FldArray(20).FldIsNum = True
           
           FldArray(21).FldName = "OtroImp"
           FldArray(21).FldValue = vFldDao(RsDao("OtroImp"))
           FldArray(21).FldIsNum = True
           
           IdCta = 0
           If vFldDao(RsDao("OtroImp")) <> 0 Then
           
              CodCuenta = ""
              If vFldDao(RsDao("IdCuentaOtroImp")) <> 0 Then
                 Q1 = "SELECT Codigo FROM Cuentas WHERE Cuentas.IdCuenta = " & vFldDao(RsDao("IdCuentaOtroImp"))
                 Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
                 If Not RsDaoAux.EOF Then
                    CodCuenta = vFldDao(RsDaoAux("Codigo"))
                 End If
                 Call CloseRs(RsDaoAux)
                 
                 Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & CodCuenta & "'"
                 Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
                 Set Rs = OpenRs(DbMain, Q1)
                 If Not Rs.EOF Then
                    IdCta = vFld(Rs("IdCuenta"))
                 End If
                 Call CloseRs(Rs)
              End If
                       
           End If
           FldArray(22).FldName = "IdCuentaOtroImp"
           FldArray(22).FldValue = IdCta
           FldArray(22).FldIsNum = True
           
           FldArray(23).FldName = "Total"
           FldArray(23).FldValue = vFldDao(RsDao("Total"))
           FldArray(23).FldIsNum = True
           
              IdCta = 0
              Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & vFldDao(RsDao("CodCtaTotalOld")) & "'"
              Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
              Set Rs = OpenRs(DbMain, Q1)
              If Not Rs.EOF Then
                 IdCta = vFld(Rs("IdCuenta"))
              End If
              Call CloseRs(Rs)
              
           FldArray(24).FldName = "IdCuentaTotal"
           FldArray(24).FldValue = IdCta
           FldArray(24).FldIsNum = True
           
           FldArray(25).FldName = "IdUsuario"
           FldArray(25).FldValue = gUsuario.IdUsuario
           FldArray(25).FldIsNum = True
           
           FldArray(26).FldName = "FechaCreacion"
           FldArray(26).FldValue = vFldDao(RsDao("FechaCreacion"))
           FldArray(26).FldIsNum = True
           
           FldArray(27).FldName = "FEmisionOri"
           FldArray(27).FldValue = vFldDao(RsDao("FEmisionOri"))
           FldArray(27).FldIsNum = True
           
           FldArray(28).FldName = "CorrInterno"
           FldArray(28).FldValue = vFldDao(RsDao("CorrInterno"))
           FldArray(28).FldIsNum = True
           
           FldArray(29).FldName = "SaldoDoc"
           FldArray(29).FldValue = vFldDao(RsDao("SaldoDoc"))
           FldArray(29).FldIsNum = True
           
           FldArray(30).FldName = "DTE"
           FldArray(30).FldValue = Abs(vFldDao(RsDao("DTE")))
           FldArray(30).FldIsNum = True
           
           FldArray(31).FldName = "PorcentRetencion"
           FldArray(31).FldValue = vFldDao(RsDao("PorcentRetencion"))
           FldArray(31).FldIsNum = True
                 
           FldArray(32).FldName = "TipoRetencion"
           FldArray(32).FldValue = vFldDao(RsDao("TipoRetencion"))
           FldArray(32).FldIsNum = True
           
           FldArray(33).FldName = "MovEdited"
           FldArray(33).FldValue = Abs(vFldDao(RsDao("MovEdited")))
           FldArray(33).FldIsNum = True
           
           FldArray(34).FldName = "OtrosVal"
           FldArray(34).FldValue = vFldDao(RsDao("OtrosVal"))
           FldArray(34).FldIsNum = True
           
           FldArray(35).FldName = "FImporF29"
           FldArray(35).FldValue = vFldDao(RsDao("FImporF29"))
           FldArray(35).FldIsNum = True
           
           FldArray(36).FldName = "NumDocRef"
           FldArray(36).FldValue = vFldDao(RsDao("NumDocRef"))
           FldArray(36).FldIsNum = False
           
              IdCta = 0
              CodCuenta = ""
              If vFldDao(RsDao("IdCtaBanco")) <> 0 Then
                 Q1 = "SELECT Codigo FROM Cuentas WHERE Cuentas.IdCuenta = " & vFldDao(RsDao("IdCtaBanco"))
                 Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
                 If Not RsDaoAux.EOF Then
                    CodCuenta = vFldDao(RsDaoAux("Codigo"))
                 End If
                 Call CloseRs(RsDaoAux)
                 
                 Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & CodCuenta & "'"
                 Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
                 Set Rs = OpenRs(DbMain, Q1)
                 If Not Rs.EOF Then
                    IdCta = vFld(Rs("IdCuenta"))
                 End If
                 Call CloseRs(Rs)
              End If
           
           FldArray(37).FldName = "IdCtaBanco"
           FldArray(37).FldValue = IdCta
           FldArray(37).FldIsNum = True
           
           FldArray(38).FldName = "TipoRelEnt"
           FldArray(38).FldValue = vFldDao(RsDao("TipoRelEnt"))
           FldArray(38).FldIsNum = True
           
              IdSuc = 0
              CodSuc = ""
              If vFldDao(RsDao("IdSucursal")) <> 0 Then
                 Q1 = "SELECT Codigo FROM Sucursales WHERE Sucursales.IdSucursal = " & vFldDao(RsDao("IdSucursal"))
                 Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
                 If Not RsDaoAux.EOF Then
                    CodSuc = vFldDao(RsDaoAux("Codigo"))
                 End If
                 Call CloseRs(RsDaoAux)
                 
                 Q1 = "SELECT IdSucursal FROM Sucursales WHERE Sucursales.Codigo = '" & CodSuc & "'"
                 Q1 = Q1 & " AND Sucursales.IdEmpresa = " & IdEmpresa
                 Set Rs = OpenRs(DbMain, Q1)
                 If Not Rs.EOF Then
                    IdSuc = vFld(Rs("IdSucursal"))
                 End If
                 Call CloseRs(Rs)
              End If
              
           FldArray(39).FldName = "IdSucursal"
           FldArray(39).FldValue = IdSuc
           FldArray(39).FldIsNum = True
           
           FldArray(40).FldName = "TotPagadoAnoAnt"
           FldArray(40).FldValue = vFldDao(RsDao("TotPagadoAnoAnt"))
           FldArray(40).FldIsNum = True
               
           FldArray(41).FldName = "CodCtaAfectoOld"
           FldArray(41).FldValue = vFldDao(RsDao("CodCtaAfectoOld"))
           FldArray(41).FldIsNum = False
               
           FldArray(42).FldName = "CodCtaExentoOld"
           FldArray(42).FldValue = vFldDao(RsDao("CodCtaExentoOld"))
           FldArray(42).FldIsNum = False
               
           FldArray(43).FldName = "CodCtaTotalOld"
           FldArray(43).FldValue = vFldDao(RsDao("CodCtaTotalOld"))
           FldArray(43).FldIsNum = False
               
           IdDocNew = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray)
        
           Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, 0 as IdCompCent, 0 as IdCompPago, Orden, MovDocumento.Debe, MovDocumento.Haber, Glosa, IdTipoValLib, EsTotalDoc, IdCCosto, IdAreaNeg, Codigo as CodCuenta, 0 as IdCuenta "
           Fld2 = " IdEmpresa, Ano, IdCompCent, IdCompPago, Orden, Debe, Haber, Glosa, IdTipoValLib, EsTotalDoc, IdCCosto, IdAreaNeg, CodCuentaOld, IdCuenta "
           
           Q1 = "SELECT " & Fld & " FROM MovDocumento "    ' WHERE Cuentas.IdEmpresa = " & IdEmpresa  ' dado que el año anterior es Access y empresas separadas
           Q1 = Q1 & " INNER JOIN Cuentas ON Cuentas.IdCuenta = MovDocumento.IdCuenta "
           Q1 = Q1 & " WHERE IdDoc = " & IdDoc
           Q1 = Q1 & " ORDER BY IdDoc, IdMovDoc "
           Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
           
           Do While Not RsDaoAux.EOF
                    
                    
              'obtenemos el IdCuenta en SQL Server e insertamos en MovDocumento
              IdCta = 0
              Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & RsDaoAux("CodCuenta") & "'"
              Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
              Set Rs = OpenRs(DbMain, Q1)
              If Not Rs.EOF Then
                 IdCta = vFld(Rs("IdCuenta"))
              End If
              Call CloseRs(Rs)
                                   
              'obtenemos el Área de Negocio y el Centro de Costo correspondientes en SQL Server
              IdAreaNeg = 0
              CodAreaNeg = ""
              If vFldDao(RsDaoAux("IdAreaNeg")) > 0 Then
                 Q1 = "SELECT Codigo FROM AreaNegocio WHERE AreaNegocio.IdAreaNegocio = " & vFldDao(RsDaoAux("IdAreaNeg"))
                 Set RsDaoAux2 = OpenRsDao(DbAnoAnt, Q1)
                 If Not RsDaoAux2.EOF Then
                    CodAreaNeg = vFldDao(RsDaoAux2("Codigo"))
                 End If
                 Call CloseRs(RsDaoAux2)
                 
                 Q1 = "SELECT IdAreaNegocio FROM AreaNegocio WHERE AreaNegocio.Codigo = '" & CodAreaNeg & "'"
                 Q1 = Q1 & " AND AreaNegocio.IdEmpresa = " & IdEmpresa
                 Set Rs = OpenRs(DbMain, Q1)
                 If Not Rs.EOF Then
                    IdAreaNeg = vFld(Rs("IdAreaNegocio"))
                 End If
                 Call CloseRs(Rs)
              End If
              
              IdCCosto = 0
              CodCCosto = ""
              If vFldDao(RsDaoAux("IdCCosto")) > 0 Then
                 Q1 = "SELECT Codigo FROM CentroCosto WHERE CentroCosto.IdCCosto = " & vFldDao(RsDaoAux("IdCCosto"))
                 Set RsDaoAux2 = OpenRsDao(DbAnoAnt, Q1)
                 If Not RsDaoAux2.EOF Then
                    CodCCosto = vFldDao(RsDaoAux2("Codigo"))
                 End If
                 Call CloseRs(RsDaoAux2)
                 
                 Q1 = "SELECT IdCCosto FROM CentroCosto WHERE CentroCosto.Codigo = '" & CodCCosto & "'"
                 Q1 = Q1 & " AND CentroCosto.IdEmpresa = " & IdEmpresa
                 Set Rs = OpenRs(DbMain, Q1)
                 If Not Rs.EOF Then
                    IdCCosto = vFld(Rs("IdCCosto"))
                 End If
                 Call CloseRs(Rs)
              End If
              
              'Ahora el Insert del Mov Documento
              Q1 = "INSERT INTO MovDocumento ( IdDoc, " & Fld2 & ") VALUES(" & IdDocNew & ","
              
              For i = 0 To RsDaoAux.Fields.Count - 1
                 If RsDaoAux(i).Name = "IdCCosto" Then
                    Q1 = Q1 & IdCCosto & ","
                 ElseIf RsDaoAux(i).Name = "IdAreaNeg" Then
                    Q1 = Q1 & IdAreaNeg & ","
                 ElseIf RsDaoAux(i).Name = "IdCuenta" Then
                    Q1 = Q1 & IdCta & ","
                    Exit For
                 ElseIf RsDaoAux(i).Type = dbText Or RsDaoAux(i).Type = dbMemo Or RsDaoAux(i).Type = dbChar Then
                    Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDaoAux(i))) & "',"
                 Else
                    Q1 = Q1 & vFldDao(RsDaoAux(i)) & ","
                 End If
                 
              Next i
              
              Q1 = Left(Q1, Len(Q1) - 1) & ")"
              Call ExecSQL(DbMain, Q1)
             
              RsDaoAux.MoveNext
           
           Loop
           
           Call CloseRs(RsDaoAux)
      End If
      Call CloseRs(Rs2)
      RsDao.MoveNext
   
   Loop
   
   Call CloseRs(RsDao)
   
   'Mensaje con cantidad
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM Documento "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano - 1
   'Set RsDao2 = OpenRsDao(DbMain, Q1)
   Set Rs2 = OpenRs(DbMain, Q1)
   
   If Rs2.EOF = False Then
      MsgBox1 "Se encontraron " & Rs2("N") & " documentos del año anterior, en estado Centralizado o Pagado, con saldo distinto de cero.", vbInformation
   End If
   
   Call CloseRs(Rs2)

   'limpiamos FExported en tabla nueva
   'limpiamos IdCompCent e IdCompPago que apuntan a comprobantes del año anterior
   Q1 = "UPDATE Documento "
   Q1 = Q1 & " SET IdCompCent = 0, IdCompPago = 0, FExported = 0"
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
  

End Function



Public Function CopyActFijoFromAccessToSQLServer(ByVal DbAnoAnt As Database, ByVal IdEmpresa As Long, ByVal Ano As Long)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim RsDao As dao.Recordset
   Dim RsDaoAux As dao.Recordset, RsDaoAux2 As dao.Recordset
   Dim i As Integer
   Dim Where As String
   Dim IdCta As Long
   Dim FldArray(37) As AdvTbAddNew_t
   Dim CodCuenta As String
   Dim Fld As String, Fld2 As String
   Dim IdActFijoNew As Long, IdActFijo As Long
   Dim IdGrupo As Long, IdGrupoOld As Long, idcomp As Long
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   Dim NImported As Long
  
   'marcamos los ActFijos que vamos a exportar con -1
   Q1 = "UPDATE MovActivoFijo SET IdActFijoOldTmp = IdActFijo, FExported = -1"
'   Q1 = Q1 & " WHERE (FExported = 0 OR FExported Is NULL) "    'para que siempre pueda volver a importar por si tiene que volver a hacerlo
   Q1 = Q1 & " WHERE ( VidaUtilResidual > 0 OR NoDepreciable <> 0 OR ValorLibro = 1 )"
   Call ExecSQLDao(DbAnoAnt, Q1)   'IdEmpresa y Año no se requiere, dado que son empresas/Año separados

   Where = " WHERE FExported < 0"
   
   'Agregamos Act Fijos marcados para exportar
   'Calculamos Total Pagado Año Anterior para que tengamos los saldos OK
   Q1 = " SELECT IdActFijo, TipoMovAF, Fecha, Cantidad, Descrip, Neto, IVA, 0 As Cred4Porc, DepNormal, DepAcelerada, IdCuenta, DepNormalHist, DepAceleradaHist, NetoVenta, IVAVenta, FechaVentaBaja, TipoDep, TipoDepHist, DepAcumHist, VidaUtil, DepAcumFinal, VidaUtilResidual, FExported, FechaUtilizacion, NoDepreciable, -1 As ValCred33, ValReajustadoNeto, IdActFijoOldTmp As IdActFijoOld, TotalmenteDepreciado, ValorLibro, iif(MovActivoFijo.Cred4Porc <> 0, MovActivoFijo.Cred4Porc, MovActivoFijo.Cred4PorcAnoInit) As Cred4PorcAnoInit, DepInstant, DepDecimaParte, DepInstantHist, DepDecimaParteHist, VidaUtilAnos "
   Q1 = Q1 & "," & IdEmpresa & " AS IdEmpresa, " & Ano & " As Ano "
   Q1 = Q1 & " FROM MovActivoFijo "
   Q1 = Q1 & " WHERE MovActivoFijo.FExported < 0 "
   Q1 = Q1 & " ORDER BY MovActivoFijo.Fecha "
   Set RsDao = OpenRsDao(DbAnoAnt, Q1)
   
      
   Do While Not RsDao.EOF
       
      IdActFijo = vFldDao(RsDao("IdActFijo"))
                  
      FldArray(0).FldName = "IdEmpresa"
      FldArray(0).FldValue = IdEmpresa
      FldArray(0).FldIsNum = True
                  
      FldArray(1).FldName = "Ano"
      FldArray(1).FldValue = Ano        ' hay que almacenar el año al que corresponde el documento en la DB no el año del documento mismo. Esto no: Year(vFldDao(RsDao("FEmision")))
      FldArray(1).FldIsNum = True
      
      FldArray(2).FldName = "TipoMovAF"
      FldArray(2).FldValue = vFldDao(RsDao("TipoMovAF"))
      FldArray(2).FldIsNum = True
      
      FldArray(3).FldName = "Fecha"
      FldArray(3).FldValue = vFldDao(RsDao("Fecha"))
      FldArray(3).FldIsNum = True
      
      FldArray(4).FldName = "Cantidad"
      FldArray(4).FldValue = vFldDao(RsDao("Cantidad"))
      FldArray(4).FldIsNum = True
            
      FldArray(5).FldName = "Descrip"
      FldArray(5).FldValue = vFldDao(RsDao("Descrip"))
      FldArray(5).FldIsNum = False
            
      FldArray(6).FldName = "Neto"
      FldArray(6).FldValue = vFldDao(RsDao("Neto"))
      FldArray(6).FldIsNum = True
                        
      FldArray(7).FldName = "IVA"
      FldArray(7).FldValue = vFldDao(RsDao("IVA"))
      FldArray(7).FldIsNum = True
            
      FldArray(8).FldName = "Cred4Porc"
      FldArray(8).FldValue = Abs(vFldDao(RsDao("Cred4Porc")))
      FldArray(8).FldIsNum = True
            
      FldArray(9).FldName = "DepNormal"
      FldArray(9).FldValue = vFldDao(RsDao("DepNormal"))
      FldArray(9).FldIsNum = True
            
      FldArray(10).FldName = "DepAcelerada"
      FldArray(10).FldValue = vFldDao(RsDao("DepAcelerada"))
      FldArray(10).FldIsNum = True
      
      IdCta = 0
      CodCuenta = ""
      If vFldDao(RsDao("IdCuenta")) <> 0 Then
      
         Q1 = "SELECT Codigo FROM Cuentas WHERE Cuentas.IdCuenta = " & vFldDao(RsDao("IdCuenta"))
         Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
         If Not RsDaoAux.EOF Then
            CodCuenta = vFldDao(RsDaoAux("Codigo"))
         End If
         Call CloseRs(RsDaoAux)
         
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & CodCuenta & "'"
         Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdCta = vFld(Rs("IdCuenta"))
         End If
         Call CloseRs(Rs)
      
      End If
      
      FldArray(11).FldName = "IdCuenta"
      FldArray(11).FldValue = IdCta
      FldArray(11).FldIsNum = True
      
      FldArray(12).FldName = "DepNormalHist"
      FldArray(12).FldValue = vFldDao(RsDao("DepNormalHist"))
      FldArray(12).FldIsNum = True
      
      FldArray(13).FldName = "DepAceleradaHist"
      FldArray(13).FldValue = vFldDao(RsDao("DepAceleradaHist"))
      FldArray(13).FldIsNum = True
               
      FldArray(14).FldName = "NetoVenta"
      FldArray(14).FldValue = vFldDao(RsDao("NetoVenta"))
      FldArray(14).FldIsNum = True
      
      FldArray(15).FldName = "IVAVenta"
      FldArray(15).FldValue = vFldDao(RsDao("IVAVenta"))
      FldArray(15).FldIsNum = True
            
      FldArray(16).FldName = "FechaVentaBaja"
      FldArray(16).FldValue = vFldDao(RsDao("FechaVentaBaja"))
      FldArray(16).FldIsNum = True
      
      FldArray(17).FldName = "TipoDep"
      FldArray(17).FldValue = vFldDao(RsDao("TipoDep"))
      FldArray(17).FldIsNum = True
      
      FldArray(18).FldName = "TipoDepHist"
      FldArray(18).FldValue = vFldDao(RsDao("TipoDepHist"))
      FldArray(18).FldIsNum = True
      
      FldArray(19).FldName = "DepAcumHist"
      FldArray(19).FldValue = vFldDao(RsDao("DepAcumHist"))
      FldArray(19).FldIsNum = True
      
      FldArray(20).FldName = "VidaUtil"
      FldArray(20).FldValue = vFldDao(RsDao("VidaUtil"))
      FldArray(20).FldIsNum = True
      
      FldArray(21).FldName = "DepAcumFinal"
      FldArray(21).FldValue = vFldDao(RsDao("DepAcumFinal"))
      FldArray(21).FldIsNum = True
      
      FldArray(22).FldName = "VidaUtilResidual"
      FldArray(22).FldValue = vFldDao(RsDao("VidaUtilResidual"))
      FldArray(22).FldIsNum = True
      
      FldArray(23).FldName = "FExported"
      FldArray(23).FldValue = -1
      FldArray(23).FldIsNum = True
      
      FldArray(24).FldName = "FechaUtilizacion"
      FldArray(24).FldValue = vFldDao(RsDao("FechaUtilizacion"))
      FldArray(24).FldIsNum = True
      
      FldArray(25).FldName = "NoDepreciable"
      FldArray(25).FldValue = Abs(vFldDao(RsDao("NoDepreciable")))
      FldArray(25).FldIsNum = True
      
      FldArray(26).FldName = "ValCred33"
      FldArray(26).FldValue = vFldDao(RsDao("ValCred33"))
      FldArray(26).FldIsNum = True
      
      FldArray(27).FldName = "ValReajustadoNeto"
      FldArray(27).FldValue = vFldDao(RsDao("ValReajustadoNeto"))
      FldArray(27).FldIsNum = True
      
      FldArray(28).FldName = "IdActFijoOld"
      FldArray(28).FldValue = 0           'vFldDao(RsDao("IdActFijoOld"))
      FldArray(28).FldIsNum = True
      
      FldArray(29).FldName = "TotalmenteDepreciado"
      FldArray(29).FldValue = Abs(vFldDao(RsDao("TotalmenteDepreciado")))
      FldArray(29).FldIsNum = True
      
      FldArray(30).FldName = "ValorLibro"
      FldArray(30).FldValue = vFldDao(RsDao("ValorLibro"))
      FldArray(30).FldIsNum = True
      
      FldArray(31).FldName = "Cred4PorcAnoInit"
      FldArray(31).FldValue = Abs(vFldDao(RsDao("Cred4PorcAnoInit")))
      FldArray(31).FldIsNum = True
            
      FldArray(32).FldName = "FechaImportFile"
      FldArray(32).FldValue = 0
      FldArray(32).FldIsNum = True
      
      FldArray(33).FldName = "DepInstant"
      FldArray(33).FldValue = vFldDao(RsDao("DepInstant"))
      FldArray(33).FldIsNum = True
      
      FldArray(34).FldName = "DepDecimaParte"
      FldArray(34).FldValue = vFldDao(RsDao("DepDecimaParte"))
      FldArray(34).FldIsNum = True
      
      FldArray(35).FldName = "DepInstantHist"
      FldArray(35).FldValue = vFldDao(RsDao("DepInstantHist"))
      FldArray(35).FldIsNum = True
      
      FldArray(36).FldName = "DepDecimaParteHist"
      FldArray(36).FldValue = vFldDao(RsDao("DepDecimaParteHist"))
      FldArray(36).FldIsNum = False
      
      FldArray(37).FldName = "VidaUtilAnos"
      FldArray(37).FldValue = vFldDao(RsDao("VidaUtilAnos"))
      FldArray(37).FldIsNum = True
      
          
      IdActFijoNew = AdvTbAddNewMult(DbMain, "MovActivoFijo", "IdActFijo", FldArray)
      
      'Insertamos la ficha financiera del Act Fijo
      Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, PrecioFactura, DerechosIntern, Transporte, ObrasAdapt, PrecioAdquis, IVARecuperable, FormacionPers, ObrasReubic, TotalGastos, FechaIncorporacion, FechaDisponible, AdquiOtrosConceptos, GastoOtrosConceptos, SinDetComps, 0 as IdFichaOld, AFGrupos.NombGrupo, AFGrupos.IdGrupo "
      Fld2 = " IdEmpresa, Ano, PrecioFactura, DerechosIntern, Transporte, ObrasAdapt, PrecioAdquis, IVARecuperable, FormacionPers, ObrasReubic, TotalGastos, FechaIncorporacion, FechaDisponible, AdquiOtrosConceptos, GastoOtrosConceptos, SinDetComps, IdFichaOld "
   
      Q1 = "SELECT " & Fld & " FROM ActFijoFicha "
      Q1 = Q1 & " INNER JOIN AFGrupos ON ActFijoFicha.IdGrupo = AFGrupos.IdGrupo "
      Q1 = Q1 & " WHERE IdActFijo = " & IdActFijo
      Q1 = Q1 & " ORDER BY IdActFijo "
      Set RsDaoAux = OpenRsDao(DbAnoAnt, Q1)
      
      Do While Not RsDaoAux.EOF
               
         'obtenemos el IdGrupo en SQL Server e insertamos en MovActtivoFijo
         
         IdGrupoOld = RsDaoAux("IdGrupo")   'idGrupo en el archivo Access
         IdGrupo = 0
         Q1 = "SELECT IdGrupo FROM AFGrupos WHERE AFGrupos.NombGrupo = '" & RsDaoAux("NombGrupo") & "'"
         Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then
            IdGrupo = vFld(Rs("IdGrupo"))
         End If
         Call CloseRs(Rs)
                              
         'Ahora el Insert del ActFijoFicha
         Q1 = "INSERT INTO ActFijoFicha ( IdActFijo, IdGrupo, " & Fld2 & ") VALUES(" & IdActFijoNew & "," & IdGrupo & ","
         
         For i = 0 To RsDaoAux.Fields.Count - 1
            If RsDaoAux(i).Name = "NombGrupo" Then
               Exit For
            ElseIf RsDaoAux(i).Type = dbText Or RsDaoAux(i).Type = dbMemo Or RsDaoAux(i).Type = dbChar Then
               Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDaoAux(i))) & "',"
            Else
               Q1 = Q1 & vFldDao(RsDaoAux(i)) & ","
            End If
            
         Next i
         
         Q1 = Left(Q1, Len(Q1) - 1) & ")"
         Call ExecSQL(DbMain, Q1)
         
         'Insertamos las componentes de ActiFijo de este grupo
         Fld = IdEmpresa & " As IdEmpresa, " & Ano & " As Ano, PjeDivComp, ValorCompra, ValorResidual, PjeAmortizacion, VidaUtil, CostosAdicionales, TasaDesc, CostoDesmant, ValActCostoDesmant, ValorBien, ValorRazonable_31_12, NoExisteValRazonable, OtrasDiferencias, DepAcum, VidaUtilDep, ReservaAcum, DepAcumuladaAnoAnt, VidaUtilYaDep, ReservaAcumAnt,  IdCompFichaOld, DepPeriodo, Factor, Revalorizacion,  AFComponentes.NombComp "
         Fld2 = " IdEmpresa, Ano, PjeDivComp, ValorCompra, ValorResidual, PjeAmortizacion, VidaUtil, CostosAdicionales, TasaDesc, CostoDesmant, ValActCostoDesmant, ValorBien, ValorRazonable_31_12, NoExisteValRazonable, OtrasDiferencias, DepAcum, VidaUtilDep, ReservaAcum, DepAcumuladaAnoAnt, VidaUtilYaDep, ReservaAcumAnt,  IdCompFichaOld, DepPeriodo, Factor, Revalorizacion "
      
         Q1 = "SELECT " & Fld & " FROM ActFijoCompsFicha "
         Q1 = Q1 & " INNER JOIN AFComponentes ON ActFijoCompsFicha.IdComp = AFComponentes.IdComp "
         Q1 = Q1 & " WHERE IdActFijo = " & IdActFijo ' & " AND IdGrupo = " & IdGrupoOld (no es necesario porque un ActFijo pertenece a un sólo Grupo)
         Q1 = Q1 & " ORDER BY IdActFijo "
         Set RsDaoAux2 = OpenRsDao(DbAnoAnt, Q1)
         
         Do While Not RsDaoAux2.EOF
                  
            'obtenemos el IdGrupo en SQL Server e insertamos en MovActtivoFijo
            
            idcomp = 0
            Q1 = "SELECT IdComp FROM AFComponentes WHERE AFComponentes.NombComp = '" & RsDaoAux2("NombComp") & "'"
            Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               idcomp = vFld(Rs("IdComp"))
            End If
            Call CloseRs(Rs)
                                 
            'Ahora el Insert del ActFijoCompsFicha
            Q1 = "INSERT INTO ActFijoCompsFicha ( IdActFijo, IdGrupo, IdComp, " & Fld2 & ") VALUES(" & IdActFijoNew & "," & IdGrupo & "," & idcomp & ","
            
            For i = 0 To RsDaoAux2.Fields.Count - 1
               If RsDaoAux2(i).Name = "NombComp" Then
                  Exit For
               ElseIf RsDaoAux2(i).Type = dbText Or RsDaoAux2(i).Type = dbMemo Or RsDaoAux2(i).Type = dbChar Then
                  Q1 = Q1 & "'" & ParaSQL(vFldDao(RsDaoAux2(i))) & "',"
               Else
                  Q1 = Q1 & vFldDao(RsDaoAux2(i)) & ","
               End If
               
            Next i
            
            Q1 = Left(Q1, Len(Q1) - 1) & ")"
            Call ExecSQL(DbMain, Q1)
        
            RsDaoAux2.MoveNext
        
         Loop
         
         Call CloseRs(RsDaoAux2)
         
         RsDaoAux.MoveNext
      
      Loop
      
      Call CloseRs(RsDaoAux)
      
      RsDao.MoveNext
   
   Loop
   
   Call CloseRs(RsDao)
   
   'Obtenemos cantidad de registros importados
   Q1 = "SELECT Count(*) As N "
   Q1 = Q1 & " FROM MovActivoFijo "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      NImported = vFld(Rs("N"))
   End If
   
   Call CloseRs(Rs)
   
   'actualizamos campos de ActFijoCompsFicha

   Tbl = " ActFijoCompsFicha "
   sFrom = " ActFijoCompsFicha INNER JOIN MovActivoFijo ON ActFijoCompsFicha.IdActFijo =  MovActivoFijo.IdActFijo "
   sFrom = sFrom & " AND ActFijoCompsFicha.IdEmpresa = MovActivoFijo.IdEmpresa AND ActFijoCompsFicha.Ano = MovActivoFijo.Ano "
   sSet = "  ActFijoCompsFicha.ValorBien = iif( ActFijoCompsFicha.NoExisteValRazonable <> 0, ActFijoCompsFicha.ValorBien, ActFijoCompsFicha.ValorBien * ActFijoCompsFicha.Factor)"
   sSet = sSet & ", ActFijoCompsFicha.DepAcumuladaAnoAnt = iif( ActFijoCompsFicha.NoExisteValRazonable <> 0, ActFijoCompsFicha.DepPeriodo, (ActFijoCompsFicha.DepPeriodo + ActFijoCompsFicha.DepAcum) * ActFijoCompsFicha.Factor )"
   sSet = sSet & ", ActFijoCompsFicha.VidaUtilYaDep = ActFijoCompsFicha.VidaUtilDep "
   sSet = sSet & ", ActFijoCompsFicha.ReservaAcumAnt = ActFijoCompsFicha.ReservaAcum "
   sWhere = " WHERE MovActivoFijo.FExported < 0"
   sWhere = sWhere & " AND MovActivoFijo.IdEmpresa = " & IdEmpresa & " AND MovActivoFijo.Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)


   'actualizamos campos de MovActiFijo

   Tbl = "MovActivoFijo"
   sFrom = " MovActivoFijo "
   sSet = "  IdDoc = 0 "
   sSet = sSet & ", IdComp = 0 "
   sSet = sSet & ", IdMovComp = 0 "
   sSet = sSet & ", TipoDepHist = TipoDep"

   sSet = sSet & ", DepNormalHist = iif( TipoDep = " & DEP_NORMAL & ", DepNormalHist + DepNormal, 0)"
   sSet = sSet & ", DepAceleradaHist = iif( TipoDep = " & DEP_ACELERADA & ", DepAceleradaHist + DepAcelerada, 0)"
   sSet = sSet & ", DepInstantHist = iif( TipoDep = " & DEP_INSTANTANEA & ", DepInstantHist + DepInstant, 0)"
   sSet = sSet & ", DepDecimaParteHist = iif( TipoDep = " & DEP_DECIMAPARTE & ", DepDecimaParteHist + DepDecimaParte, 0)"

   sSet = sSet & ", ValReajustadoNetoAnt = ValReajustadoNeto"
   sSet = sSet & ", DepAcumHist = DepAcumFinal"
   sWhere = Where
   sWhere = sWhere & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
   
   'actualizamos campos de depreciación en un query aparte para evitar que se pisen los cambios
   
   Q1 = "UPDATE MovActivoFijo SET "
   Q1 = Q1 & "  FImported = " & CLng(Int(Now))
   Q1 = Q1 & ", DepNormal = iif( TipoDep = " & DEP_NORMAL & ",iif( VidaUtilResidual >= 12, 12, VidaUtilResidual), 0)"
   Q1 = Q1 & ", DepAcelerada = iif( TipoDep = " & DEP_ACELERADA & ", iif( VidaUtilResidual >= 12, 12, VidaUtilResidual), 0)"
   Q1 = Q1 & ", DepInstant = iif( TipoDep = " & DEP_INSTANTANEA & ", iif( VidaUtilResidual >= 12, 12, VidaUtilResidual), 0)"
   Q1 = Q1 & ", DepDecimaParte = iif( TipoDep = " & DEP_DECIMAPARTE & ", iif( VidaUtilResidual >= 12, 12, VidaUtilResidual), 0)"
   Q1 = Q1 & ", TotalmenteDepreciado = iif(ValorLibro = 1, 1, TotalmenteDepreciado) "
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
            
   'Mensaje con cantidad
   MsgBox1 "Se importaron " & NImported & " Activos Fijos del año anterior, con vida útil residual o no depreciables.", vbInformation
   
   'limpiamos FExported en tabla nueva
   'limpiamos IdCompCent e IdCompPago que apuntan a comprobantes del año anterior
   Q1 = "UPDATE MovActivoFijo "
   Q1 = Q1 & " SET FExported = 0"
   Q1 = Q1 & Where
   Q1 = Q1 & " AND IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
  

End Function


'genera saldos de apertura de cuentas en SQL Servera partir de año anterior que está en Access
Public Function GenSaldosAperturaAccessFromSQLServer(ByVal DbAnoAnt As Database, ByVal IdEmpresa As Long, ByVal Ano As Long) As Boolean
   Dim RutMdb As String
   Dim Q1 As String
   Dim FCierre As Long
   Dim Rs As Recordset
   Dim RsDao As dao.Recordset
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String
   Dim IdCta As Long
      
   GenSaldosAperturaAccessFromSQLServer = False
         
   'limpiamos los saldos de apertura
   
   Q1 = "UPDATE Cuentas SET Debe = 0, Haber = 0, DebeTrib = 0, HaberTrib = 0 "
   Q1 = Q1 & " WHERE IdEmpresa = " & IdEmpresa & " AND Ano = " & Ano
   Call ExecSQL(DbMain, Q1)
   
   
   'generamos los saldos financieros y ambos de las cuentas y actualizamos los saldos en el plan de cuentas
   Q1 = "SELECT  MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano, Sum(MovComprobante.Debe) AS SumDebe, Sum(MovComprobante.Haber) AS SumHaber "
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta) "
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp "
   Q1 = Q1 & " WHERE (Clasificacion=" & CLASCTA_ACTIVO & " OR Clasificacion=" & CLASCTA_PASIVO & ")"
   Q1 = Q1 & " AND Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & ")"
   Q1 = Q1 & " AND Comprobante.TipoAjuste IN(" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
'   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & IdEmpresa & " AND Comprobante.Ano = " & Ano - 1   'no se requiere para año anterior en Access
   Q1 = Q1 & " GROUP BY MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano "
   Set RsDao = OpenRsDao(DbAnoAnt, Q1)
   
   Do While Not RsDao.EOF
   
      IdCta = 0
      Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & vFldDao(RsDao("Codigo")) & "'"
      Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         IdCta = vFld(Rs("IdCuenta"))
      End If
      Call CloseRs(Rs)
   
      If IdCta <> 0 Then
   
         Tbl = "Cuentas"
         sFrom = " Cuentas "
         sSet = " Cuentas.Debe = " & IIf(vFldDao(RsDao("SumDebe")) > vFldDao(RsDao("SumHaber")), vFldDao(RsDao("SumDebe")) - vFldDao(RsDao("SumHaber")), 0)
         sSet = sSet & ", Cuentas.Haber = " & IIf(vFldDao(RsDao("SumDebe")) > vFldDao(RsDao("SumHaber")), 0, vFldDao(RsDao("SumHaber")) - vFldDao(RsDao("SumDebe")))
         sWhere = " WHERE Cuentas.IdCuenta = " & IdCta & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
   
         Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
      End If
      
       RsDao.MoveNext
      
   Loop
   
   Call CloseRs(RsDao)
   
   'generamos los saldos tributarios de las cuentas y actualizamos los saldos en el plan de cuentas
   Q1 = "SELECT  MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano, Sum(MovComprobante.Debe) AS SumDebe, Sum(MovComprobante.Haber) AS SumHaber "
   Q1 = Q1 & " FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.idComp=Comprobante.idComp "
   Q1 = Q1 & " WHERE (Clasificacion=" & CLASCTA_ACTIVO & " OR Clasificacion=" & CLASCTA_PASIVO & ")"
   Q1 = Q1 & " AND Comprobante.Estado IN(" & EC_APROBADO & "," & EC_PENDIENTE & ")"
   Q1 = Q1 & " AND Comprobante.TipoAjuste IN(" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & ")"
'   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & IdEmpresa & " AND Comprobante.Ano = " & Ano - 1   'no se requiere para año anterior en Access
   Q1 = Q1 & " GROUP BY MovComprobante.IdCuenta, Cuentas.Codigo, MovComprobante.IdEmpresa, MovComprobante.Ano "
   Set RsDao = OpenRsDao(DbAnoAnt, Q1)
   
   Do While Not RsDao.EOF
   
      IdCta = 0
      Q1 = "SELECT IdCuenta FROM Cuentas WHERE Cuentas.Codigo = '" & vFldDao(RsDao("Codigo")) & "'"
      Q1 = Q1 & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         IdCta = vFld(Rs("IdCuenta"))
      End If
      Call CloseRs(Rs)
   
      If IdCta <> 0 Then
   
         Tbl = "Cuentas"
         sFrom = " Cuentas "
         sSet = " Cuentas.DebeTrib = " & IIf(vFldDao(RsDao("SumDebe")) > vFldDao(RsDao("SumHaber")), vFldDao(RsDao("SumDebe")) - vFldDao(RsDao("SumHaber")), 0)
         sSet = sSet & ", Cuentas.HaberTrib = " & IIf(vFldDao(RsDao("SumDebe")) > vFldDao(RsDao("SumHaber")), 0, vFldDao(RsDao("SumHaber")) - vFldDao(RsDao("SumDebe")))
         sWhere = " WHERE Cuentas.IdCuenta = " & IdCta & " AND Cuentas.IdEmpresa = " & IdEmpresa & " AND Cuentas.Ano = " & Ano
   
         Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
      End If
      
      RsDao.MoveNext
      
   Loop
   
   Call CloseRs(RsDao)
   
   GenSaldosAperturaAccessFromSQLServer = True
   
End Function

