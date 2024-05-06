Attribute VB_Name = "Administrador"
Option Explicit

'Public Const ID_ADMIN = 1
Global Const PRV_ADM = 1

Global Const OPER_NEW = 1
Global Const OPER_EDIT = 2
Global Const OPER_VIEW = 3
Global Const OPER_REN = 4
Global Const OPER_COPY = 5
Global Const OPER_REPAIR = 6

Public gPrtReportes As ClsPrtFlxGrid

'aplicacion actual
Public gApp     As Integer

'2850275
'Public lDbRemu As Database

'fin 2850275

Public Sub Main()
   Dim i As Integer
   Dim Msg As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Frm As FrmMain
      
   Call PamInit
    
   gDebug = GetDebug()
      
   If App.PrevInstance Then
      MsgBox "Esta aplicación ya se está ejecutando." & Chr(10) & "Use Alt+Tab hasta encontrarla", vbExclamation
      End
   End If
     
   Call AddLog("Main: 34", 1)

   Call ChkSystem(True)
    
   Call AddLog("Main: 38", 1)
    
   Call InitLexComun
   
   Call AddLog("Main: 42", 1)
   
   Debug.Print "&" & Hex(FwVersion("", 0))
   If FwVersion("", 0) >= &H20004 Then ' *** por ahora
      Call FwInit("", 8725387) ' permite que el DLL funcione
   End If
   
   Call AddLog("Main: 49", 1)
   
   gDbPath = GetCmdParam("DbPath")
   If gDbPath = "" Then
      gDbPath = W.AppPath & "\Datos"
      If APP_DEMO Then
         gDbPath = W.AppPath & "\Datos" & "Demo"
      End If
   Else
      gDbPath = ReplaceStr(gDbPath, "%AppPath%", W.AppPath)
   End If
   Call AddLog("Main 60: gDbPath=<" & gDbPath & ">", 1)
   
   On Error Resume Next
'   MkDir gDbPath & "\Importar"
'   MkDir gDbPath & "\Exportar"
'   MkDir W.AppPath & "\Log"
   
   If DB_MSSQL = False Then
      MkDir gDbPath & "\Empresas"
      Name gDbPath & "\HyperCont.mdb" As gDbPath & "\LexContab.mdb"
      gLicFile = gDbPath & "\Empresas\Info.cfg"
   Else
      gLicFile = W.AppPath & "\InfoSQL.cfg" ' 15 jul 2019: no tiene carpeta Datos
   End If
   
   ' Verificación de Inscripción del equipo
   gAppCode.Demo = True ' por defecto
   
#If Inscr2 = 0 Then
   Call FwRegist ' - Antigua inscripción
#Else
   Call AddLog("Main: 82", 1)
   ' Esquema nuevo
   Call InscribPC  ' para poder ejecutar
   Call AddLog("Main: 84", 1)
   Call CheckInscPC  ' Nueva inscripción
   
#End If

  If APP_DEMO Then
      Debug.Print "  **** SIEMPRE DEMO ****"
     
      MsgBox1 "ATENCIÓN" & vbCrLf & "Este ejecutable sólo funciona en modo Demo." & vbCrLf & "Si tiene la mantención al día, por favor, baje la actualización del programa." & vbCrLf & "Esto puede hacerlo desde el menú Ayuda o desde el sitio web.", vbInformation
     
      gAppCode.Demo = True
   End If
   
   If gAppCode.Demo Then
'      gAppCode.NivProd = VER_DEMO
      gAppCode.NivProd = VER_5EMP
      gMaxEmpLicencia = 3 'al crear las empresas se verifica que sólo sea 1-9, 2-7 y 3-5
   Else
      Select Case gAppCode.NivProd
         Case VER_ILIM
             gMaxEmpLicencia = 1000
#If DATACON = 2 Then
         Case VER_50EMP
             gMaxEmpLicencia = 50
         Case VER_100EMP
             gMaxEmpLicencia = 100
         Case VER_200EMP
             gMaxEmpLicencia = 200
         Case VER_400EMP
             gMaxEmpLicencia = 400
         Case VER_800EMP
             gMaxEmpLicencia = 800
#End If
         Case Else
             gMaxEmpLicencia = 5
      End Select
             
   End If
   
   If gAppCode.Demo Then
      Call AddLog("Version DEMO - " & APP_DEMO)
   End If
    
   gDbType = IIf(DB_MSSQL, SQL_SERVER, SQL_ACCESS)
   
'   App.Title = App.Title & "-" & IIf(gDbType = SQL_ACCESS, "Access", "SQL Server")
    
#If DATACON = 1 Then
   If OpenDbAdm() = False Then
      End
   End If
#Else
   If OpenMsSql() = False Then
      End
   End If
#End If

   gHRPath = GetCmdParam("HR")
   If gHRPath = "" Then
      i = rInStr(gDbPath, "\")
      If i Then ' asumimos que viene al final viene "\Datos"
         gHRPath = Left(gDbPath, i) & ".."
      End If
      ' gHRPath = W.AppPath & "\.."
   End If

   'si es APP_DEMO y la base no es de demo, pa' fuera para no dañar los datos con CorrigeBase
   
   If APP_DEMO Then
      
      'tiene las empresas demo? Si no, las agregamos
      Q1 = "SELECT Count(*) As N FROM Empresas WHERE RUT IN ('1','2','3')"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
      
         If vFld(Rs("N")) = 0 Then   'no hay empresas Demo, las ceramos
         
            Q1 = "INSERT INTO Empresas(Rut, NombreCorto) VALUES ('1', 'Empresa1')"
            Call ExecSQL(DbMain, Q1)
            
            Q1 = "INSERT INTO Empresas(Rut, NombreCorto) VALUES ('2', 'Empresa2')"
            Call ExecSQL(DbMain, Q1)
            
            Q1 = "INSERT INTO Empresas(Rut, NombreCorto) VALUES ('3', 'Empresa3')"
            Call ExecSQL(DbMain, Q1)
            
         End If
         
      End If
      
      Call CloseRs(Rs)
      
      'tiene más de 3 empersas con RUT distinto de 1, 2, 3
           
      Q1 = "SELECT Count(*) As N FROM Empresas WHERE RUT NOT IN ('1','2','3')"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
      
         If vFld(Rs("N")) > 0 Then
            MsgBox1 "La base de datos NO corresponde a la DEMO de LP Contabilidad.", vbCritical
            Call CloseRs(Rs)
            Call CloseDb(DbMain)
            End
         End If
         
      End If
      
      Call CloseRs(Rs)
      
   End If

   
   Call ReadOficina
   
#If DATACON = 1 Then
   Call CorrigeBaseAdm
   
  '3092471
  Call CampoImportEmpresas
   '3092471
   
#Else
   Call CorrigeBaseAdmSQLServer
   
    '3092471
  Call CampoImportEmpresas
   '3092471
#End If
      
   Call IniAdmin
   
   FrmStart.Show vbModeless
   DoEvents
   
   Sleep 1500
   
   If FrmidUsuario.FShow = vbCancel Then
      DbMain.Close
      End
   End If
   
   Call AddDebug("A show Modeless1")
   Set Frm = New FrmMain
   
   On Error GoTo 0
   
   Call AddDebug("A show Modeless2 " & (Frm Is Nothing))
   
   Frm.Show vbModeless
   
   DoEvents
   Unload FrmStart
   
End Sub
Private Sub IniAdmin()

   gAdmUser = "administ"
   gValidRut = True
   
   Set gPrtReportes = New ClsPrtFlxGrid
   gPrtReportes.PrtDemo = gAppCode.Demo
   
   Call ReadComun

   ReDim gPrivilegios(Log2(LAST_PRV))

   gPrivilegios(Log2(PRV_ADM_SIS)) = "Configurar Usuarios y Administrar Sistema"
   gPrivilegios(Log2(PRV_CFG_EMP)) = "Definir y Configurar Empresa"
   gPrivilegios(Log2(PRV_ADM_EMPRESA)) = "Administrar Períodos Contables Empresa"
   gPrivilegios(Log2(PRV_ADM_CTAS)) = "Administrar Plan de Cuentas"
   gPrivilegios(Log2(PRV_ING_COMP)) = "Ingresar Comprobantes"
   gPrivilegios(Log2(PRV_ADM_COMP)) = "Administrar Comprobantes (anular, eliminar)"
   gPrivilegios(Log2(PRV_ING_DOCS)) = "Ingresar Documentos"
   gPrivilegios(Log2(PRV_ADM_DOCS)) = "Administrar Documentos (centralizar, generar pagos)"
   gPrivilegios(Log2(PRV_ADM_DEF)) = "Administrar Entidades, Áreas de Negocio, Centros de Gestión"
   gPrivilegios(Log2(PRV_VER_INFO)) = "Ver Informes, Reportes y Libros"
   gPrivilegios(Log2(PRV_IMP_LIBOF)) = "Imprimir Libros Oficiales"
   gPrivilegios(Log2(PRV_ADM_TIMB)) = "Administrar Folios Timbraje"
   gPrivilegios(Log2(PRV_ADM_CONCIL)) = "Realizar conciliación bancaria"
   gPrivilegios(Log2(PRV_ADM_ACTFIJOS)) = "Administrar Activos Fijos"
   
   Call ReadTipoRazFin

End Sub
Public Function LoadDbList(Lst As Control, ByVal Qry As String, ByVal SelItem As Long) As Integer
   Dim i As Integer
   Dim Idx As Integer
   Dim Rs As Recordset
   
   Lst.Clear
  
   'Qry: Rs(0)=Id, Rs(1)=nombre
   Set Rs = OpenRs(DbMain, Qry)

   i = 0
   Idx = 0
   
   Do While Rs.EOF = False
      Lst.AddItem DeSQL(Rs(1))
      Lst.ItemData(Lst.NewIndex) = Rs(0)
      If SelItem > 0 And Rs(0) = SelItem Then
         Idx = i
      End If
      i = i + 1
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   Set Rs = Nothing
   
   If Lst.ListCount > 0 Then
      Lst.ListIndex = Idx
   End If
   
   LoadDbList = Idx
End Function

Public Function GetMaxTableId(ByVal IdName As String, ByVal TableName As String, ByVal Where As String) As Long
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Max(" & IdName & ") FROM " & TableName & " " & Where
   Set Rs = OpenRs(DbMain, Q1)

   If Rs.EOF = True Then
      GetMaxTableId = 1
   ElseIf IsNull(Rs(0)) Then
      GetMaxTableId = 1
   Else
      GetMaxTableId = Rs(0) + 1
   End If
   
   Call CloseRs(Rs)
   
End Function

Public Function TienePrivilegio(Priv As Long, PrivSet As Long) As Boolean
   TienePrivilegio = ((Priv And PrivSet) <> 0)
End Function

'Public Function ChkPriv(Priv As Long) As Boolean
'   ChkPriv = ((Priv And gUsuario.Priv) <> 0)
'End Function
#If DATACON <> 1 Then
Public Function ImpListEmpFromAccess() As Boolean
   Dim PathDb As String
   Dim FNBaseAccess As String
   Dim DbAccess As Database
   Dim FrmSelBase As FrmSelRuta
   Dim Q1 As String
   Dim RsDao As dao.Recordset
   Dim ConnStr As String


   'veamos si existe archivo LPContab.mdb en el path de la aplicación
   PathDb = gDbPath & "\LPContab.mdb"
   If Not ExistFile(PathDb) = True Then 'no existe archivo LPContab.mdb en Access

      'querrá seleccionar la Ruta del archivo?
      If MsgBox1("No existe archivo LPContab.mdb en la siguiente Ruta " & vbCrLf & vbCrLf & PathDb & vbCrLf & vbCrLf & "Desea seleccionar otra Ruta para el archivo Access?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
         
         MsgBox1 "No es posible traer la lista de empresas desde Access dado que no se encuentra el archivo LPContab.mdb.", vbExclamation
         Exit Function

      Else
         'permitimos seleccionar la Ruta del archivo
         Set FrmSelBase = New FrmSelRuta
         
         If FrmSelBase.FSelFile("Seleccionar base de datos Access LPContab.mdb", "Archivos MDB (*.mdb)|*.mdb", "LPContab.mdb", FNBaseAccess) = vbCancel Then
            
            Exit Function
   
         Else
            'seleccionada la ruta, veamos si existe
            PathDb = FNBaseAccess
            
            If Not ExistFile(PathDb) = True Then 'no existe archivo año anterior en Access
               MsgBox1 "No es posible abrir el archivo LPContab.mdb en la ruta indicada." & vbCrLf & vbCrLf & "Proceso de importación de empresas finalizado sin éxito.", vbExclamation
               Exit Function
      
            End If
            
         End If
         
      End If
         
   End If
   
   
   ConnStr = ";PWD=" & PASSW_LEXCONT & ";"
    
   Set DbAccess = OpenDatabase(PathDb, False, False, ConnStr)

   If Err <> 0 Or DbAccess Is Nothing Then
      MsgBox1 "No fue posible abrir el archivo LpContab.mdb en Access para esta empresa. (" & Error & ")", vbExclamation
      Exit Function
   End If

   gFrmMain.MousePointer = vbHourglass
   
   'leemos la lista de empresas
   Q1 = "SELECT Rut, NombreCorto FROM Empresas WHERE Estado = 0 "
   Set RsDao = OpenRsDao(DbAccess, Q1)
   
   Do While RsDao.EOF = False
      Q1 = "INSERT INTO Empresas (RUT, NombreCorto, Estado, RutDisp)"
      Q1 = Q1 & "VALUES ( '" & vFldDao(RsDao("Rut")) & "'"
      Q1 = Q1 & ", '" & vFldDao(RsDao("NombreCorto")) & "'"
      Q1 = Q1 & ", 0, ' ' )"
      
      
      Call ExecSQL(DbMain, Q1, False)
      
      RsDao.MoveNext
      
   Loop
   
   Call CloseRs(RsDao)
   
   '3217545 FPR SE CREO ESTE METODO PARA EL TRASPASO COMENTAR SI ES NECESARIO
   'Call CorrigeBaseAdmSQLServer
   If DbMain Is Nothing Then
    Call OpenMsSql
   End If
   Call TrasEmpresas(DbMain, DbAccess)
   ' 3217545 FPR
   
   Call CloseDb(DbAccess)
   
   gFrmMain.MousePointer = vbDefault
   
   MsgBox1 "Proceso de importación finalizado", vbInformation
   
End Function

#End If
'ImpListEmpLpRemuFromAccess
'2850275
Public Function ImpListEmpLpRemuFromAccess() As Boolean
   Dim PathDbLpRemu As String
    Dim PathDbLpContab As String
   Dim FNBaseAccess As String
   Dim DbAccess As Database
   Dim FrmSelBase As FrmSelRuta
   Dim Q1 As String
   Dim RsDao As dao.Recordset
   Dim ConnStr As String
    Dim FNBaseAccess2 As String
   Dim DbAccess2 As Database
   Dim Q2 As String
   Dim RsDao2 As dao.Recordset
   Dim ConnStr2 As String
   Dim bErrMsg As String
   
   Dim Rs1 As Recordset
   Dim Rs2 As Recordset
   
    Dim Rc As Long

  If gDbType = SQL_ACCESS Then
  

   'veamos si existe archivo LPRemu.mdb en el path de la aplicación
   PathDbLpRemu = GetIniString(gIniFile, "Config", "PathRemu", "")
   
   If Not ExistFile(PathDbLpRemu) = True Then 'no existe archivo LPContab.mdb en Access

      'querrá seleccionar la Ruta del archivo?
      If MsgBox1("No existe archivo LPRemu.mdb en la siguiente Ruta " & vbCrLf & vbCrLf & PathDbLpRemu & vbCrLf & vbCrLf & "Desea seleccionar otra Ruta para el archivo Access?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
         
         MsgBox1 "No es posible traer la lista de empresas desde Access dado que no se encuentra el archivo LPRemu.mdb.", vbExclamation
         Exit Function

      Else
         'permitimos seleccionar la Ruta del archivo
         Set FrmSelBase = New FrmSelRuta
         
         If FrmSelBase.FSelFile("Seleccionar base de datos Access LPRemu.mdb", "Archivos MDB (*.mdb)|*.mdb", "LPRemu.mdb", FNBaseAccess) = vbCancel Then
            
            Exit Function
   
         Else
            'seleccionada la ruta, veamos si existe
            PathDbLpRemu = FNBaseAccess
            
            If Not ExistFile(PathDbLpRemu) = True Then 'no existe archivo año anterior en Access
               MsgBox1 "No es posible abrir el archivo LPRemu.mdb en la ruta indicada." & vbCrLf & vbCrLf & "Proceso de importación de empresas finalizado sin éxito.", vbExclamation
               Exit Function
      
            End If
            
         End If
         
      End If
         
   End If
   
   
   ConnStr = ";PWD=" & SG_PASSW_FAIRPAY & ";"
    
   Set DbAccess = OpenDatabase(PathDbLpRemu, False, False, ConnStr)

   If Err <> 0 Or DbAccess Is Nothing Then
      MsgBox1 "No fue posible abrir el archivo LpRemu.mdb en Access para esta empresa. (" & Error & ")", vbExclamation
      Exit Function
   End If

   gFrmMain.MousePointer = vbHourglass
   
   'leemos la lista de empresas
   Q1 = "SELECT * FROM Empresas "
   Set RsDao = OpenRsDao(DbAccess, Q1)
   
   Dim contador As Integer
   
   contador = 0
   
   Do While RsDao.EOF = False
   
    'veamos si existe archivo LPContab.mdb en el path de la aplicación
   PathDbLpContab = gDbPath & "\LPContab.mdb"
   If Not ExistFile(PathDbLpContab) = True Then 'no existe archivo LPContab.mdb en Access

      'querrá seleccionar la Ruta del archivo?
      If MsgBox1("No existe archivo LPContab.mdb en la siguiente Ruta " & vbCrLf & vbCrLf & PathDbLpContab & vbCrLf & vbCrLf & "Desea seleccionar otra Ruta para el archivo Access?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
         
         MsgBox1 "No es posible traer la lista de empresas desde Access dado que no se encuentra el archivo LPContab.mdb.", vbExclamation
         Exit Function

      Else
         'permitimos seleccionar la Ruta del archivo
         Set FrmSelBase = New FrmSelRuta
         
         If FrmSelBase.FSelFile("Seleccionar base de datos Access LPContab.mdb", "Archivos MDB (*.mdb)|*.mdb", "LPContab.mdb", FNBaseAccess2) = vbCancel Then
            
            Exit Function
   
         Else
            'seleccionada la ruta, veamos si existe
            PathDbLpContab = FNBaseAccess2
            
            If Not ExistFile(PathDbLpContab) = True Then 'no existe archivo año anterior en Access
               MsgBox1 "No es posible abrir el archivo LPContab.mdb en la ruta indicada." & vbCrLf & vbCrLf & "Proceso de importación de empresas finalizado sin éxito.", vbExclamation
               Exit Function
      
            End If
            
         End If
         
      End If
         
   End If
   
   
   ConnStr2 = ";PWD=" & PASSW_LEXCONT & ";"
    
   Set DbAccess2 = OpenDatabase(PathDbLpContab, False, False, ConnStr2)

   If Err <> 0 Or DbAccess2 Is Nothing Then
      MsgBox1 "No fue posible abrir el archivo LpContab.mdb en Access para esta empresa. (" & Error & ")", vbExclamation
      Exit Function
   End If

   gFrmMain.MousePointer = vbHourglass
   
   'leemos la lista de empresas
   Q1 = "SELECT Rut, NombreCorto FROM Empresas WHERE rut = '" & vFldDao(RsDao("Rut")) & "'"
   Set RsDao2 = OpenRsDao(DbAccess2, Q1)
   
   If RsDao2.EOF = True Then
       Q1 = "INSERT INTO Empresas (RUT, NombreCorto, Estado, RutDisp)"
       Q1 = Q1 & "VALUES ( '" & vFldDao(RsDao("Rut")) & "'"
       Q1 = Q1 & ", '" & vFldDao(RsDao("RazonSoc")) & "'"
       Q1 = Q1 & ", 0, ' ' )"
              
       Call ExecSQL(DbMain, Q1, False)
         
         contador = contador + 1
    End If
    
    RsDao.MoveNext
   Call CloseRs(RsDao2)
   Call CloseDb(DbAccess2)
   Loop
   
   Call CloseRs(RsDao)
   
   Call CloseDb(DbAccess)
   
   
   Else 'sql server
   
   
  If OpenMsSqlRemu() = True Then
   
    'leemos la lista de empresas
   Q1 = "SELECT * FROM Empresas "
   Set RsDao = OpenRsDao(lDbRemu, Q1)
   
   Do While RsDao.EOF = False
   
        'leemos la lista de empresas
        Q1 = ""
        Q1 = "SELECT Rut, NombreCorto FROM Empresas WHERE rut = '" & vFldDao(RsDao("Rut")) & "'"
        Set Rs2 = OpenRs(DbMain, Q1)
        
        If Rs2.EOF = True Then
            Q1 = "INSERT INTO Empresas (RUT, NombreCorto, Estado, RutDisp)"
            Q1 = Q1 & "VALUES ( '" & vFldDao(RsDao("Rut")) & "'"
            Q1 = Q1 & ", '" & vFldDao(RsDao("RazonSoc")) & "'"
            Q1 = Q1 & ", 0, ' ' )"
                   
            Call ExecSQL(DbMain, Q1, False)
              
              contador = contador + 1
         End If
    
         RsDao.MoveNext
        Call CloseRs(Rs2)
        'Call CloseDb(DbMain)
        Loop
   
        Call CloseRs(RsDao)
        
        Call CloseDb(lDbRemu)
        
        Else
        MsgBox1 "Problemas al abrir la base de datos de Remuneraciones.", vbExclamation
    
        Exit Function
        
        End If
   
   End If
   
   gFrmMain.MousePointer = vbDefault
   
   MsgBox1 "Proceso de importación finalizado, se capturaron " & contador & " empresas desde LpRemu", vbInformation
   
End Function
