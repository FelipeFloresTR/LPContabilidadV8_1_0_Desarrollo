Attribute VB_Name = "LpContMain"
Option Explicit

Public Sub Main()
   Dim Rc As Integer, i As Integer, bDemo As Boolean
   Dim BoolIniEmpresa As Boolean
   Dim Msg As String, Key As Long, PKey As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Usr As String
      
   Call PamInit
   Call PamRandomize
      
   gEmprSeparadas = False

   ' pam: Nueva Instancia... despues de PamInit y PamRandomize
   Key = GenInstanceKey()
   PKey = Val(GetCmdParam("i"))
   ' MsgBox "Key=" & Key & " ? " & PKey
   gNuevaInstancia = (Key = PKey)

   gDebug = GetDebug()
   
   Call ChkSystem(True)

  'no se permite más de un usuario en un mismo equipo para evitar que algunos usuarios multipliquen sus licencias utilizando concección remota con Terminal Server
   If App.PrevInstance And gNuevaInstancia = False Then
      MsgBox "Esta aplicación ya se está ejecutando." & Chr(10) & "Use Alt+Tab hasta encontrarla", vbExclamation
      End
   End If

   Call InitLexComun

   Debug.Print "&" & Hex(FwVersion("", 0))
   If FwVersion("", 0) >= &H20004 Then ' *** por ahora
      Call FwInit("", 8725387) ' permite que el DLL funcione
   End If

   gDbPath = GetCmdParam("DbPath")
   If gDbPath = "" Then
      gDbPath = W.AppPath & "\Datos"
      If APP_DEMO Then
         gDbPath = W.AppPath & "\Datos" & "Demo"
      End If
   Else
      gDbPath = ReplaceStr(gDbPath, "%AppPath%", W.AppPath)
   End If
   Call AddLog("Main: gDbPath=[" & gDbPath & "]", 1)
   

   gImportPath = W.AppPath & "\Importar"
   gExportPath = W.AppPath & "\Exportar"
   
   On Error Resume Next
'   MkDir gDbPath & "\Empresas"
'   MkDir gDbPath & "\Importar"
'   MkDir gDbPath & "\Exportar"
   
   RmDir (gDbPath & "\Importar")
   RmDir (gDbPath & "\Exportar")
   
   MkDir gImportPath
   MkDir gImportPath & "\Log"
   MkDir gExportPath
   MkDir gExportPath & "\Log"
   
   MkDir W.AppPath & "\Log"
   

   ' Verificación de Inscripción del equipo
   If DB_MSSQL = False Then
      MkDir gDbPath & "\Empresas"
      Name gDbPath & "\HyperCont.mdb" As gDbPath & "\LexContab.mdb"
      gLicFile = gDbPath & "\Empresas\Info.cfg"
   Else
      gLicFile = W.AppPath & "\InfoSQL.cfg" ' 15 jul 2019: no tiene carpeta Datos
   End If
   
   gAppCode.Demo = True ' por defecto
   
#If Inscr2 = 0 Then
   Call FwRegist ' - Antigua inscripción
#Else
   ' Esquema nuevo
   Call InscribPC  ' para poder ejecutar
   Call CheckInscPC  ' Nueva inscripción
   
#End If
   
   If APP_DEMO Then
      gAppCode.Demo = True
   End If
   
   If gAppCode.Demo Then
      gAppCode.NivProd = VER_DEMO
   End If
   
   If gAppCode.Demo Then
      Call AddLog("Version DEMO - " & APP_DEMO)
   End If

   gDbType = IIf(DB_MSSQL, SQL_SERVER, SQL_ACCESS)
   
'   App.Title = App.Title & "-" & IIf(gDbType = SQL_ACCESS, "Access", "SQL Server")
  
   gEmpresa.Rut = ""
   gEmpresa.Ano = 0
   
   
   'Call CambiarClaveOLD
   
#If DATACON = 2 Then
   If OpenMsSql() = False Then
      End
   End If
#Else
   'LEO LA B.D. LEXCONTAB
   If OpenDbAdm() = False Then
      End
   End If
   
   Call SetupEmpSeparadas
#End If
   
   'si es APP_DEMO y la base no es de demo, pa' fuera para no dañar los datos con CorrigeBase
   
   If APP_DEMO Then
      
      'tiene más de 3 empersas con RUT distinto de 1, 2, 3
      
      Q1 = "SELECT Count(*) As N FROM Empresas WHERE RUT NOT IN ('1','2','3')"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
      
         If vFld(Rs("N")) > 0 Then
            MsgBox1 "La base de datos NO corresponde a la DEMO de LP Contabilidad." & vbCrLf & vbCrLf & gDbPath, vbCritical
            Call CloseRs(Rs)
            Call CloseDb(DbMain)
            End
         End If
         
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   Call AddDebug("Main: a CorrigeBaseAdm")
   
#If DATACON = 1 Then
   Call CorrigeBaseAdm
#Else
   Call CorrigeBaseAdmSQLServer

#End If
   
   Call AddDebug("Main: a InscribPC")
   
   If ContRegisterPc("", gCantLicencias) = False Then
      Call CloseDb(DbMain)
      End
   End If
   
   Usr = ContRegisteredUsr()
   
   Call AddDebug("Main: ContRegisteredUsr: '" & Usr & "'")
   
'   Q1 = "SELECT Pid FROM PcUsr WHERE PC = '" & ParaSQL(W.PcName) & "' AND Usr = '" & ParaSQL(W.UserName) & "'"
'   Set Rs = OpenRs(DbMain, Q1)
'
'   If Not Rs.EOF Then
'      Call AddDebug("Main: SELECT Pid: " & vFld(Rs("Pid")))
'   Else
'      Call AddDebug("Main: SELECT Pid: NULL")
'   End If
'
'   Call CloseRs(Rs)
      
      
   Call ReadOficina
   
   'Call CheckInscPC
   
   Call AddDebug("Main: a FrmStart.show")
   
   FrmStart.Show vbModeless
   DoEvents
   
   If gAppCode.Demo Then
      gAppCode.NivProd = VER_DEMO ' en modo DEMO mostramos todo lo del producto
      MsgBox1 "Este programa no está registrado y funcionará en modo DEMO." & vbCrLf & "Para registrarlo utilice el módulo Administrador.", vbInformation
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
   
   Sleep 500
      
   ' ****** 19 ago 2013 ************
   bDemo = gAppCode.Demo
   Call ReadPrimerUso   ' obtiene o registra el primero uso de la versión actual del programa
   If W.InDesign = False Then
   
      If FwChkActive(0) <> vbYes Then
   '      Call CloseDb(DbMain)
   '      End
      End If
   
      Call CleanPrimerUso
   End If
   
   Call AddLog("Main: DM: " & bDemo & " => " & gAppCode.Demo & " - " & APP_DEMO)
   If bDemo <> gAppCode.Demo Then   ' por si paso de No demo a Si demo        OJO VER CON PABLO   CREAR EMPRESAS 1-9, 2-7 y 3-4?
   
      Call CloseDb(DbMain)
             
      If OpenDbAdm() = False Then
         End
      End If
   End If
   ' ****** 19 ago 2013 ************
      
   Call AddDebug("Main: a SetDbPath")
      
   Call SetDbPath(FrmStart.Drive1) ' se absolutiza el gDbPath
   Call AddLog("Main: de SetDbPath ==> gDbPath=[" & gDbPath & "]", 1)
   
   gHRPath = GetCmdParam("HR")
   If gHRPath = "" Then
      i = rInStr(gDbPath, "\")
      If i Then ' asumimos que viene al final viene "\Datos"
         gHRPath = Left(gDbPath, i) & ".."
      End If
      ' gHRPath = W.AppPath & "\.."
   End If
      
   FrmIdUser.Show vbModal
   Call AddDebug("Main: después de IdUser, gUsuario.Rc=" & gUsuario.Rc)
   
   If gUsuario.Rc = vbCancel Then
      Call ContUnregisterPc(2)
      Call CloseDb(DbMain)
      End
   End If
   
   'inicializamos arreglos básicos constantes y leemos archivo Ini
   Call IniHyperCont
   Call AddDebug("Main: después IniHyperCont")
   Call IniHyperContFca
   Call AddDebug("Main: después IniHyperContFca")
   
   'Mostramos la pantalla de selección de empresas según usuario
   BoolIniEmpresa = False
   
   Do While BoolIniEmpresa = False
   
      If FrmSelEmpresas.FSelect() = vbCancel Then
         Call ContUnregisterPc(3)
         Call CloseDb(DbMain)          'OJO VER CON PABLO
         End
      End If
      
      If gEmprSeparadas Then
         'Cerramos la DB LexContab
         Call CloseDb(DbMain)
      End If
      
      'Se abre la base de datos de la empresa y se inicializan sus datos básicos
      BoolIniEmpresa = IniEmpresa()
      If BoolIniEmpresa = False Then
         If gEmprSeparadas Then
            If OpenDbAdm() = False Then
               End
            End If
         End If
      End If
      
   Loop
   
   Call AddDebug("Main: pasamos Loop IniEmpresa", 1)

   'creamos clases de impresión de grillas seteamos los datos de la empresa
   Call CreatePrtFormats
   Call SetPrtData
   
   'PS, por el error de cliente
   'Call CheckCompAPertura     'ya no es válido!
         
   Call AddDebug("Main: Vamos a FrmMain.Show", 1)
      
   FrmMain.Show vbModeless
   
   DoEvents
   Unload FrmStart
   
End Sub

Private Sub ReadPrimerUso()
   Dim Q1 As String, Rs As Recordset, Rc As Long
   
   If gAppCode.Demo Then
      Exit Sub
   End If
   
   ' Primer uso de la version actual del programa
   Q1 = "SELECT Valor FROM Param WHERE Tipo='FUVER' And Codigo=" & W.FVersion
   Set Rs = OpenRs(DbMain, Q1)
   Q1 = ""
   If Rs.EOF Then
      gAppCode.FUsoVersion = Int(Now)
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor ) VALUES ( 'FUVER', " & W.FVersion & ", '" & gAppCode.FUsoVersion & "' )"
   Else
      gAppCode.FUsoVersion = Val(vFld(Rs("Valor")))
      
      If gAppCode.FUsoVersion <= 0 Or gAppCode.FUsoVersion > Int(Now) Then
         gAppCode.FUsoVersion = Int(Now)
         Q1 = "UPDATE Param  Set Valor='" & gAppCode.FUsoVersion & "' WHERE Tipo='FUVER' And Codigo=" & W.FVersion
      End If
   End If
   Call CloseRs(Rs)
   
   If Q1 <> "" Then
      Rc = ExecSQL(DbMain, Q1)
   End If
   
End Sub

Private Sub CleanPrimerUso()
   Dim Q1 As String
   
   If gAppCode.Demo Then
      Exit Sub
   End If
   
   ' Elimina los registros antiguos, de versiones de más de
      
   Call DeleteSQL(DbMain, "Param", "WHERE Tipo='FUVER' And Codigo <" & (W.FVersion - 180), True)
      
End Sub








