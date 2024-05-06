Attribute VB_Name = "ModIFRS"
Option Explicit


'Tipos de Informes IFRS
Public Const IFRS_ESTFIN = 1
Public Const IFRS_ESTRES = 2
Public Const IFRS_BALEJEC = 3
Public Const IFRS_BAL8COL = 4


Public gInformeIFRS(IFRS_BAL8COL) As String

Public Const IFRS_MAXNIVEL = 4

#If DATACON = 1 Then       'Access

Public Function InsertTblIFRS() As Boolean
   Dim DbIFRS As Database
   Dim DbName As String
   Dim Buf As String, Rs As Recordset, SqlErr As String
   Dim ConnStr As String
   Dim TblName As String
   Dim Q1 As String
   Dim NewName As String

   On Error Resume Next

   InsertTblIFRS = True

   DbName = gDbPath & "\" & BD_IFRS

'2868088
   Call SetDbSecurity(DbName, PASSW_LEXCONT, gCfgFile, SG_SEGCFG, ConnStr)
   'Call SetDbSecurity(DbName, PASSW_LEXCONT_NEW, gCfgFile, SG_SEGCFG, ConnStr)
'fin 2868088

   ERR.Clear
   Set DbIFRS = OpenDatabase(DbName, False, False, ConnStr)

   If DbIFRS Is Nothing Then
      SqlErr = " Error " & ERR & ", '" & Error & "'"
      Buf = "Falló OpenDB: [" & DbName & "] ConnStr=" & (ConnStr <> "") & ", " & SqlErr
      Call AddLog(Buf)
   End If

   ConnStr = Mid(ConnStr, 2)  'sin el ; del principio

   If ERR = 3356 Then
      MsgBox1 "Ya existe algún usuario trabajando con la empresa seleccionada.", vbExclamation
      InsertTblIFRS = False
   End If

   If ERR = 3343 Then
      MsgBox1 "Se ha detectado fallo en la base de datos " & BD_IFRS & ", se tratará de reparar. Intente ingresar nuevamente.", vbExclamation
      Call RepairDb(DbName)
      InsertTblIFRS = False
   End If

   If (ERR Or DbMain Is Nothing) And ERR <> 3356 And ERR <> 3343 Then
      MsgBox SqlErr & vbCrLf & "'" & DbName & "'", vbExclamation
      InsertTblIFRS = False
   End If

   Call CloseDb(DbIFRS)

   'linkeamos las tablas de IFRS
   TblName = "EstadoResultado"
   NewName = "IFRS_" & TblName

   Call LinkMdbTable(DbMain, DbName, TblName, , , , ConnStr, True)

   DbMain.TableDefs.Delete NewName   ' Si ya existía, la eliminamos

   Q1 = "SELECT * INTO " & NewName & " FROM " & TblName
   Call ExecSQL(DbMain, Q1)
   Call ExecSQL(DbMain, "DROP TABLE " & TblName)

   TblName = "EstadoSituacionFinanciera"
   NewName = "IFRS_" & TblName

   Call LinkMdbTable(DbMain, DbName, TblName, , , , ConnStr, True)

   DbMain.TableDefs.Delete NewName   ' Si ya existía, la eliminamos

   Q1 = "SELECT * INTO " & NewName & " FROM " & TblName
   Call ExecSQL(DbMain, Q1)
   Call ExecSQL(DbMain, "DROP TABLE " & TblName)

End Function
Public Function InsertTblIFRS_50() As Boolean
   Dim DbIFRS As Database
   Dim DbName As String
   Dim Buf As String, Rs As Recordset, SqlErr As String
   Dim ConnStr As String
   Dim TblName As String
   Dim Q1 As String
   Dim NewName As String
   Dim sSet As String, sFrom As String, sWhere As String, Tbl As String

   On Error Resume Next

   InsertTblIFRS_50 = True

   DbName = gDbPath & "\" & BD_IFRS_50

   'Call SetDbSecurity(DbName, PASSW_LEXCONT, gCfgFile, SG_SEGCFG, ConnStr)

  '2868088
   Call SetDbSecurity(DbName, PASSW_LEXCONT, gCfgFile, SG_SEGCFG, ConnStr)
   'Call SetDbSecurity(DbName, PASSW_LEXCONT_NEW, gCfgFile, SG_SEGCFG, ConnStr)
'fin 2868088
   

   ERR.Clear
   Set DbIFRS = OpenDatabase(DbName, False, False, ConnStr)

   If DbIFRS Is Nothing Then
      SqlErr = " Error " & ERR & ", '" & Error & "'"
      Buf = "Falló OpenDB: [" & DbName & "] ConnStr=" & (ConnStr <> "") & ", " & SqlErr
      Call AddLog(Buf)
   End If

   ConnStr = Mid(ConnStr, 2)  'sin el ; del principio

   If ERR = 3356 Then
      MsgBox1 "Ya existe algún usuario trabajando con la empresa seleccionada.", vbExclamation
      InsertTblIFRS_50 = False
   End If

   If ERR = 3343 Then
      MsgBox1 "Se ha detectado fallo en la base de datos " & BD_IFRS & ", se tratará de reparar. Intente ingresar nuevamente.", vbExclamation
      Call RepairDb(DbName)
      InsertTblIFRS_50 = False
   End If

   If (ERR Or DbMain Is Nothing) And ERR <> 3356 And ERR <> 3343 Then
      MsgBox SqlErr & vbCrLf & "'" & DbName & "'", vbExclamation
      InsertTblIFRS_50 = False
   End If

   Call CloseDb(DbIFRS)

   'linkeamos las tablas de IFRS
   TblName = "PlanIFRS"
   NewName = "IFRS_" & TblName

   Call LinkMdbTable(DbMain, DbName, TblName, , , , ConnStr, True)

   DbMain.TableDefs.Delete NewName   ' Si ya existía, la eliminamos

   Q1 = "SELECT * INTO " & NewName & " FROM " & TblName
   Call ExecSQL(DbMain, Q1)
   Call ExecSQL(DbMain, "DROP TABLE " & TblName)
   
   
   'creamos los indices
   
   Q1 = "CREATE UNIQUE INDEX IdCuenta ON " & NewName & " (IdCuenta) WITH PRIMARY"
   Call ExecSQL(DbMain, Q1, False)
   
   Q1 = "CREATE UNIQUE INDEX Codigo ON " & NewName & " (Codigo)"
   Call ExecSQL(DbMain, Q1, False)

   Q1 = "CREATE UNIQUE INDEX Nombre ON " & NewName & " (Nombre)"
   Call ExecSQL(DbMain, Q1, False)
   

   
   'ahora asignamos los códigos IFRS que vienen en la tabla PlanIFRS, a los planes predefinidos
   
'   Q1 = "UPDATE PlanAvanzado INNER JOIN " & NewName & " ON PlanAvanzado.Codigo = " & NewName & ".CodPlanAvanzado"
'   Q1 = Q1 & " SET PlanAvanzado.CodIFRS = " & NewName & ".Codigo"
'   Call ExecSQL(DbMain, Q1)
   Tbl = " PlanAvanzado "
   sFrom = " PlanAvanzado INNER JOIN " & NewName & " ON PlanAvanzado.Codigo = " & NewName & ".CodPlanAvanzado"
   sSet = " PlanAvanzado.CodIFRS = " & NewName & ".Codigo"
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
'   Q1 = "UPDATE PlanIntermedio INNER JOIN " & NewName & " ON PlanIntermedio.Codigo = " & NewName & ".CodPlanAvanzado"
'   Q1 = Q1 & " SET PlanIntermedio.CodIFRS = " & NewName & ".Codigo"
'   Call ExecSQL(DbMain, Q1)
   Tbl = " PlanIntermedio "
   sFrom = " PlanIntermedio INNER JOIN " & NewName & " ON PlanIntermedio.Codigo = " & NewName & ".CodPlanAvanzado"
   sSet = " PlanIntermedio.CodIFRS = " & NewName & ".Codigo"
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
'   Q1 = "UPDATE PlanBasico INNER JOIN " & NewName & " ON PlanBasico.Codigo = " & NewName & ".CodPlanAvanzado"
'   Q1 = Q1 & " SET PlanBasico.CodIFRS = " & NewName & ".Codigo"
'   Call ExecSQL(DbMain, Q1)
   Tbl = " PlanBasico"
   sFrom = " PlanBasico INNER JOIN " & NewName & " ON PlanBasico.Codigo = " & NewName & ".CodPlanAvanzado"
   sSet = " PlanBasico.CodIFRS = " & NewName & ".Codigo"
   Call UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   
End Function

#End If

Public Function AsignarCodIFRSaPlanesPreDefinidos()
   Dim QBase As String
   Dim Wh As String
   Dim Q1 As String
   Dim CodIFRS(200) As String
   Dim CodPlan(200) As String
   Dim AuxCod As String
   Dim i As Integer
            
   'Situación Financiera
         
   CodPlan(1) = "1010101": CodIFRS(1) = "1101010"
   CodPlan(2) = "1010102": CodIFRS(2) = "1101010"
   CodPlan(3) = "1010103": CodIFRS(3) = "1101010"
   CodPlan(4) = "1010104": CodIFRS(4) = "1101010"
   CodPlan(5) = "1010203": CodIFRS(5) = "1101010"
   
   CodPlan(6) = "1010201": CodIFRS(6) = "1101020"
   CodPlan(7) = "1010202": CodIFRS(7) = "1101020"
   CodPlan(8) = "1010301": CodIFRS(8) = "1101020"
   CodPlan(9) = "1010302": CodIFRS(9) = "1101020"
   CodPlan(10) = "1010303": CodIFRS(10) = "1101020"
   
   CodPlan(11) = "1011002": CodIFRS(11) = "1101030"
   
   CodPlan(12) = "1010401": CodIFRS(12) = "1101040"
   CodPlan(13) = "1010402": CodIFRS(13) = "1101040"
   CodPlan(14) = "1010403": CodIFRS(14) = "1101040"
   CodPlan(15) = "1010404": CodIFRS(15) = "1101040"
   CodPlan(16) = "1010501": CodIFRS(16) = "1101040"
   CodPlan(17) = "1010502": CodIFRS(17) = "1101040"
   CodPlan(18) = "1010503": CodIFRS(18) = "1101040"
   CodPlan(19) = "1010504": CodIFRS(19) = "1101040"
   CodPlan(20) = "1010505": CodIFRS(20) = "1101040"
   CodPlan(21) = "1010506": CodIFRS(21) = "1101040"
   CodPlan(22) = "1010507": CodIFRS(22) = "1101040"
   CodPlan(23) = "1010508": CodIFRS(23) = "1101040"
   CodPlan(24) = "1010509": CodIFRS(24) = "1101040"
   CodPlan(25) = "1010599": CodIFRS(25) = "1101040"
   CodPlan(26) = "1010601": CodIFRS(26) = "1101040"
   CodPlan(27) = "1010602": CodIFRS(27) = "1101040"
   CodPlan(28) = "1010603": CodIFRS(28) = "1101040"
   CodPlan(29) = "1010604": CodIFRS(29) = "1101040"
   CodPlan(30) = "1010605": CodIFRS(30) = "1101040"
   CodPlan(31) = "1010606": CodIFRS(31) = "1101040"
   CodPlan(32) = "1010607": CodIFRS(32) = "1101040"
   CodPlan(33) = "1010608": CodIFRS(33) = "1101040"
   CodPlan(34) = "1010609": CodIFRS(34) = "1101040"
   CodPlan(35) = "1010610": CodIFRS(35) = "1101040"
   CodPlan(36) = "1010611": CodIFRS(36) = "1101040"
   CodPlan(37) = "1010699": CodIFRS(37) = "1101040"
   
   CodPlan(38) = "1010701": CodIFRS(38) = "1101050"
   CodPlan(39) = "1010702": CodIFRS(39) = "1101050"
   CodPlan(40) = "1010799": CodIFRS(40) = "1101050"
   
   CodPlan(41) = "1010801": CodIFRS(41) = "1101060"
   CodPlan(42) = "1010802": CodIFRS(42) = "1101060"
   CodPlan(43) = "1010803": CodIFRS(43) = "1101060"
   CodPlan(44) = "1010804": CodIFRS(44) = "1101060"
   CodPlan(45) = "1010805": CodIFRS(45) = "1101060"
   CodPlan(46) = "1010806": CodIFRS(46) = "1101060"
   CodPlan(47) = "1010898": CodIFRS(47) = "1101060"
   CodPlan(48) = "1010899": CodIFRS(48) = "1101060"
   
   CodPlan(49) = "1010901": CodIFRS(49) = "1101080"
   CodPlan(50) = "1010902": CodIFRS(50) = "1101080"
   CodPlan(51) = "1010903": CodIFRS(51) = "1101080"
   CodPlan(52) = "1010904": CodIFRS(52) = "1101080"
   CodPlan(53) = "1010905": CodIFRS(53) = "1101080"
   CodPlan(54) = "1010906": CodIFRS(54) = "1101080"
   CodPlan(55) = "1010999": CodIFRS(55) = "1101080"
   
   CodPlan(56) = "1030201": CodIFRS(56) = "1102010"
   
   CodPlan(57) = "1011001": CodIFRS(57) = "1102020"
   CodPlan(58) = "1011003": CodIFRS(58) = "1102020"
   CodPlan(59) = "1011099": CodIFRS(59) = "1102020"
   
   CodPlan(60) = "1030501": CodIFRS(60) = "1102030"
   
   CodPlan(61) = "1030601": CodIFRS(61) = "1102040"
   CodPlan(62) = "1030602": CodIFRS(62) = "1102040"
   CodPlan(63) = "1030603": CodIFRS(63) = "1102040"
   
   CodPlan(64) = "1030101": CodIFRS(64) = "1102060"
   
   CodPlan(65) = "1020305": CodIFRS(65) = "1102080"
   CodPlan(66) = "1020709": CodIFRS(66) = "1102080"
   CodPlan(67) = "1030801": CodIFRS(67) = "1102080"
   CodPlan(68) = "1030802": CodIFRS(68) = "1102080"
   CodPlan(69) = "1030803": CodIFRS(69) = "1102080"
   CodPlan(70) = "1030902": CodIFRS(70) = "1102080"
   CodPlan(71) = "1030903": CodIFRS(71) = "1102080"
   CodPlan(72) = "1031001": CodIFRS(72) = "1102080"
   CodPlan(73) = "1031002": CodIFRS(73) = "1102080"
   CodPlan(74) = "1031003": CodIFRS(74) = "1102080"
   CodPlan(75) = "1031004": CodIFRS(75) = "1102080"
   CodPlan(76) = "1031005": CodIFRS(76) = "1102080"
   CodPlan(77) = "1031006": CodIFRS(77) = "1102080"
   CodPlan(78) = "1031101": CodIFRS(78) = "1102080"
   CodPlan(79) = "1031102": CodIFRS(79) = "1102080"
   CodPlan(80) = "1031103": CodIFRS(80) = "1102080"
   CodPlan(81) = "1031105": CodIFRS(81) = "1102080"
   CodPlan(82) = "1031106": CodIFRS(82) = "1102080"
   
   CodPlan(83) = "1030301": CodIFRS(83) = "1102090"
   CodPlan(84) = "1030401": CodIFRS(84) = "1102090"
   
   CodPlan(85) = "1020101": CodIFRS(85) = "1102100"
   CodPlan(86) = "1020102": CodIFRS(86) = "1102100"
   CodPlan(87) = "1020202": CodIFRS(87) = "1102100"
   CodPlan(88) = "1020203": CodIFRS(88) = "1102100"
   CodPlan(89) = "1020299": CodIFRS(89) = "1102100"
   CodPlan(90) = "1020301": CodIFRS(90) = "1102100"
   CodPlan(91) = "1020302": CodIFRS(91) = "1102100"
   CodPlan(92) = "1020303": CodIFRS(92) = "1102100"
   CodPlan(93) = "1020304": CodIFRS(93) = "1102100"
   CodPlan(94) = "1020306": CodIFRS(94) = "1102100"
   CodPlan(95) = "1020307": CodIFRS(95) = "1102100"
   CodPlan(96) = "1020399": CodIFRS(96) = "1102100"
   CodPlan(97) = "1020401": CodIFRS(97) = "1102100"
   CodPlan(98) = "1020601": CodIFRS(98) = "1102100"
   CodPlan(99) = "1020602": CodIFRS(99) = "1102100"
   CodPlan(100) = "1020603": CodIFRS(100) = "1102100"
   CodPlan(101) = "1020604": CodIFRS(101) = "1102100"
   CodPlan(102) = "1020605": CodIFRS(102) = "1102100"
   CodPlan(103) = "1020606": CodIFRS(103) = "1102100"
   CodPlan(104) = "1020607": CodIFRS(104) = "1102100"
   CodPlan(105) = "1020608": CodIFRS(105) = "1102100"
   CodPlan(106) = "1020610": CodIFRS(106) = "1102100"
   CodPlan(107) = "1020611": CodIFRS(107) = "1102100"
   CodPlan(108) = "1020612": CodIFRS(108) = "1102100"
   CodPlan(109) = "1020613": CodIFRS(109) = "1102100"
   CodPlan(110) = "1020614": CodIFRS(110) = "1102100"
   CodPlan(111) = "1020701": CodIFRS(111) = "1102100"
   CodPlan(112) = "1020702": CodIFRS(112) = "1102100"
   CodPlan(113) = "1020703": CodIFRS(113) = "1102100"
   CodPlan(114) = "1020704": CodIFRS(114) = "1102100"
   CodPlan(115) = "1020705": CodIFRS(115) = "1102100"
   CodPlan(116) = "1020706": CodIFRS(116) = "1102100"
   CodPlan(117) = "1020707": CodIFRS(117) = "1102100"
   CodPlan(118) = "1020708": CodIFRS(118) = "1102100"
   CodPlan(119) = "1020710": CodIFRS(119) = "1102100"
   CodPlan(120) = "1020711": CodIFRS(120) = "1102100"
   CodPlan(121) = "1020712": CodIFRS(121) = "1102100"
   CodPlan(122) = "1020713": CodIFRS(122) = "1102100"
   CodPlan(123) = "1020714": CodIFRS(123) = "1102100"
   
   CodPlan(124) = "1020201": CodIFRS(124) = "1102120"
   
   CodPlan(125) = "1030702": CodIFRS(125) = "2101010"
   CodPlan(126) = "2010101": CodIFRS(126) = "2101010"
   CodPlan(127) = "2010201": CodIFRS(127) = "2101010"
   CodPlan(128) = "2010301": CodIFRS(128) = "2101010"
   CodPlan(129) = "2010401": CodIFRS(129) = "2101010"
   CodPlan(130) = "2010501": CodIFRS(130) = "2101010"
   
   CodPlan(131) = "2010601": CodIFRS(131) = "2101020"
   CodPlan(132) = "2010602": CodIFRS(132) = "2101020"
   CodPlan(133) = "2010603": CodIFRS(133) = "2101020"
   CodPlan(134) = "2010604": CodIFRS(134) = "2101020"
   CodPlan(135) = "2010605": CodIFRS(135) = "2101020"
   CodPlan(136) = "2010606": CodIFRS(136) = "2101020"
   CodPlan(137) = "2010607": CodIFRS(137) = "2101020"
   CodPlan(138) = "2010699": CodIFRS(138) = "2101020"
   CodPlan(139) = "2010902": CodIFRS(139) = "2101020"
   CodPlan(140) = "2011101": CodIFRS(140) = "2101020"
   CodPlan(141) = "2011102": CodIFRS(141) = "2101020"
   CodPlan(142) = "2011103": CodIFRS(142) = "2101020"
   CodPlan(143) = "2011104": CodIFRS(143) = "2101020"
   CodPlan(144) = "2011105": CodIFRS(144) = "2101020"
   CodPlan(145) = "2011106": CodIFRS(145) = "2101020"
   CodPlan(146) = "2011107": CodIFRS(146) = "2101020"
   CodPlan(147) = "2011108": CodIFRS(147) = "2101020"
   CodPlan(148) = "2011109": CodIFRS(148) = "2101020"
   CodPlan(149) = "2011110": CodIFRS(149) = "2101020"
   CodPlan(150) = "2011199": CodIFRS(150) = "2101020"
   CodPlan(151) = "2011203": CodIFRS(151) = "2101020"
   CodPlan(152) = "2011204": CodIFRS(152) = "2101020"
   CodPlan(153) = "2011205": CodIFRS(153) = "2101020"
   CodPlan(154) = "2011206": CodIFRS(154) = "2101020"
   CodPlan(155) = "2011299": CodIFRS(155) = "2101020"
   CodPlan(156) = "2011402": CodIFRS(156) = "2101020"
   CodPlan(157) = "2011501": CodIFRS(157) = "2101020"
   CodPlan(158) = "2011601": CodIFRS(158) = "2101020"
   
   CodPlan(159) = "2010701": CodIFRS(159) = "2101030"
   CodPlan(160) = "2010801": CodIFRS(160) = "2101030"
   
   CodPlan(161) = "2011201": CodIFRS(161) = "2101050"
   CodPlan(162) = "2011202": CodIFRS(162) = "2101050"
   CodPlan(163) = "2011301": CodIFRS(163) = "2101050"
   
   CodPlan(164) = "2011001": CodIFRS(164) = "2101060"
   CodPlan(165) = "2011003": CodIFRS(165) = "2101060"
   CodPlan(166) = "2011004": CodIFRS(166) = "2101060"
   CodPlan(167) = "2011099": CodIFRS(167) = "2101060"
   
   CodPlan(168) = "2010901": CodIFRS(168) = "2101070"
   
   CodPlan(169) = "2010903": CodIFRS(169) = "2102010"
   CodPlan(170) = "2020101": CodIFRS(170) = "2102010"
   
   CodPlan(171) = "1011202": CodIFRS(171) = "2102020"    'deberia ser 2-1-02-020
   CodPlan(172) = "2020201": CodIFRS(172) = "2102020"
   CodPlan(173) = "2020702": CodIFRS(173) = "2102020"
   CodPlan(174) = "2020801": CodIFRS(174) = "2102020"
   
   CodPlan(175) = "2020301": CodIFRS(175) = "2102030"
   CodPlan(176) = "2020401": CodIFRS(176) = "2102030"
   CodPlan(177) = "2020501": CodIFRS(177) = "2102030"
   
   CodPlan(178) = "2011401": CodIFRS(178) = "2102050"
   CodPlan(179) = "2020701": CodIFRS(179) = "2102050"
   
   CodPlan(180) = "1030701": CodIFRS(180) = "2102060"
   CodPlan(181) = "2011002": CodIFRS(181) = "2102060"
   CodPlan(182) = "2020601": CodIFRS(182) = "2102060"
   
   CodPlan(183) = "2020699": CodIFRS(183) = "2102070"
   
   CodPlan(184) = "2030101": CodIFRS(184) = "2201010"
   
   CodPlan(185) = "2031101": CodIFRS(185) = "2201020"
   CodPlan(186) = "2031201": CodIFRS(186) = "2201020"
   CodPlan(187) = "2031301": CodIFRS(187) = "2201020"
   
   CodPlan(188) = "2030201": CodIFRS(188) = "2201030"
   
   CodPlan(189) = "2030301": CodIFRS(189) = "2201040"
   CodPlan(190) = "2030401": CodIFRS(190) = "2201040"
   CodPlan(191) = "2030501": CodIFRS(191) = "2201040"
   
   QBase = "UPDATE PlanAvanzado SET CodIFRS_EstFin = '"
   Wh = " WHERE Codigo = '"
   
   For i = 1 To UBound(CodPlan)
      If CodPlan(i) = "" Then
         Exit For
      End If
      
      AuxCod = ReplaceStr(CodIFRS(i), "-", "")
      
      Q1 = QBase & AuxCod & "' " & Wh & CodPlan(i) & "'"
      Call ExecSQL(DbMain, Q1)
      
   Next i
   
   QBase = "UPDATE PlanIntermedio SET CodIFRS_EstFin = '"
   Wh = " WHERE Codigo = '"
   
   For i = 1 To UBound(CodPlan)
      If CodPlan(i) = "" Then
         Exit For
      End If
      
      AuxCod = ReplaceStr(CodIFRS(i), "-", "")
      
      Q1 = QBase & AuxCod & "' " & Wh & CodPlan(i) & "'"
      Call ExecSQL(DbMain, Q1)
      
   Next i
   
   QBase = "UPDATE PlanBasico SET CodIFRS_EstFin = '"
   Wh = " WHERE Codigo = '"
   
   For i = 1 To UBound(CodPlan)
      If CodPlan(i) = "" Then
         Exit For
      End If
      
      AuxCod = ReplaceStr(CodIFRS(i), "-", "")
      
      Q1 = QBase & AuxCod & "' " & Wh & CodPlan(i) & "'"
      Call ExecSQL(DbMain, Q1)
      
      CodIFRS(i) = ""
      CodPlan(i) = ""
   Next i
   
   
   'Estado de Resultado
   CodPlan(1) = "3010101": CodIFRS(1) = "3101010"
   CodPlan(2) = "3010102": CodIFRS(2) = "3101010"
   
   CodPlan(3) = "3010201": CodIFRS(3) = "3101020"
   CodPlan(4) = "3010202": CodIFRS(4) = "3101020"
   CodPlan(5) = "3010301": CodIFRS(5) = "3101020"
   CodPlan(6) = "3010302": CodIFRS(6) = "3101020"
   CodPlan(7) = "3010303": CodIFRS(7) = "3101020"
   CodPlan(8) = "3010304": CodIFRS(8) = "3101020"
   CodPlan(9) = "3010305": CodIFRS(9) = "3101020"
   CodPlan(10) = "3010306": CodIFRS(10) = "3101020"
   CodPlan(11) = "3010307": CodIFRS(11) = "3101020"
   CodPlan(12) = "3010308": CodIFRS(12) = "3101020"
   CodPlan(13) = "3010309": CodIFRS(13) = "3101020"
   CodPlan(14) = "3010310": CodIFRS(14) = "3101020"
   CodPlan(15) = "3010311": CodIFRS(15) = "3101020"
   CodPlan(16) = "3010312": CodIFRS(16) = "3101020"
   CodPlan(17) = "3010313": CodIFRS(17) = "3101020"
   CodPlan(18) = "3010314": CodIFRS(18) = "3101020"
   CodPlan(19) = "3010315": CodIFRS(19) = "3101020"
   CodPlan(20) = "3010401": CodIFRS(20) = "3101020"
   CodPlan(21) = "3010402": CodIFRS(21) = "3101020"
   CodPlan(22) = "3010403": CodIFRS(22) = "3101020"
   CodPlan(23) = "3010404": CodIFRS(23) = "3101020"
   CodPlan(24) = "3010405": CodIFRS(24) = "3101020"
   CodPlan(25) = "3010406": CodIFRS(25) = "3101020"
   CodPlan(26) = "3010407": CodIFRS(26) = "3101020"
   CodPlan(27) = "3010408": CodIFRS(27) = "3101020"
   CodPlan(28) = "3010409": CodIFRS(28) = "3101020"
   CodPlan(29) = "3010410": CodIFRS(29) = "3101020"
   CodPlan(30) = "3010411": CodIFRS(30) = "3101020"
   CodPlan(31) = "3010412": CodIFRS(31) = "3101020"
   CodPlan(32) = "3010413": CodIFRS(32) = "3101020"
   CodPlan(33) = "3010414": CodIFRS(33) = "3101020"
   CodPlan(34) = "3010415": CodIFRS(34) = "3101020"
   CodPlan(35) = "3010416": CodIFRS(35) = "3101020"
   CodPlan(36) = "3010417": CodIFRS(36) = "3101020"
   CodPlan(37) = "3010418": CodIFRS(37) = "3101020"
   CodPlan(38) = "3010419": CodIFRS(38) = "3101020"
   CodPlan(39) = "3010420": CodIFRS(39) = "3101020"
   CodPlan(40) = "3010421": CodIFRS(40) = "3101020"
   CodPlan(41) = "3010422": CodIFRS(41) = "3101020"
   CodPlan(42) = "3010423": CodIFRS(42) = "3101020"
   CodPlan(43) = "3010424": CodIFRS(43) = "3101020"
   CodPlan(44) = "3010425": CodIFRS(44) = "3101020"
   CodPlan(45) = "3010426": CodIFRS(45) = "3101020"
   CodPlan(46) = "3010427": CodIFRS(46) = "3101020"
   CodPlan(47) = "3010428": CodIFRS(47) = "3101020"
   CodPlan(48) = "3010429": CodIFRS(48) = "3101020"
   CodPlan(49) = "3010430": CodIFRS(49) = "3101020"
   CodPlan(50) = "3010431": CodIFRS(50) = "3101020"
   CodPlan(51) = "3010432": CodIFRS(51) = "3101020"
   CodPlan(52) = "3010433": CodIFRS(52) = "3101020"
   CodPlan(53) = "3010442": CodIFRS(53) = "3101020"
   
   CodPlan(54) = "3030101": CodIFRS(54) = "3102000"
   
   CodPlan(55) = "3010501": CodIFRS(55) = "3102030"
   CodPlan(56) = "3010502": CodIFRS(56) = "3102030"
   CodPlan(57) = "3010503": CodIFRS(57) = "3102030"
   CodPlan(58) = "3010504": CodIFRS(58) = "3102030"
   CodPlan(59) = "3010505": CodIFRS(59) = "3102030"
   CodPlan(60) = "3010506": CodIFRS(60) = "3102030"
   CodPlan(61) = "3010507": CodIFRS(61) = "3102030"
   CodPlan(62) = "3010508": CodIFRS(62) = "3102030"
   CodPlan(63) = "3010509": CodIFRS(63) = "3102030"
   CodPlan(64) = "3010510": CodIFRS(64) = "3102030"
   CodPlan(65) = "3010511": CodIFRS(65) = "3102030"
   CodPlan(66) = "3010512": CodIFRS(66) = "3102030"
   CodPlan(67) = "3010513": CodIFRS(67) = "3102030"
   CodPlan(68) = "3010514": CodIFRS(68) = "3102030"
   CodPlan(69) = "3010515": CodIFRS(69) = "3102030"
   CodPlan(70) = "3010601": CodIFRS(70) = "3102030"
   CodPlan(71) = "3010602": CodIFRS(71) = "3102030"
   CodPlan(72) = "3010603": CodIFRS(72) = "3102030"
   CodPlan(73) = "3010604": CodIFRS(73) = "3102030"
   CodPlan(74) = "3010605": CodIFRS(74) = "3102030"
   CodPlan(75) = "3010606": CodIFRS(75) = "3102030"
   CodPlan(76) = "3010607": CodIFRS(76) = "3102030"
   CodPlan(77) = "3010608": CodIFRS(77) = "3102030"
   CodPlan(78) = "3010609": CodIFRS(78) = "3102030"
   CodPlan(79) = "3010610": CodIFRS(79) = "3102030"
   CodPlan(80) = "3010611": CodIFRS(80) = "3102030"
   CodPlan(81) = "3010612": CodIFRS(81) = "3102030"
   CodPlan(82) = "3010613": CodIFRS(82) = "3102030"
   CodPlan(83) = "3010614": CodIFRS(83) = "3102030"
   CodPlan(84) = "3010615": CodIFRS(84) = "3102030"
   CodPlan(85) = "3010616": CodIFRS(85) = "3102030"
   CodPlan(86) = "3010617": CodIFRS(86) = "3102030"
   CodPlan(87) = "3010618": CodIFRS(87) = "3102030"
   CodPlan(88) = "3010619": CodIFRS(88) = "3102030"
   CodPlan(89) = "3010620": CodIFRS(89) = "3102030"
   CodPlan(90) = "3010621": CodIFRS(90) = "3102030"
   CodPlan(91) = "3010622": CodIFRS(91) = "3102030"
   CodPlan(92) = "3010623": CodIFRS(92) = "3102030"
   CodPlan(93) = "3010624": CodIFRS(93) = "3102030"
   CodPlan(94) = "3010625": CodIFRS(94) = "3102030"
   CodPlan(95) = "3010626": CodIFRS(95) = "3102030"
   CodPlan(96) = "3010627": CodIFRS(96) = "3102030"
   CodPlan(97) = "3010628": CodIFRS(97) = "3102030"
   CodPlan(98) = "3010629": CodIFRS(98) = "3102030"
   CodPlan(99) = "3010630": CodIFRS(99) = "3102030"
   CodPlan(100) = "3010631": CodIFRS(100) = "3102030"
   CodPlan(101) = "3010632": CodIFRS(101) = "3102030"
   CodPlan(102) = "3010633": CodIFRS(102) = "3102030"
   CodPlan(103) = "3010634": CodIFRS(103) = "3102030"
   CodPlan(104) = "3010635": CodIFRS(104) = "3102030"
   CodPlan(105) = "3010636": CodIFRS(105) = "3102030"
   CodPlan(106) = "3010637": CodIFRS(106) = "3102030"
   CodPlan(107) = "3010638": CodIFRS(107) = "3102030"
   CodPlan(108) = "3010639": CodIFRS(108) = "3102030"
   CodPlan(109) = "3010640": CodIFRS(109) = "3102030"
   CodPlan(110) = "3010641": CodIFRS(110) = "3102030"
   CodPlan(111) = "3010642": CodIFRS(111) = "3102030"
   CodPlan(112) = "3010643": CodIFRS(112) = "3102030"
   CodPlan(113) = "3010644": CodIFRS(113) = "3102030"
   CodPlan(114) = "3010645": CodIFRS(114) = "3102030"
   CodPlan(115) = "3010701": CodIFRS(115) = "3102030"
   
   CodPlan(116) = "3020701": CodIFRS(116) = "3102040"
   CodPlan(117) = "3020702": CodIFRS(117) = "3102040"
   CodPlan(118) = "3020703": CodIFRS(118) = "3102040"
   CodPlan(119) = "3020704": CodIFRS(119) = "3102040"
   CodPlan(120) = "3020801": CodIFRS(120) = "3102040"
   
   CodPlan(121) = "3020301": CodIFRS(121) = "3102050"
   CodPlan(122) = "3020302": CodIFRS(122) = "3102050"
   CodPlan(123) = "3020303": CodIFRS(123) = "3102050"
   
   CodPlan(124) = "3020101": CodIFRS(124) = "3102060"
   CodPlan(125) = "3020102": CodIFRS(125) = "3102060"
   CodPlan(126) = "3020103": CodIFRS(126) = "3102060"
   CodPlan(127) = "3020104": CodIFRS(127) = "3102060"
   CodPlan(128) = "3020105": CodIFRS(128) = "3102060"
   CodPlan(129) = "3020106": CodIFRS(129) = "3102060"
   
   CodPlan(130) = "3020501": CodIFRS(130) = "3102070"
   CodPlan(131) = "3020502": CodIFRS(131) = "3102070"
   CodPlan(132) = "3020503": CodIFRS(132) = "3102070"
   CodPlan(133) = "3020504": CodIFRS(133) = "3102070"
   CodPlan(134) = "3020505": CodIFRS(134) = "3102070"
   CodPlan(135) = "3020506": CodIFRS(135) = "3102070"
   CodPlan(136) = "3020601": CodIFRS(136) = "3102070"
   CodPlan(137) = "3020602": CodIFRS(137) = "3102070"
   CodPlan(138) = "3020603": CodIFRS(138) = "3102070"
   CodPlan(139) = "3020604": CodIFRS(139) = "3102070"
   
   CodPlan(140) = "3020201": CodIFRS(140) = "3102080"
   
   CodPlan(141) = "3021001": CodIFRS(141) = "3102090"
   
   CodPlan(142) = "3040101": CodIFRS(142) = "3103010"
   CodPlan(143) = "3040102": CodIFRS(143) = "3103010"
   CodPlan(144) = "3040103": CodIFRS(144) = "3103010"
   CodPlan(145) = "3050101": CodIFRS(145) = "3104010"
   CodPlan(146) = "2030301": CodIFRS(146) = "3201010"
   
   CodPlan(147) = "2030401": CodIFRS(147) = "3201020"
   
   QBase = "UPDATE PlanAvanzado SET CodIFRS_EstRes = '"
   Wh = " WHERE Codigo = '"
   
   For i = 1 To UBound(CodPlan)
      If CodPlan(i) = "" Then
         Exit For
      End If
      
      AuxCod = ReplaceStr(CodIFRS(i), "-", "")
      
      Q1 = QBase & AuxCod & "' " & Wh & CodPlan(i) & "'"
      Call ExecSQL(DbMain, Q1)
            
   Next i
   
   QBase = "UPDATE PlanIntermedio SET CodIFRS_EstRes = '"
   Wh = " WHERE Codigo = '"
   
   For i = 1 To UBound(CodPlan)
      If CodPlan(i) = "" Then
         Exit For
      End If
      
      AuxCod = ReplaceStr(CodIFRS(i), "-", "")
      
      Q1 = QBase & AuxCod & "' " & Wh & CodPlan(i) & "'"
      Call ExecSQL(DbMain, Q1)
            
   Next i
   
   QBase = "UPDATE PlanBasico SET CodIFRS_EstRes = '"
   Wh = " WHERE Codigo = '"
   
   For i = 1 To UBound(CodPlan)
      If CodPlan(i) = "" Then
         Exit For
      End If
      
      AuxCod = ReplaceStr(CodIFRS(i), "-", "")
      
      Q1 = QBase & AuxCod & "' " & Wh & CodPlan(i) & "'"
      Call ExecSQL(DbMain, Q1)
            
   Next i
  

End Function

'función que marca un comprobante para indicar si es de CCMM
'Basta que tenga alguna de las siguientes cuentas:

'3-02-09-00           Corrección Monetaria
'3-02-09-01              C.M. Activos
'3-02-09-02              C.M. Automóviles
'3-02-09-03              C.M. Activos en Leasing
'3-02-09-04              C.M. Capital Propio Financiero
'3-02-09-05              C.M. Pasivos
'3-02-09-06              C.M. Obligaciones por Leasing
'3-02-09-07              Otras Correcciones Monetarias

'Public Sub MarkCompCCMM()
'   Dim Q1 As String
'   Dim Rs As Recordset
'   Dim CurPlan As String
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS'"
'   Set Rs = OpenRs(DbMain, Q1)
'
'   If Rs.EOF = False Then
'      CurPlan = vFld(Rs("Valor"))
'   End If
'
'   Call CloseRs(Rs)
'
'   'esto es sólo para los clientes que usan nuestros planes pre-definidos
'   If CurPlan <> "BÁSICO" And CurPlan <> "INTERMEDIO" And CurPlan <> "AVANZADO" Then
'      Call CloseRs(Rs)
'      Exit Sub
'   End If
'
'
'   Call CloseRs(Rs)
'
'   Q1 = "UPDATE (Comprobante INNER JOIN MovComprobante ON Comprobante.IdComp = MovComprobante.IdComp) "
'   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta"
'   Q1 = Q1 & " SET Comprobante.EsCCMM = -1"
'   Q1 = Q1 & " WHERE Comprobante.Estado = " & EC_APROBADO & " AND left(Cuentas.Codigo, 5) = '30209'"
'
'   Call ExecSQL(DbMain, Q1)
'
'End Sub


Public Function GetCCMMApertura() As Double
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Saldo As Double
   
   Q1 = "SELECT Sum(MovComprobante.Haber - MovComprobante.Debe) as Saldo "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta"
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante")
   Q1 = Q1 & " WHERE Comprobante.Tipo = " & TC_APERTURA & " AND Cuentas.Codigo IN ('3020901', '3020902', '3020903', '3020904', '3020905', '3020906', '3020907')"
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      Saldo = vFld(Rs("Saldo"))
   End If
   Call CloseRs(Rs)
   
   GetCCMMApertura = 0   'Saldo


End Function

Public Function SaldosSinClasifIFRS() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As String
   
   SaldosSinClasifIFRS = False
      
   Q1 = "SELECT Count(*)"
   Q1 = Q1 & " FROM (MovComprobante  INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovComprobante", "Cuentas") & " )"
   Q1 = Q1 & " LEFT JOIN IFRS_PlanIFRS ON IFRS_PlanIFRS.Codigo = Cuentas.CodIFRS"
   Q1 = Q1 & " WHERE IFRS_PlanIFRS.Codigo IS NULL"
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY Cuentas.Codigo, IFRS_PlanIFRS.Codigo"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
   
      If vFld(Rs(0)) > 0 Then
         SaldosSinClasifIFRS = True
      End If
   End If
   
   Call CloseRs(Rs)
   
End Function

