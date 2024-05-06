Attribute VB_Name = "MCorrigeBaseAdm"
Option Explicit
Const MDBV21 = "Actualizav21.mdb"
Private lDbVerAdm As Integer
Dim lUpdOK As Boolean

Public Sub CorrigeBaseAdm()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Rc As Long
   
   'Call SetupEmpSeparadas
   
   
   If Not gEmprSeparadas Then
      Exit Sub
   End If


   lDbVerAdm = 0
   lUpdOK = True
      
   If Not CorrigeBaseAdm_2005_01() Then
      Exit Sub
   End If
   
   Call AddDebug("18: CorrigeBaseAdm lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V21() Then    ' 30 Sept. 2005
      Exit Sub
   End If
    
   If Not CorrigeBaseAdm_V22() Then    ' 23 Enero 2006
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V23() Then    ' 17 Marzo 2006 'PS
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V24() Then    ' 30 Marzo 2006
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V25() Then    ' 24 Abril 2006
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V26() Then    ' 25 Agosto 2006
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V27() Then    ' 11 mar 2008
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V28() Then    ' 16 may 2008
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V29() Then    ' 3 jul 2008 *** chequea que no sea
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V30() Then    ' 8 jul 2008
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V31() Then    '10 jul 2008 valores IPC
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V32() Then    '14 jul 2008
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V33() Then    '9 oct 2008
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V34() Then    '15 oct 2008
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V35() Then    '19 dic 2008
      Exit Sub
   End If
   
   Call AddDebug("80: CorrigeBaseAdm 2009 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V36() Then    '8 may 2009
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V37() Then    '12 may 2009
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V38() Then    '8 Jul 2009
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V39() Then    '21 Ago 2009
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V40() Then    '21 Sep 2009
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V41() Then    '4 dic 2009
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V42() Then    '30 dic 2009
      Exit Sub
   End If
   
   Call AddDebug("110: CorrigeBaseAdm 2010 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V43() Then    '28 ene 2010
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V44() Then    '19 ago 2010
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V45() Then    '31 ago 2010
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V46() Then    '1 sept 2010
      Exit Sub
   End If
   
   Call AddDebug("128: CorrigeBaseAdm 2011 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V47() Then    '29 mar 2011
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V48() Then    '15 abril 2011
      Exit Sub
   End If
   
   'aqui se da inicio a versión 3.0
   
   If Not CorrigeBaseAdm_V299() Then    '13 julio 2011
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V300() Then    '10 agosto 2011
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V301() Then    '18 agosto 2011
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V302() Then    '6 sept. 2011
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V303() Then    '23 nov. 2011
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V304() Then    '24 nov. 2011
      Exit Sub
   End If
   
   Call AddDebug("164: CorrigeBaseAdm 2012 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V305() Then    '16 abr 2012
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V306() Then    '26 sep 2012     Tablas IFRS
      Exit Sub
   End If

   If Not CorrigeBaseAdm_V307() Then    '26 sep 2012
      Exit Sub
   End If

   Call AddDebug("178: CorrigeBaseAdm 2013 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V308() Then    '6 mayo 2013
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V309() Then    '7 mayo 2013
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V310() Then    '7 mayo 2013
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V311() Then    '30 mayo 2013
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V312() Then     '31 jul 2013
      Exit Sub
   End If
   
   Call AddDebug("200: CorrigeBaseAdm 2013-2 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V313() Then     '19 ago 2013
      Exit Sub
   End If

   Call AddDebug("206: CorrigeBaseAdm 2013 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V314() Then     'entregada  10 sep 2013
      Exit Sub
   End If

   Call AddDebug("212: CorrigeBaseAdm 2013 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V315() Then     'entregada  7 oct 2013
      Exit Sub
   End If
   
   Call AddDebug("218: CorrigeBaseAdm 2013 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V316() Then     'entregada 13 nov 2013
      Exit Sub
   End If
   
   Call AddDebug("224: CorrigeBaseAdm 2013 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V317() Then     'agregada 19 nov 2013
      Exit Sub
   End If
      
   Call AddDebug("230: CorrigeBaseAdm 2014 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V318() Then     'entregada 26 oct 2014
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V319() Then     'entregada 6 nov 2014
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V320() Then     'entregada 11 nov 2014
      Exit Sub
   End If
   
   Call AddDebug("234: CorrigeBaseAdm 2015 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V321() Then     'entregada 6 mar 2015
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V322() Then     'entregada 10 mar 2015
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V323() Then     'entregada abril 2015
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V324() Then     'entregada 28 mayo 2015
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V325() Then     'agregada 17 junio 2015
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V326() Then     'agregada 19 nov 2015
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V327() Then     'entregada 26 nov 2015
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V328() Then     'entregada 2 mar 2015
      Exit Sub
   End If
      
   Call AddDebug("270: CorrigeBaseAdm 2016 lDbVerAdm=" & lDbVerAdm)
   
   If Not CorrigeBaseAdm_V329() Then     'agregada 20 jul 2016 / entregada
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V330() Then     'agregada 21 nov 2016 / entregada
      Exit Sub
   End If
      
   Call AddDebug("336: CorrigeBaseAdm 2017 lDbVerAdm=" & lDbVerAdm)
      
   If Not CorrigeBaseAdm_V331() Then     'agregada 25 ene 2017 / entregada
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V332() Then     'agregada 30 ene 2017 / entregada
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V333() Then     'agregada 31 ene 2017 / entregada
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V334() Then     'agregada 2 feb 2017 / entregada
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V335() Then     'agregada 10 mar 2017 / entregada 11 mar 2017
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V336() Then     'agregada 28 mar 2017 / entregada ?
      Exit Sub
   End If
      
      
   If Not CorrigeBaseAdm_V337() Then     'agregada y entregada 5 abr 2017
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V338() Then     'agregada 10 abr 2017
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V339() Then     'agregada 17 may 2017
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V340() Then     'agregada 1 jun 2017
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V341() Then     'agregada 16 jun 2017
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V342() Then     'agregada 20 jun 2017
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V343() Then     'agregada 23 ago 2017
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V344() Then     'agregada 14 sept 2017
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V345() Then     'agregada 13 nov 2017
      Exit Sub
   End If
      
      
   If Not CorrigeBaseAdm_V346() Then     'agregada 28 nov 2017
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V347() Then     'agregada 19 dic 2017
      Exit Sub
   End If
            
   If Not CorrigeBaseAdm_V348() Then     'agregada 28 dic 2017
      Exit Sub
   End If
      
   Call AddDebug("336: CorrigeBaseAdm 2018 lDbVerAdm=" & lDbVerAdm)
      
   If Not CorrigeBaseAdm_V349() Then     'agregada 5 abr 2018
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V350() Then     'agregada 6 jul 2018
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V351() Then     'agregada 17 oct 2018
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V352() Then     'agregada 31 oct 2018
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V353() Then     'agregada 15 nov 2018
      Exit Sub
   End If
            
   If Not CorrigeBaseAdm_V354() Then     'agregada 17 ene 2019
      Exit Sub
   End If
            
   If Not CorrigeBaseAdm_V355() Then     'agregada 2 sep 2019
      Exit Sub
   End If
            
   If Not CorrigeBaseAdm_V356() Then     'agregada 24 dic 2019
      Exit Sub
   End If
            
   If Not CorrigeBaseAdm_V357() Then     'agregada 16 mar 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V358() Then     'agregada 31 mar 2020
      Exit Sub
   End If
            
   If Not CorrigeBaseAdm_V359() Then     'agregada 30 jul 2020
      Exit Sub
   End If
            
   If Not CorrigeBaseAdm_V360() Then     'agregada 24 ago 2020
      Exit Sub
   End If
            
   If Not CorrigeBaseAdm_V361() Then     'agregada 31 ago 2020
      Exit Sub
   End If
            
   If Not CorrigeBaseAdm_V362() Then     'agregada 2 sep 2020
      Exit Sub
   End If
            
   If Not CorrigeBaseAdm_V363() Then     'agregada 14 sep 2020
      Exit Sub
   End If
            
   If Not CorrigeBaseAdm_V364() Then     'agregada 17 sep 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V365() Then     'agregada 20 oct 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V366() Then     'agregada 2 nov 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V367() Then     'agregada 30 nov 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V368() Then     'agregada 12 mar 2021
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V369() Then     'agregada 14 abr 2021
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V370() Then     'agregada 10 sep 2021 gcb10092021
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V371() Then     'agregada 5 oct 2021 gcb10092021
      Exit Sub
   End If
   
   '2814014 pipe
   If Not CorrigeBaseAdm_V372() Then     'agregada 27 may 2022 ffv 2814014
      Exit Sub
   End If
   'fin 2814014
   
  
   
   
   
'   If lDbVerAdm > 44 Then
'      MsgBox1 "¡ ATENCION !" & vbCrLf & vbCrLf & "La base de datos corresponde a una versión posterior de este programa." & vbCrLf & "Debe actualizar el programa antes de continuar, de lo contrario podría dañar la información..", vbCritical
'      Call CloseDb(DbMain)
'      End
'   End If
   
   Call AddDebug("350: CorrigeBaseAdm FIN lDbVerAdm=" & lDbVerAdm)
   
#If DATACON = 1 Then
   Q1 = "SELECT Codigo FROM Param WHERE Tipo='VERSION'"
   Set Rs = OpenRs(DbMain, Q1)
   Rc = Rs.EOF
   Call CloseRs(Rs)
   
   If Rc Then
      Call AlterField(DbMain, "Param", "Codigo", dbLong)
   
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor ) VALUES ('VERSION', " & W.FVersion & ", '" & ParaSQL(W.Version) & "' )"
   Else
      Q1 = "UPDATE Param SET Codigo= " & W.FVersion & ", Valor= '" & ParaSQL(W.Version) & "' WHERE Tipo='VERSION'"
   End If
   Rc = ExecSQL(DbMain, Q1)
#End If
  
End Sub


Public Function CorrigeBaseAdm_V371() As Boolean   'agregada 4 oct 2021
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String


   On Error Resume Next

   '--------------------- Versión 371 -----------------------------------
   
   '2803829
   ActImpAdicionales2016
   'fin 2803829

   If lDbVerAdm = 371 And lUpdOK = True Then
           
      'Insertamos Retención 3% préstamo solidario a libro de retenciones
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
      Q1 = Q1 & " VALUES(" & LIB_RETEN & "," & LIBRETEN_RET3PORC & ", 'Retención 3%', ' ', ' ', 0, ' ', ' ', 'Retención 3%',' ',' ', 5, 0, 0, ' ', 'Retención 3% Prést. Sol.')"
      Call ExecSQL(DbMain, Q1)
      
      
           
     If lUpdOK Then
         lDbVerAdm = 372
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V371 = lUpdOK
   
End Function

Private Function CorrigeBaseAdm_V370() As Boolean
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim Tb As TableDef


   On Error Resume Next
   
   'gcb10092021
   
   
   
   
   '--------------------- Versión 370 -----------------------------------
   If lDbVerAdm = 370 And lUpdOK = True Then
   
      ' no hace se hace acá sino en el CoprrigeBase   FCA 29 sep 2021
'      Q1 = "CREATE TABLE AjusteIVAMensual ( IDEmpresa Long, Ano Integer, "
'      Q1 = Q1 & "Mes Byte,Valor Double )"
'      Call ExecSQL(DbMain, Q1)
'      DbMain.TableDefs.Refresh
     
                          
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 371
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V370 = lUpdOK

   
End Function

Public Function CorrigeBaseAdm_V369() As Boolean   'agregada 14 abr 2021
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String


   On Error Resume Next

   '--------------------- Versión 369 -----------------------------------

   If lDbVerAdm = 369 And lUpdOK = True Then
           
      'Insertamos impuestos específicos Diesel y Gasolina en libro de Ventas
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
      Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IMPDIESEL & ", 'Impuesto Específico Diesel', ' ', ' ', 0, ' ', ',1,3,4,', 'Imto. Esp.','Diesel',' ', 44, 100, 0, '28', 'Impuesto Específico Diesel')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
      Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IMPGASOLINA & ", 'Impuesto Específico Gasolina', ' ', ' ', 0, ' ', ',1,3,4,', 'Imto. Esp.','Gasolina',' ', 45, 100, 0, '35', 'Impuesto Específico Gasolina')"
      Call ExecSQL(DbMain, Q1)
      
           
     If lUpdOK Then
         lDbVerAdm = 370
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V369 = lUpdOK
   
End Function


Public Function CorrigeBaseAdm_V368() As Boolean   'agregada 12 mar 2021
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String


   On Error Resume Next

   '--------------------- Versión 368 -----------------------------------

   If lDbVerAdm = 368 And lUpdOK = True Then
            
      Set Tbl = DbMain.TableDefs("EmpresasAno")

      'agregamos campo CPS_INRPropiosPerdidas a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_INRPropiosPerdidas", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_INRPropiosPerdidas", vbExclamation
         lUpdOK = False

      End If
      
      'agregamos campo CPS_UtilidadesPerdida a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_UtilidadesPerdida", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_UtilidadesPerdida", vbExclamation
         lUpdOK = False

      End If
   
      'agregamos campo CPS_IngresoDiferido a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_IngresoDiferido", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_IngresoDiferido", vbExclamation
         lUpdOK = False

      End If
           
       'agregamos campo CPS_CTDImputableIPE a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_CTDImputableIPE", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_CTDImputableIPE", vbExclamation
         lUpdOK = False

      End If
           
      'agregamos campo CPS_IncentivoAhorro a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_IncentivoAhorro", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_IncentivoAhorro", vbExclamation
         lUpdOK = False

      End If
           
      'agregamos campo CPS_IDPCVoluntario a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_IDPCVoluntario", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_IDPCVoluntario", vbExclamation
         lUpdOK = False

      End If
           
      'agregamos campo CPS_CredActFijos a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_CredActFijos", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_CredActFijos", vbExclamation
         lUpdOK = False

      End If
           
      'agregamos campo CPS_CredParticipaciones a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_CredParticipaciones", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_CredParticipaciones", vbExclamation
         lUpdOK = False

      End If
           
     If lUpdOK Then
         lDbVerAdm = 369
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V368 = lUpdOK
   
End Function


Public Function CorrigeBaseAdm_V367() As Boolean   'agregada 30 nov 2020
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim QBase As String, QEnd As String


   On Error Resume Next

   '--------------------- Versión 367 -----------------------------------

   If lDbVerAdm = 367 And lUpdOK = True Then
            
      'agregamos campo AnoDesde a PlanCuentasSII  para indicar desde que año está vigente una cuenta
      Set Tbl = DbMain.TableDefs("PlanCuentasSII")

      Err.Clear
      Set Fld = Tbl.CreateField("AnoDesde", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanCuentasSII.AnoDesde", vbExclamation
         lUpdOK = False

      End If
      
      'agregamos dos cuentas adicionales a plan de cuentas SII
      QBase = "INSERT INTO PlanCuentasSII (CodigoSII, DescripSII, FmtCodigoSII, Clasificacion, AnoDesde) VALUES("
      QEnd = "," & CLASCTA_ACTIVO & ", 2020)"
     
      Q1 = "'1016300', 'Gastos Anticipados', '1.01.63.00'"
      Call ExecSQL(DbMain, QBase & Q1 & QEnd)

      Q1 = "'1016400', 'Impuesto Diferido', '1.01.64.00'"
      Call ExecSQL(DbMain, QBase & Q1 & QEnd)

          
      If lUpdOK Then
         lDbVerAdm = 368
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V367 = lUpdOK
   
End Function


Public Function CorrigeBaseAdm_V366() As Boolean   'agregada 2 nov 2020
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String


   On Error Resume Next

   '--------------------- Versión 366 -----------------------------------

   If lDbVerAdm = 366 And lUpdOK = True Then
            
      Set Tbl = DbMain.TableDefs("EmpresasAno")

      'agregamos campo CPS_CapPropioTrib a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_CapPropioTrib", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_CapPropioTrib", vbExclamation
         lUpdOK = False

      End If
      
      'agregamos campo CPS_CapPropioTribAnoAnt a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_CapPropioTribAnoAnt", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_CapPropioTribAnoAnt", vbExclamation
         lUpdOK = False

      End If
           
      'agregamos campo CPS_RepPerdidaArrastre a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_RepPerdidaArrastre", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_RepPerdidaArrastre", vbExclamation
         lUpdOK = False

      End If
      
      'agregamos campo CPS_CapPropioSimplVarAnual a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_CapPropioSimplVarAnual", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_CapPropioSimplVarAnual", vbExclamation
         lUpdOK = False

      End If
          
      If lUpdOK Then
         lDbVerAdm = 367
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V366 = lUpdOK
   
End Function


Public Function CorrigeBaseAdm_V365() As Boolean   'agregada 20 oct 2020
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String


   On Error Resume Next

   '--------------------- Versión 365 -----------------------------------

   If lDbVerAdm = 365 And lUpdOK = True Then
            
      Set Tbl = DbMain.TableDefs("EmpresasAno")

      'agregamos campo CPS_AumentosCapital a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_AumentosCapital", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_AumentosCapital", vbExclamation
         lUpdOK = False

      End If
      
      'agregamos campo CPS_GastosRechazadosNoPagan40 a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_GastosRechazadosNoPagan40", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_GastosRechazadosNoPagan40", vbExclamation
         lUpdOK = False

      End If
      
      'agregamos campo CPS_INRPropios a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_INRPropios", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_INRPropios", vbExclamation
         lUpdOK = False

      End If
      
       'agregamos campo CPS_OtrosAjustesAumentos a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_OtrosAjustesAumentos", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_OtrosAjustesAumentos", vbExclamation
         lUpdOK = False

      End If
     
       'agregamos campo CPS_OtrosAjustesDisminuciones a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_OtrosAjustesDisminuciones", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_OtrosAjustesDisminuciones", vbExclamation
         lUpdOK = False

      End If
      
           
      If lUpdOK Then
         lDbVerAdm = 366
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V365 = lUpdOK
   
End Function



Public Function CorrigeBaseAdm_V364() As Boolean   'agregada 17 sep 2020
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String


   On Error Resume Next

   '--------------------- Versión 364 -----------------------------------

   If lDbVerAdm = 364 And lUpdOK = True Then
      
      'Tipo IVA Retenido en Libro de Compras
      
      Set Tbl = DbMain.TableDefs("EmpresasAno")

      'agregamos campo CPS_CapitalAportado a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_CapitalAportado", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_CapitalAportado", vbExclamation
         lUpdOK = False

      End If
      
      'agregamos campo CPS_BaseImpPrimCat 14DN3 a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_BaseImpPrimCat_14DN3", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_BaseImpPrimCat_14DN3", vbExclamation
         lUpdOK = False

      End If
      
      'agregamos campo CPS_BaseImpPrimCat 14DN8 a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_BaseImpPrimCat_14DN8", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_BaseImpPrimCat_14DN8", vbExclamation
         lUpdOK = False

      End If
      
       'agregamos campo CPS_Participaciones a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_Participaciones", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_Participaciones", vbExclamation
         lUpdOK = False

      End If
     
       'agregamos campo CPS_Disminuciones a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_Disminuciones", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_Disminuciones", vbExclamation
         lUpdOK = False

      End If
      
       'agregamos campo CPS_GastosRechazados a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_GastosRechazados", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_GastosRechazados", vbExclamation
         lUpdOK = False

      End If
      
       'agregamos campo CPS_RetirosDividendos a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_RetirosDividendos", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_RetirosDividendos", vbExclamation
         lUpdOK = False

      End If
      
        'agregamos campo CPS_CapPropioSimplificado a EmpresasAno
      Err.Clear
      Set Fld = Tbl.CreateField("CPS_CapPropioSimplificado", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CPS_CapPropioSimplificado", vbExclamation
         lUpdOK = False

      End If
      
      'Agregamos tabla CapPropioSimplAnual
      
      Set Tbl = New TableDef
      Tbl.Name = "CapPropioSimplAnual"

      Err.Clear
      Set Fld = Tbl.CreateField("IdCapPropioSimplAnual", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "CapPropioSimplAnual.IdCapPropioSimplAnual", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set Fld = Tbl.CreateField("IdEmpresa", dbLong)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "CapPropioSimplAnual.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set Fld = Tbl.CreateField("TipoDetCPS", dbByte)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "CapPropioSimplAnual.TipoDetCPS", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set Fld = Tbl.CreateField("IngresoManual", dbBoolean)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "CapPropioSimplAnual.IngresoManual", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set Fld = Tbl.CreateField("AnoValor", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "CapPropioSimplAnual.AnoValor", vbExclamation
         lUpdOK = False
      End If

      Err.Clear
      Set Fld = Tbl.CreateField("Valor", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "CapPropioSimplAnual.Valor", vbExclamation
         lUpdOK = False
      End If

      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON CapPropioSimplAnual (IdCapPropioSimplAnual) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE INDEX IdxEmpAno ON CapPropioSimplAnual (IdEmpresa, TipoDetCPS, AnoValor )"
         Rc = ExecSQL(DbMain, Q1, False)
      

      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla CapPropioSimplAnual", vbExclamation
         lUpdOK = False

      End If
          
           
      If lUpdOK Then
         lDbVerAdm = 365
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V364 = lUpdOK
   
End Function


Public Function CorrigeBaseAdm_V363() As Boolean   'agregada 14 sep 2020
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String


   On Error Resume Next

   '--------------------- Versión 363 -----------------------------------

   If lDbVerAdm = 363 And lUpdOK = True Then
      
      'Tipo IVA Retenido en Libro de Compras
      
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_TOTAL & " WHERE TipoLib = " & LIB_COMPRAS
      Q1 = Q1 & " AND Codigo = " & LIBCOMPRAS_IVARETTOT
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_PARCIAL & " WHERE TipoLib = " & LIB_COMPRAS
      Q1 = Q1 & " AND Codigo = " & LIBCOMPRAS_IVARETPARC
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVerAdm = 364
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V363 = lUpdOK
   
End Function

Public Function CorrigeBaseAdm_V362() As Boolean   'agregada 2 sep 2020
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String


   On Error Resume Next

   '--------------------- Versión 362 -----------------------------------

   If lDbVerAdm = 362 And lUpdOK = True Then
      
      'Tipo IVA Retenido en Libro de Compras
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_PARCIAL & " WHERE TipoLib = " & LIB_COMPRAS
      Q1 = Q1 & " AND Codigo >=  " & LIBCOMPRAS_IVARETPARCTRIGO & " AND Codigo <= " & LIBCOMPRAS_IVARETPARCFAMBPASAS
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_TOTAL & " WHERE TipoLib = " & LIB_COMPRAS
      Q1 = Q1 & " AND ((Codigo >=  " & LIBCOMPRAS_IVARETTOTCHATARRA & " AND Codigo <= " & LIBCOMPRAS_IVARETTOTCARTONES & ")"
      Q1 = Q1 & " OR Codigo =  " & LIBCOMPRAS_IVARETORO & ")"
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVerAdm = 363
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V362 = lUpdOK
   
End Function

Public Function CorrigeBaseAdm_V361() As Boolean   'agregada 31 ago 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 361 -----------------------------------

   If lDbVerAdm = 361 And lUpdOK = True Then
      
      'Corrección de texto Título Chatarra
      Q1 = "UPDATE TipoValor SET TitCompleto = 'IVA Retenido Total Chatarra' WHERE TipoLib = " & LIB_VENTAS
      Q1 = Q1 & " AND Codigo = " & LIBVENTAS_IVA_RETTOTAL_CHATARRA
      
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVerAdm = 362
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V361 = lUpdOK
   
End Function


Public Function CorrigeBaseAdm_V360() As Boolean   'agregada 24 ago 2020
   Dim Q1 As String, Q2 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 360 -----------------------------------

   If lDbVerAdm = 360 And lUpdOK = True Then
      
'      'código SII DTE para TipoDoc FIV
'      Q1 = "UPDATE TipoDocs SET CodDocSII = '29' WHERE Diminutivo = 'FIV'"
'      Call ExecSQL(DbMain, Q1)
'
'      'código SII DTE para TipoDoc FIV
'      Q1 = "UPDATE TipoDocs SET CodDocSII = '0', CodDocDTESII = '47' WHERE Diminutivo = 'VPE'"
'      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo TipoIVARetenido a tabla TipoValor
      Set Tbl = DbMain.TableDefs("TipoValor")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoIVARetenido", dbByte)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.TipoIVARetenido", vbExclamation
         lUpdOK = False
      End If
      
      'Dejamos TAsa en NULL donde vale 0
      Q1 = "UPDATE TipoValor SET Tasa = ' ' WHERE Tasa = 0"
      Call ExecSQL(DbMain, Q1)
      
      'actualizamos códigos SII de algunos impuestos adicionales
      Q1 = "UPDATE TipoValor SET CodSIIDTE = '15' WHERE Codigo = " & LIBVENTAS_IVARETTOT
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoValor SET CodSIIDTE = '14' WHERE Codigo = " & LIBVENTAS_RETMARGENCOM
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET CodSIIDTE = '271' WHERE Codigo = " & LIBVENTAS_ILABEDANALCAZUCAR
      Call ExecSQL(DbMain, Q1)

     
      'agregamos nuevos impuestos adicionales para Vnetas
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, "
      Q1 = Q1 & " Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto, TipoIVARetenido)"
      Q1 = Q1 & " VALUES(" & LIB_VENTAS & ","

      Q2 = Q1 & LIBVENTAS_IVA_ANTICIP_FAENACARNE & ", 'IVA Anticip. Faenam. Carne', ' ', ' ', 0, ' ', ',3,4,5,', "
      Q2 = Q2 & "'IVA Anticip.','Faenamiento Carne',' ', 23, 5, 0, '17', 'IVA Anticipado Faenamiento Carne', 0)"

      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_LEGUMBRES & ", 'IVA Ret. Parcial Legumbres', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Legumbres',' ', 24, ' ', 0, '30', 'IVA Retenido Parcial Legumbres', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_LEGUMBRES & ", 'IVA Ret. Total Legumbres' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Legumbres',' ', 25, 100, 0, '301', 'IVA Retenido Total Legumbres', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_SILVESTRES & ", 'IVA Ret. Total Silvestres' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Silvestres',' ', 26, 100, 0, '31', 'IVA Retenido Total Silvestres', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_GANADO & ", 'IVA Ret. Parcial Ganado' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Ganado',' ', 27, ' ', 0, '32', 'IVA Retenido Parcial Ganado', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_GANADO & ", 'IVA Ret. Total Ganado' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Ganado',' ', 28, 100, 0, '321', 'IVA Retenido Total Ganado', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_MADERA & ", 'IVA Ret. Parcial Madera' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Madera',' ', 29, ' ', 0, '33', 'IVA Retenido Parcial Madera', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_MADERA & ", 'IVA Ret. Total Madera' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Madera',' ', 30, 100, 0, '331', 'IVA Retenido Total Madera', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_TRIGO & ", 'IVA Ret. Parcial Trigo' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Trigo',' ', 31, ' ', 0, '34', 'IVA Retenido Parcial Trigo', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_TRIGO & ", 'IVA Ret. Total Trigo' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Trigo',' ', 32, 100, 0, '341', 'IVA Retenido Total Trigo', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_ARROZ & ", 'IVA Ret. Parcial Arroz' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Arroz',' ', 33, ' ', 0, '36', 'IVA Retenido Parcial Arroz', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_ARROZ & ", 'IVA Ret. Total Arroz' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Arroz',' ', 34, 100, 0, '361', 'IVA Retenido Total Arroz', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_HIDROBIOLOGICAS & ", 'IVA Ret. Parcial Hidrobiológicas', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Hidrobiológicas',' ', 35, ' ', 0, '37', 'IVA Retenido Parcial Hidrobiológica',  " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_HIDROBIOLÓGICAS & ", 'IVA Ret. Total Hidrobiológicas' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Hidrobiológicas',' ', 36, 100, 0, '371', 'IVA Retenido Total Hidrobiológica', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_CHATARRA & ", 'IVA Ret. Total Chatarra', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Chatarra',' ', 37, 100, 0, '38', 'IVA Retenido Total Chatarrae', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_PPA & ", 'IVA Ret. Total PPA', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','PPA',' ', 38, 100, 0, '39', 'IVA Retenido Total PPA', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_CARTONES & ", 'IVA Ret. Total Cartones', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Cartones',' ', 39, 100, 0, '47', 'IVA Retenido Total Cartones', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_BERRIES & ", 'IVA Ret. Parcial Berries', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Berries',' ', 40, ' ', 0, '48', 'IVA Retenido Parcial Berries', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_BERRIES & ", 'IVA Ret. Total Berries', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Berries',' ', 41, 100, 0, '481', 'IVA Retenido Total Berries', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_FACT_COMPRA_SIN_RET & ", 'Fact. compra sin Retención', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'Factura de compra','sin Retención',' ', 42, 0, 0, '49', 'Factura de compra sin Retención', 0)"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RET_FACT_INICIO & ", 'IVA Retenido Factura de Inicio', ' ', ' ', 0, ' ', ',17,',"
      Q2 = Q2 & "'IVA Ret.','Fact. Inicio',' ', 43, 100, 0, '60', 'IVA Retenido Factura de Inicio', " & IVARET_TOTAL & ")"

      Call ExecSQL(DbMain, Q2)
      
      
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_TOTAL & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo IN ( " & LIBVENTAS_IVARETTOT & "," & LIBVENTAS_IVAADQCONSTINMUEBLES & ") "
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_PARCIAL & " WHERE  TipoLib = " & LIB_VENTAS & " AND Codigo IN ( " & LIBVENTAS_IVARETPARC & ") "
      Call ExecSQL(DbMain, Q1)
      
      
      If lUpdOK Then
         lDbVerAdm = 361
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V360 = lUpdOK
   
End Function

Public Function CorrigeBaseAdm_V359() As Boolean   'agregada 30 jul 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 359 -----------------------------------

   If lDbVerAdm = 359 And lUpdOK = True Then
      
      'código SII DTE para TipoDoc FIV
      Q1 = "UPDATE TipoDocs SET CodDocSII = '29' WHERE Diminutivo = 'FIV'"
      Call ExecSQL(DbMain, Q1)
      
      'código SII DTE para TipoDoc FIV
      Q1 = "UPDATE TipoDocs SET CodDocSII = '0', CodDocDTESII = '47' WHERE Diminutivo = 'VPE'"
      Call ExecSQL(DbMain, Q1)
      
      
      If lUpdOK Then
         lDbVerAdm = 360
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V359 = lUpdOK
   
End Function

Public Function CorrigeBaseAdm_V358() As Boolean   'agregada 31 mar 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 358 -----------------------------------

   If lDbVerAdm = 358 And lUpdOK = True Then
   
      'Agregamos tabla FactorActAnual
      
      Set Tbl = New TableDef
      Tbl.Name = "FactorActAnual"

      Err.Clear
      Set Fld = Tbl.CreateField("IdFactorActAnual", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "FactorActAnual.IdFactorActAnual", vbExclamation
         lUpdOK = False
      End If
      
     
      Err.Clear
      Set Fld = Tbl.CreateField("Ano", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "FactorActAnual.Ano", vbExclamation
         lUpdOK = False
      End If
     
      Err.Clear
      Set Fld = Tbl.CreateField("MesRow", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "FactorActAnual.MesRow", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set Fld = Tbl.CreateField("MesCol", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "FactorActAnual.MesCol", vbExclamation
         lUpdOK = False
      End If
     
      Err.Clear
      Set Fld = Tbl.CreateField("Factor", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "FactorActAnual.Factor", vbExclamation
         lUpdOK = False
      End If
      
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON FactorActAnual (IdFactorActAnual) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxAno ON FactorActAnual (Ano, MesRow, MesCol)"
         Rc = ExecSQL(DbMain, Q1, False)
      

      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla FactorActAnual", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing
   
      If lUpdOK Then
         lDbVerAdm = 359
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V358 = lUpdOK
   
End Function

Public Function CorrigeBaseAdm_V357() As Boolean   'agregada 16 mar 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 357 -----------------------------------

   If lDbVerAdm = 357 And lUpdOK = True Then
   
      'agregamos campo aIPC a IPC  (IPC acumulado)
      Set Tbl = DbMain.TableDefs("IPC")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("aIPC", dbDouble)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "IPC.aIPC", vbExclamation
         lUpdOK = False
      End If
      Err.Clear
      
      If lUpdOK Then
         lDbVerAdm = 358
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V357 = lUpdOK
   
End Function

Public Function CorrigeBaseAdm_V356() As Boolean   'agregada 24 dic 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim QBase As String, QEnd As String


   On Error Resume Next

   '--------------------- Versión 356 -----------------------------------

   If (lDbVerAdm = 356) And lUpdOK = True Then
      
      'agregamos nuevo impuesto de retención IMPNAC a partir de enero 2020
      Q1 = "INSERT INTO Impuestos (Impuesto, Porcentaje, FechaDesde) VALUES ('IMPNAC', 0.1075, " & CLng(DateSerial(2020, 1, 1)) & ")"
      Call ExecSQL(DbMain, Q1)
      
      
      'agregamos dos cuentas adicionales a plan de cuentas SII
      QBase = "INSERT INTO PlanCuentasSII (CodigoSII, DescripSII, FmtCodigoSII, Clasificacion) VALUES("
      QEnd = "," & CLASCTA_RESULTADO & ")"
     
      Q1 = "'3030700', 'Deudores Incobrables', '3.03.07.00'"
      Call ExecSQL(DbMain, QBase & Q1 & QEnd)
      
      Q1 = "'3050900', 'Utilidad (pérdida) por resultados devengados en Otras Sociedades', '3.05.09.00'"
      Call ExecSQL(DbMain, QBase & Q1 & QEnd)

      'cambiamos código plan SII para cuenta Deudores incobrables
      QBase = "UPDATE PlanBasico "

      Q1 = "SET CodCtaPlanSII ='3030700' WHERE Codigo = '3010644' AND " & GenLike(DbMain, "Deudores Incobrables", "Descripcion")
      Call ExecSQL(DbMain, QBase & Q1)

      QBase = "UPDATE PlanIntermedio "

      Call ExecSQL(DbMain, QBase & Q1)

      QBase = "UPDATE PlanAvanzado "

      Call ExecSQL(DbMain, QBase & Q1)

      
      If lUpdOK Then
         lDbVerAdm = 357
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V356 = lUpdOK
   
End Function
Public Function CorrigeBaseAdm_V355() As Boolean   'agregada 2 sep 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 355 -----------------------------------

   If lDbVerAdm = 355 And lUpdOK = True Then
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Cobquecura'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Coelemu'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Ninhue'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Portezuelo'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Quirihue'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Ranquil'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Trehuaco'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Bulnes'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Chillan Viejo'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Chillan'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'El Carmen'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Pemuco'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Pinto'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Quillon'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'San Ignacio'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Yungay'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Coihueco'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Ñiquen'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'San Carlos'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'San Fabian'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'San Nicolas'"
      Call ExecSQL(DbMain, Q1)
      
      
      If lUpdOK Then
         lDbVerAdm = 356
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V355 = lUpdOK
   
End Function
Public Function CorrigeBaseAdm_V354() As Boolean   'agregada 17 ene 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 354 -----------------------------------

   If lDbVerAdm = 354 And lUpdOK = True Then
      
      'código SII DTE para TipoDoc Importaciones
      Q1 = "UPDATE TipoDocs SET CodDocDTESII = '914' WHERE Diminutivo = 'IMP'"
      Call ExecSQL(DbMain, Q1)
      
      
      If lUpdOK Then
         lDbVerAdm = 355
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V354 = lUpdOK
   
End Function
Public Function CorrigeBaseAdm_V353() As Boolean   'agregada 15 nov 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 353 -----------------------------------

   If lDbVerAdm = 353 And lUpdOK = True Then
   
      'agregamos campo CodAduana a Monedas (es por TrFacturación)
      Set Tbl = DbMain.TableDefs("Monedas")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodAduana", dbText, 3)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "Monedas.CodAduana", vbExclamation
         lUpdOK = False
      End If
      Err.Clear
      
      'agregamos campo EsFijo a Monedas
      Tbl.Fields.Append Tbl.CreateField("EsFijo", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "Monedas.EsFijo", vbExclamation
         lUpdOK = False
      End If
      
      Q1 = "CREATE UNIQUE INDEX IdxCod ON Monedas (CodAduana)"
      Rc = ExecSQL(DbMain, Q1, False)
         
      Q1 = "UPDATE Monedas SET CodAduana = '200', EsFijo = -1 WHERE Descrip = 'Pesos'"
      Call ExecSQL(DbMain, Q1)

      Q1 = "UPDATE Monedas SET CodAduana = '013', Descrip = 'Dólar USA', EsFijo = -1  WHERE Descrip = 'Dólar'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Monedas SET EsFijo = -1  WHERE Simbolo = 'UF' OR Simbolo = 'UTM'"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVerAdm = 354
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V353 = lUpdOK
   
End Function
Public Function CorrigeBaseAdm_V352() As Boolean   'agregada 31 oct 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 352 -----------------------------------

   If lDbVerAdm = 352 And lUpdOK = True Then
   
      'Eliminamos cuenta duplicada del PlanCuentasSII
      Q1 = "DELETE * FROM PlanCuentasSII WHERE CodigoSII = '1022800'"
      Call ExecSQL(DbMain, Q1)
          
      If lUpdOK Then
         lDbVerAdm = 353
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V352 = lUpdOK

End Function
Public Function CorrigeBaseAdm_V351() As Boolean   'agregada 17 oct 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 351 -----------------------------------

   If lDbVerAdm = 351 And lUpdOK = True Then
   
      'Agregamos campo OldCodigo a CodActiv, esto por el cambio de codificación del SII a partir de Nov 2018
      Set Tbl = DbMain.TableDefs("CodActiv")
     
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("OldCodigo", dbText, 10)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "CodActiv.OldCodigo", vbExclamation
         lUpdOK = False
      End If
                 
      'achicamos campo codigo, tenía tamaño 255, lo cual es ridículo
      Call AlterField(DbMain, "CodActiv", "Codigo", dbText, 10)
     
      'agregamos los nuevos códigos de actividad económica del SII validos desde nov 2018
      Call UpdateCodActiv2018
      
      If lUpdOK Then
         lDbVerAdm = 352
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V351 = lUpdOK

End Function
Public Function CorrigeBaseAdm_V350() As Boolean   'agregada 6 jul 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 350 -----------------------------------

   If lDbVerAdm = 350 And lUpdOK = True Then
   
      'desmarcamos como descontinuado Otros Impuestos
      Q1 = "UPDATE TipoValor SET Valor = Left(Valor, Len(Valor)-3) WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_OTROSIMP
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVerAdm = 351
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V350 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V349() As Boolean   'agregada 5 abr 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 349 -----------------------------------

   If lDbVerAdm = 349 And lUpdOK = True Then
   
      'creamos y llenamos tabla PlanCuentasSII
      Call CrearTblPlanCuentasSII
      
      Call FillPlanCuentasSII
      

      'Agregamos campo CodCtaPlanSII a PlanBasico
      Set Tbl = DbMain.TableDefs("PlanBasico")
     
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaPlanSII", dbText, 10)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanBasico.CodCtaPlanSII", vbExclamation
         lUpdOK = False
      End If
                 
      'Agregamos campo IdCtaPlanSII a PlanIntermedio
      Set Tbl = DbMain.TableDefs("PlanIntermedio")
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaPlanSII", dbText, 10)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanIntermedio.CodCtaPlanSII", vbExclamation
         lUpdOK = False
      End If
                 
      'Agregamos campo IdCtaPlanSII a PlanAvanzado
      Set Tbl = DbMain.TableDefs("PlanAvanzado")
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaPlanSII", dbText, 10)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanAvanzado.CodCtaPlanSII", vbExclamation
         lUpdOK = False
      End If
                     
      'Agregamos campo IdCtaPlanSII a IFRS_PlanIFRS
      Set Tbl = DbMain.TableDefs("IFRS_PlanIFRS")
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaPlanSII", dbText, 10)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "IFRS_PlanIFRS.CodCtaPlanSII", vbExclamation
         lUpdOK = False
      End If
                     
      'actualizamos los planes predefinidos con las Cuentas del Plan de Cuentas SII indicadas por Legal Publishing
      Call UpdateCtaPlanSII("PLanBasico")
      Call UpdateCtaPlanSII("PLanIntermedio")
      Call UpdateCtaPlanSII("PLanAvanzado")
      
      Call UpdateCtaPlanSII_IFRS("IFRS_PlanIFRS")
      
      If lUpdOK Then
         lDbVerAdm = 350
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V349 = lUpdOK

End Function

            
Public Function CorrigeBaseAdm_V348() As Boolean   'agregada 28 dic 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 348 -----------------------------------

   If lDbVerAdm = 348 And lUpdOK = True Then
      
      'Agregamos campo TipoPartida a PlanBasico
      Set Tbl = DbMain.TableDefs("PlanBasico")
     
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoPartida", dbByte)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanBasico.TipoPartida", vbExclamation
         lUpdOK = False
      End If
                 
      'Agregamos campo TipoPartida a PlanIntermedio
      Set Tbl = DbMain.TableDefs("PlanIntermedio")
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoPartida", dbByte)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanIntermedio.TipoPartida", vbExclamation
         lUpdOK = False
      End If
                 
      'Agregamos campo TipoPartida a PlanAvanzado
      Set Tbl = DbMain.TableDefs("PlanAvanzado")
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoPartida", dbByte)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanAvanzado.TipoPartida", vbExclamation
         lUpdOK = False
      End If
                     
      'Agregamos campo TipoPartida a IFRS_PlanIFRS
      Set Tbl = DbMain.TableDefs("IFRS_PlanIFRS")
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoPartida", dbByte)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "IFRS_PlanIFRS.TipoPartida", vbExclamation
         lUpdOK = False
      End If
      
      
                     
      'actualizamos los planes predefinidos con las partidas indicadas por Legal Publishing
      Call UpdateTipoPartidaCtas("PLanBasico")
      Call UpdateTipoPartidaCtas("PLanIntermedio")
      Call UpdateTipoPartidaCtas("PLanAvanzado")
      
      Call UpdateTipoPartidaIFRS("IFRS_PlanIFRS")
 
 
      'agregamos tipo doc OII = "Otros Ingr. Saldo Inicial Libro de Caja" y OEI = "Otros Egr. Saldo Inicial Libro de Caja"
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_CAJAING & " AND Diminutivo = '" & LIBCAJA_OTROSINGINI & "'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja, DocImpExp, DocBoletas, CodDocSII, CodDocDTESII)"
         Q1 = Q1 & " VALUES(" & LIB_CAJAING & "," & LIBCAJA_OTROSINGINI & ", 'Otros Ingr. Saldo Inicial', '" & TDOC_OTROSINGRINI & "', 'ACTIVO', -1, 0, 1, 0, 0, 0, 0, '', '')"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
   
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_CAJAEGR & " AND Diminutivo = '" & LIBCAJA_OTROSEGRINI & "'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja, DocImpExp, DocBoletas, CodDocSII, CodDocDTESII)"
         Q1 = Q1 & " VALUES(" & LIB_CAJAEGR & "," & LIBCAJA_OTROSEGRINI & ", 'Otros Egr. Saldo Inicial', '" & TDOC_OTROSEGRINI & "', 'ACTIVO', -1, 0, 1, 0, 0, 0, 0, '', '')"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)

      If lUpdOK Then
         lDbVerAdm = 349
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V348 = lUpdOK

End Function


Public Function CorrigeBaseAdm_V347() As Boolean   'agregada 19 dic 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field

   
   On Error Resume Next
   
   If lDbVerAdm = 347 And lUpdOK = True Then
   
      'agregamos campo Activo a Usuarios
      Set Tbl = DbMain.TableDefs("EmpresasAno")

      Err.Clear
      Set Fld = Tbl.CreateField("CredArt33bis", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.CredArt33bis", vbExclamation
         lUpdOK = False

      End If
          
      If lUpdOK Then
         lDbVerAdm = 348
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V347 = lUpdOK

End Function


Public Function CorrigeBaseAdm_V346() As Boolean   'agregada 28 nov 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field

   
   On Error Resume Next
   
   If lDbVerAdm = 346 And lUpdOK = True Then
   
      'actualizamos perfil (todo) para que quede con FFFF
      
      Q1 = "UPDATE Perfiles SET Privilegios = 65535 WHERE Nombre = '(todo)'"
      Call ExecSQL(DbMain, Q1)
     
      If lUpdOK Then
         lDbVerAdm = 347
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V346 = lUpdOK

End Function


Public Function CorrigeBaseAdm_V345() As Boolean   'agregada 13 nov 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field

   
   On Error Resume Next
   
   If lDbVerAdm = 345 And lUpdOK = True Then
   
      'agregamos campo Activo a Usuarios
      Set Tbl = DbMain.TableDefs("Usuarios")

      Err.Clear
      Set Fld = Tbl.CreateField("Activo", dbBoolean)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "Usuarios.Activo", vbExclamation
         lUpdOK = False

      End If
     
      'agregamos campo HabilitadoHasta a Usuarios (fecha de deshabilitación)
      Err.Clear
      Set Fld = Tbl.CreateField("HabilitadoHasta", dbLong)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "Usuarios.HabilitadoHasta", vbExclamation
         lUpdOK = False

      End If
      
      Q1 = "UPDATE Usuarios SET Activo = 1"
      Call ExecSQL(DbMain, Q1)
     
      If lUpdOK Then
         lDbVerAdm = 346
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V345 = lUpdOK

End Function


Public Function CorrigeBaseAdm_V344() As Boolean   'agregada 14 sept 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field

   
   On Error Resume Next
   
   If lDbVerAdm = 344 And lUpdOK = True Then
   
      'agregamos campo SaldoLibroCaja a EmpresasAno
      Set Tbl = DbMain.TableDefs("EmpresasAno")

      Err.Clear
      Set Fld = Tbl.CreateField("SaldoLibroCaja", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.SaldoLibroCaja", vbExclamation
         lUpdOK = False

      End If
     
      If lUpdOK Then
         lDbVerAdm = 345
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V344 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V343() As Boolean   'agregada 23 ago 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field

   
   On Error Resume Next
   
   If lDbVerAdm = 343 And lUpdOK = True Then
   
      Call ExecSQL(DbMain, "UPDATE Regiones SET COMUNA = 'MARCHIGUE' WHERE COMUNA = 'MARCHIHUE'")
      Call ExecSQL(DbMain, "UPDATE Regiones SET COMUNA = 'SAAVEDRA' WHERE COMUNA = 'PUERTO SAAVEDRA'")
      Call ExecSQL(DbMain, "UPDATE Regiones SET COMUNA = 'NATALES' WHERE COMUNA = 'PUERTO NATALES'")
      Call ExecSQL(DbMain, "UPDATE Regiones SET COMUNA = 'SANTIAGO' WHERE COMUNA = 'SANTIAGO (*)'")
      Call ExecSQL(DbMain, "UPDATE Regiones SET COMUNA = 'SANTIAGO CENTRO (*)' WHERE COMUNA = 'SANTIAGO CENTRO'")
                  
      If lUpdOK Then
         lDbVerAdm = 344
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V343 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V342() As Boolean   'agregada 20 jun 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field

   
   On Error Resume Next
   
   If lDbVerAdm = 342 And lUpdOK = True Then
   
      'agregamos campo IngresarTotal a TipoDocs
      Set Tbl = DbMain.TableDefs("TipoDocs")

      Err.Clear
      Set Fld = Tbl.CreateField("IngresarTotal", dbByte)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.IngresarTotal", vbExclamation
         lUpdOK = False

      End If
     
      'agregamos campo IngresarTotal a TipoDocs
      Set Tbl = DbMain.TableDefs("TipoDocs")

      Err.Clear
      Set Fld = Tbl.CreateField("TieneNumDocHasta", dbByte)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.TieneNumDocHasta", vbExclamation
         lUpdOK = False

      End If
    
      'agregamos campo IngresarTotal a TipoDocs
      Set Tbl = DbMain.TableDefs("TipoDocs")

      Err.Clear
      Set Fld = Tbl.CreateField("TieneCantBoletas", dbByte)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.TieneCantBoletas", vbExclamation
         lUpdOK = False

      End If
    
      Call AlterField(DbMain, "TipoDocs", "TieneAfecto", dbByte)
      Call AlterField(DbMain, "TipoDocs", "TieneExento", dbByte)
      
      Tbl.Fields.Refresh
      
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneAfecto = " & VAL_OPCIONAL & " WHERE TieneAfecto <> 0")
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneExento = " & VAL_OPCIONAL & " WHERE TieneExento <> 0")
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET IngresarTotal = 1 WHERE TieneExento = 0 AND TieneAfecto = 0")
      
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneAfecto = " & VAL_OBLIGATORIO & " WHERE Diminutivo IN('BOV', 'BEX', 'DVB', 'FAV', 'MRG', 'VPE') ")
      
      
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneAfecto = " & VAL_NOPERMITIDO & " WHERE Diminutivo IN('BOE', 'EXP', 'FVE') ")
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneAfecto = " & VAL_OPCIONAL & " WHERE Diminutivo IN('NCV', 'NDV') ")
      

      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneExento = " & VAL_OBLIGATORIO & " WHERE Diminutivo IN('BEX', 'BOE', 'EXP', 'FVE', 'VSD') ")
      
      
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneExento = " & VAL_NOPERMITIDO & " WHERE Diminutivo IN('BOV', 'DVB', 'VPE') ")
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneExento = " & VAL_OPCIONAL & " WHERE Diminutivo IN('FAV', 'MRG', 'NCV', 'NDV') ")
      
     
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneNumDocHasta = " & VAL_OBLIGATORIO & " WHERE Diminutivo IN('BOV', 'BEX', 'BOE', 'MRG') ")
         
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneNumDocHasta = " & VAL_NOPERMITIDO & " WHERE Diminutivo IN('EXP', 'FAV', 'FVE', 'NCV', 'NDV', 'VPE', 'DVB') ")
                  
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneCantBoletas = " & VAL_OBLIGATORIO & " WHERE Diminutivo IN('VPE') ")
                  
      If lUpdOK Then
         lDbVerAdm = 343
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V342 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V341() As Boolean   'agregada 16 jun 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field

   
   On Error Resume Next
   
   If lDbVerAdm = 341 And lUpdOK = True Then
   
      'cambiamos MRG para que no edite Afecto pero si Exento
      Q1 = "UPDATE TipoDocs SET TieneAfecto = 0, TieneExento = 1, DocBoletas = -1  WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = '" & TDOC_MAQREGISTRADORA & "'"
      Call ExecSQL(DbMain, Q1)
         
      Q1 = "UPDATE TipoDocs SET DocBoletas = -1  WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = '" & TDOC_BOLVENTAEX & "'"
      Call ExecSQL(DbMain, Q1)
         
      If lUpdOK Then
         lDbVerAdm = 342
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V341 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V340() As Boolean   'agregada 1 jun 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field

   
   On Error Resume Next
   
   If lDbVerAdm = 340 And lUpdOK = True Then
   
      'cambiamos MRG para que no edite Afecto y Exento
      Q1 = "UPDATE TipoDocs SET TieneAfecto = 0, TieneExento = 0  WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'MRG'"
      Call ExecSQL(DbMain, Q1)
         
      If lUpdOK Then
         lDbVerAdm = 341
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V340 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V339() As Boolean   'agregada 17 may 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field

   
   On Error Resume Next
   
   If lDbVerAdm = 339 And lUpdOK = True Then
   
      'cambiamos ExigeRut para documento OTC
      Q1 = "UPDATE TipoDocs SET ExigeRUT = -1 WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'OTC'"
      Call ExecSQL(DbMain, Q1)
         
      If lUpdOK Then
         lDbVerAdm = 340
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V339 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V338() As Boolean   'agregada 10 abr 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field

   
   On Error Resume Next
   
   If lDbVerAdm = 338 And lUpdOK = True Then
   
      'agregamos tipo doc OIN = "Otros Ingresos Libro de Caja" y OEG = "Otros Egresos Libro de Caja"
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_CAJAING & " AND Diminutivo = 'OIN'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja, DocImpExp, DocBoletas, CodDocSII, CodDocDTESII)"
         Q1 = Q1 & " VALUES(" & LIB_CAJAING & "," & LIBCAJA_OTROSING & ", 'Otros Ingresos', '" & TDOC_OTROSINGRESOS & "', 'ACTIVO', -1, 0, 1, 0, 0, 0, 0, '', '')"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
   
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_CAJAEGR & " AND Diminutivo = 'OEG'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja, DocImpExp, DocBoletas, CodDocSII, CodDocDTESII)"
         Q1 = Q1 & " VALUES(" & LIB_CAJAEGR & "," & LIBCAJA_OTROSEGR & ", 'Otros Egresos', '" & TDOC_OTROSEGRESOS & "', 'ACTIVO', -1, 0, 1, 0, 0, 0, 0, '', '')"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
  
      If lUpdOK Then
         lDbVerAdm = 339
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V338 = lUpdOK

End Function


Public Function CorrigeBaseAdm_V337() As Boolean   'agregada y entregada 5 abr 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field


   On Error Resume Next
   
   If lDbVerAdm = 337 And lUpdOK = True Then

      'agregamos IVA Irrecuperable a Factura de Compra
      Q1 = "UPDATE TipoValor SET TipoDoc = ',1,3,4,5,6,9,13,' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo IN( " & LIBCOMPRAS_IVAIRREC1 & "," & LIBCOMPRAS_IVAIRREC2 & "," & LIBCOMPRAS_IVAIRREC3 & "," & LIBCOMPRAS_IVAIRREC4 & "," & LIBCOMPRAS_IVAIRREC9 & ")"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVerAdm = 338
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)

      End If
   End If
  
  CorrigeBaseAdm_V337 = lUpdOK

End Function


Public Function CorrigeBaseAdm_V336() As Boolean   'agregada 28 mar 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field


   On Error Resume Next
   
   If lDbVerAdm = 336 And lUpdOK = True Then
   
      Call AlterField(DbMain, "TipoValor", "CodSIIDTE", dbText, 3)
   
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
      Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPBEDANALCAZUCAR & ", 'Impto. Beb. analc. alto azúcar', ' ', ' ', 0, ' ', ',1,3,4,', 'Impto. Bebidas','Analc. alto azucar',' ', 13, 18, -1, '271', 'Impto. Bebidas analcohólicas con alto contenido de azúcar')"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVerAdm = 337
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)

      End If
   End If
  
  CorrigeBaseAdm_V336 = lUpdOK

End Function


Public Function CorrigeBaseAdm_V335() As Boolean   'agregada 10 mar 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field


   On Error Resume Next
   
   If lDbVerAdm = 335 And lUpdOK = True Then

      'modificamos Multiple en IVA Irrecuperable
      Q1 = "UPDATE TipoValor SET Multiple = '-1' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo IN( " & LIBCOMPRAS_IVAIRREC1 & "," & LIBCOMPRAS_IVAIRREC2 & "," & LIBCOMPRAS_IVAIRREC3 & "," & LIBCOMPRAS_IVAIRREC4 & "," & LIBCOMPRAS_IVAIRREC9 & ")"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVerAdm = 336
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)

      End If
   End If
  
  CorrigeBaseAdm_V335 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V334() As Boolean   'agregada 2 feb 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field


   On Error Resume Next
   
   If lDbVerAdm = 334 And lUpdOK = True Then

      'agregamos documentos al IVA Irrecuperable
      Q1 = "UPDATE TipoValor SET TipoDoc = ',1,3,4,6,9,13,' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo IN( " & LIBCOMPRAS_IVAIRREC1 & "," & LIBCOMPRAS_IVAIRREC2 & "," & LIBCOMPRAS_IVAIRREC3 & "," & LIBCOMPRAS_IVAIRREC4 & "," & LIBCOMPRAS_IVAIRREC9 & ")"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVerAdm = 335
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)

      End If
   End If
  
  CorrigeBaseAdm_V334 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V333() As Boolean   'agregada 31 ene 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field


   On Error Resume Next
   
   If lDbVerAdm = 333 And lUpdOK = True Then

      'marcamos como descontinuado Otros Impuestos
      Q1 = "UPDATE TipoValor SET Orden = Orden + 1000 WHERE Right(Valor, 3) = '(*)' AND TipoLib = " & LIB_COMPRAS
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVerAdm = 334
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)

      End If
   End If
  
  CorrigeBaseAdm_V333 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V332() As Boolean   'agregada 30 ene 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field


   
   On Error Resume Next
   
   If lDbVerAdm = 332 And lUpdOK = True Then
               
      'marcamos como descontinuado Otros Impuestos
      Q1 = "UPDATE TipoValor SET Valor = Valor & '(*)', Atributo = 'OTROSIMP' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_OTROSIMP
      Call ExecSQL(DbMain, Q1)
                  
      'códigos SII DTE para TipoDoc
      Q1 = "UPDATE TipoDocs SET CodDocDTESII = '61' WHERE Diminutivo = 'NCF'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET CodDocDTESII = '56' WHERE Diminutivo = 'NDF'"
      Call ExecSQL(DbMain, Q1)
                 
      'Insertamos IVA Anticipado Harina en Libro de Ventas
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE)"
      Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IVAANTICIPADOHARINA & ", 'IVA Anticipado Harina', ' ', ' ', -1, ' ', ',1,3,4,6,', 'IVA Anticipado','Harina',' ', 20, 12, 0, '19')"
      Call ExecSQL(DbMain, Q1)
                                    
      'Actualizamos TipoDoc de IVA Anticipado Carne
      Q1 = "UPDATE TipoValor SET TipoDoc = ',1,3,4,6,'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IVAANTICIPADOCARNE
      Call ExecSQL(DbMain, Q1)
            
      If lUpdOK Then
         lDbVerAdm = 333
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V332 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V331() As Boolean   'agregada 25 ene 2017
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field


   
   On Error Resume Next
   
   If lDbVerAdm = 331 And lUpdOK = True Then
         
      Call ActImpAdicionales2016
      
      'códigos SII para TipoDoc
      Q1 = "UPDATE TipoDocs SET CodDocSII = '60' WHERE Diminutivo = 'NCF'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET CodDocSII = '61' WHERE Diminutivo = 'NDF'"
      Call ExecSQL(DbMain, Q1)
                  
      Q1 = "UPDATE TipoDocs SET CodDocSII = '29' WHERE Diminutivo = 'FIC'"
      Call ExecSQL(DbMain, Q1)
                  
      If lUpdOK Then
         lDbVerAdm = 332
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V331 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V330() As Boolean   'agregada 21 nov 2016
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field


   
   On Error Resume Next
   
   If lDbVerAdm = 330 And lUpdOK = True Then
   
      'agregamos campo TitCompleto a TipoValor
      Set Tbl = DbMain.TableDefs("TipoValor")

      Err.Clear
      Set Fld = Tbl.CreateField("TitCompleto", dbText, 100)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.TitCompleto", vbExclamation
         lUpdOK = False

      End If
                        
      If lUpdOK Then
         lDbVerAdm = 331
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V330 = lUpdOK

End Function



Private Function CorrigeBaseAdm_V329() As Boolean     'agregada 20 jul 2016 / entregada
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 328 -----------------------------------
   If lDbVerAdm = 329 And lUpdOK = True Then
   
      Err.Clear
      
      Set Tbl = DbMain.TableDefs("TipoValor")
      
      'agregamos campos a tabla TipoValor: Tasa, Es Recuperable
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("Tasa", dbSingle)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.Tasa", vbExclamation
         lUpdOK = False
      End If
                 
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("EsRecuperable", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.EsRecuperable", vbExclamation
         lUpdOK = False
      End If
                 
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodSIIDTE", dbText, 2)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.CodSIIDTE", vbExclamation
         lUpdOK = False
      End If
                  
      'se corrige error de ortografía
      Q1 = "UPDATE TipoValor SET  Valor = 'Impto. Pisco, Licores, Whisky, Aguard.' "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPPISCO
      Call ExecSQL(DbMain, Q1)
  
      'se corrige CodDocDTESII de Factura de Exportación
      Q1 = "UPDATE TipoDocs SET  CodDocDTESII = '110'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'EXP'"    'Factura de exportación
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos tasa a impuestos adicionales
      Q1 = "UPDATE TipoValor SET  Tasa = 31.50, EsRecuperable = -1, CodSIIDTE = '24' "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPPISCO
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET  Tasa = 20.50, EsRecuperable = -1, CodSIIDTE = '25' "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPVINOS
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET  Tasa = 20.50, EsRecuperable = -1, CodSIIDTE = '26' "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPCERVEZA
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET  Tasa = 10, EsRecuperable = -1, CodSIIDTE = '27' "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPBEBANHALC
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET  Valor = 'Imp. Adic. Alfombras, Tapices, Casas Rodantes, Caviar', "
      Q1 = Q1 & " Tit1 = 'Imp. Alfombras', Tit2 = 'Tapices', Tasa = 15, EsRecuperable = -1, CodSIIDTE = '44' "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPART37E
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET  Valor = 'Imp. Adic. Pirotecnia', "
      Q1 = Q1 & " Tit1 = 'Imp. Adic.', Tit2 = 'Pirotecnia', Tasa = 50, EsRecuperable = -1, CodSIIDTE = '45' "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPART37J
      Call ExecSQL(DbMain, Q1)
      
'      Q1 = "UPDATE TipoValor SET  Valor = 'IVA Anticipado Harina', "
'      Q1 = Q1 & " Tit1 = 'IVA Anticipado', Tit2 = 'Harina', Tasa = 12, EsRecuperable = -1, CodSIIDTE = '19' "
'      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_RETANTCAMBIOSUJHARINA
'      Call ExecSQL(DbMain, Q1)
'
      Q1 = "UPDATE TipoValor SET  Valor = 'IVA Anticipado Carne', "
      Q1 = Q1 & " Tit1 = 'IVA Anticipado', Tit2 = 'Carne', Tasa = 5, EsRecuperable = -1, CodSIIDTE = '18' "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IVAANTICIPADOCARNE
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET  Valor = 'IVA Retenido Construcción', "
      Q1 = Q1 & " Tit1 = 'IVA Retenido', Tit2 = 'Constr.', Tasa = 100, EsRecuperable = -1, CodSIIDTE = '41' "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IVAADQCONSTINMUEBLES
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE)"
      Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_JOYAS & ", 'Imp. Adic. Joyas, Piedras Preciosas, Pieles Finas', ' ', ' ', -1, ' ', ',1,', 'Imp. Adic.','Joyas, Pieles',' ', 22, 15, -1, '23')"
      Call ExecSQL(DbMain, Q1)
     
      
         
   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 330
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V329 = lUpdOK

End Function
Private Function CorrigeBaseAdm_V328() As Boolean     'entregada 2 mar 2015
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 328 -----------------------------------
   If lDbVerAdm = 328 And lUpdOK = True Then
   
      'insertamos nuevos impuestos
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden)"
      Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAADQCONSTINMUEBLES & ", 'IVA por Adq. o Const. Inmuebles', ' ', ' ', -1, 766, ',1,', 'IVA','Inmuebles','25', 26)"
      Call ExecSQL(DbMain, Q1)
   
  
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden)"
      Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IVAADQCONSTINMUEBLES & ", 'IVA por Adq. o Const. Inmuebles', ' ', ' ', -1, 764, ',1,', 'IVA','Inmuebles',' ', 21 )"
      Call ExecSQL(DbMain, Q1)
  
         
   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 329
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V328 = lUpdOK

End Function
Private Function CorrigeBaseAdm_V327() As Boolean     'entregada 26 nov 2015
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 327 -----------------------------------
   If lDbVerAdm = 327 And lUpdOK = True Then
   
      Q1 = "UPDATE TipoValor SET TipoDoc = ',3,4,5,10,11,' WHERE TipoLib = " & LIB_COMPRAS & " AND Valor IN ('IVA Retenido Parcial', 'IVA Retenido Total')"
      Call ExecSQL(DbMain, Q1)
       
   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 328
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V327 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V326() As Boolean   'entregada 19 nov 2015
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Base As Integer, Base2 As Integer
   
   Base = 10000
   Base2 = 20000

   
   On Error Resume Next
   
   If lDbVerAdm = 326 And lUpdOK = True Then
   
      'agregamos campos CodF29CountSuper, CodF29IVASuper para traspaso a F29 de compras con FAC en supermercados
      Set Tbl = DbMain.TableDefs("TipoDocs")

      Err.Clear
      Set Fld = Tbl.CreateField("CodF29CountSuper", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29CountSuper", vbExclamation
         lUpdOK = False

      End If
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29IVASuper", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29IVASuper", vbExclamation
         lUpdOK = False

      End If
      
      'Actualizamos códigos F29 FAC para Supermercados
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & " CodF29CountSuper = 761, CodF29IVASuper = 762"
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'FAC'"
      Call ExecSQL(DbMain, Q1)
      
      'Arreglamos codificación para exportación a Form29
      
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29CountDTE = " & Base + 511
      Q1 = Q1 & ", CodF29IVAIrrecDTE = " & Base + 514
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo IN('FAC', 'NCC', 'NDC', 'IMP')"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Count = " & Base + 519
      Q1 = Q1 & ", CodF29Neto = " & Base + 520
      Q1 = Q1 & ", CodF29CountDTE = " & Base + 511
      Q1 = Q1 & ", CodF29IVAIrrecDTE = " & Base + 514
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'FCC'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Neto = " & Base + 502
      Q1 = Q1 & ", CodF29NetoNoGiro = " & Base + 717
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'FAV'"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Neto = " & Base + 510
      Q1 = Q1 & ", CodF29NetoNoGiro = " & Base + 734
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'NCV'"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Neto = " & Base + 513
      Q1 = Q1 & ", CodF29NetoNoGiro = " & Base + 717
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'NDV'"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29DifIVARetParcial = " & Base + 517
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'FCV'"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Neto = " & Base + 501
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'LFV'"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Neto = " & Base + 111
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'BOV'"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Count = " & Base + 519
      Q1 = Q1 & ", CodF29Neto = " & -(Base + 520)
      Q1 = Q1 & ", CodF29CountDTE = " & Base + 511
      Q1 = Q1 & ", CodF29IVAIrrecDTE = " & Base + 514
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'NCF'"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Count = " & Base + 519
      Q1 = Q1 & ", CodF29Neto = " & Base + 520
      Q1 = Q1 & ", CodF29CountDTE = " & Base + 511
      Q1 = Q1 & ", CodF29IVAIrrecDTE = " & Base + 514
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'NDF'"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Neto = " & Base + 111
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'BEX'"
      Call ExecSQL(DbMain, Q1)
      
    
     
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Count = " & Base + 519
      Q1 = Q1 & ", CodF29Neto = " & Base + 520
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'FIC'"
      Call ExecSQL(DbMain, Q1)

      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29IVA = " & Base + 587
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'FIV'"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Neto = " & Base + 759
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'VPE'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET"
      Q1 = Q1 & "  CodF29 = " & Base + 556
      Q1 = Q1 & " WHERE Valor = 'IVA anticipado del periodo Harina'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET"
      Q1 = Q1 & "  CodF29 = " & Base2 + 556
      Q1 = Q1 & " WHERE Valor = 'IVA anticipado del periodo Carne'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET"
      Q1 = Q1 & "  CodF29 = " & Base + 587
      Q1 = Q1 & " WHERE Valor = 'IVA Retenido Total'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET"
      Q1 = Q1 & "  CodF29 = " & Base + 555
      Q1 = Q1 & " WHERE Valor = 'Ret. anticipo cambio sujeto Harina'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET"
      Q1 = Q1 & "  CodF29 = " & Base2 + 555
      Q1 = Q1 & " WHERE Valor = 'Ret. anticipo cambio sujeto Carne'"
      Call ExecSQL(DbMain, Q1)
      
      
      If lUpdOK Then
         lDbVerAdm = 327
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V326 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V325() As Boolean   'agregada 17 jun 2015
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field


   
   On Error Resume Next
   
   If lDbVerAdm = 325 And lUpdOK = True Then
   
      'agregamos campo RemIVAUTM a EmpresasAno para almacenar Remanete de IVA de este año al cierre del mismo
      Set Tbl = DbMain.TableDefs("EmpresasAno")

      Err.Clear
      Set Fld = Tbl.CreateField("RemIVAUTM", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.RemIVAUTM", vbExclamation
         lUpdOK = False

      End If
         
      'agregamos camp RemIVAUTMAnoAnt a EmpresasAno para almacenar Remanete de IVA de año anterior, que se obtiene en la Apertura desde el año anterior o por ingreso directo
      Err.Clear
      Set Fld = Tbl.CreateField("RemIVAUTMAnoAnt", dbDouble)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.RemIVAUTMAnoAnt", vbExclamation
         lUpdOK = False

      End If
         
      If lUpdOK Then
         lDbVerAdm = 326
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V325 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V324() As Boolean   'entregada 28 mayo 2015
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field


   
   On Error Resume Next
   
   If lDbVerAdm = 324 And lUpdOK = True Then
   
      'Actualizamos códigos F29 VPE
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & " CodF29Count = 758, CodF29Neto = 5759, CodF29IVA = 759"
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'VPE'"
      Call ExecSQL(DbMain, Q1)
         
      If lUpdOK Then
         lDbVerAdm = 325
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V324 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V323() As Boolean   'entregada abril 2015
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long
   Dim Tbl As TableDef
   Dim Fld As Field


   
   On Error Resume Next
   
   If lDbVerAdm = 323 And lUpdOK = True Then
   
      'agregamos tipo doc VCPME (VPE) = Vales comprobantes de Pago Medios Electrónicos en Libro Ventas
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = '" & TDOC_VALEPAGOELECTR & "'"
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
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, CodF29Count, CodF29Neto, CodF29IVA, TieneAfecto, TieneExento, ExigeRUT, EsRebaja, DocImpExp, DocBoletas, CodDocSII)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Vale Pago Electrónico', '" & TDOC_VALEPAGOELECTR & "', 'ACTIVO', -1, 110, 6111, 111, 0, 0, 0, 0, 0, -1, '48')"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
   
      If lUpdOK Then
         lDbVerAdm = 324
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V323 = lUpdOK

End Function
Private Function CorrigeBaseAdm_V322() As Boolean     '10 mar 2015
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 322 -----------------------------------
   If lDbVerAdm = 322 And lUpdOK = True Then
   
      Set Tbl = DbMain.TableDefs("Impuestos")

      Err.Clear
      Set Fld = Tbl.CreateField("FechaDesde", dbLong)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "Impuestos.FechaDesde", vbExclamation
         lUpdOK = False

      End If
      
      'agregamos nuevo impuesto de retención IMPEXT a partir de enero 2015
      Q1 = "INSERT INTO Impuestos (Impuesto, Porcentaje, FechaDesde) VALUES ('IMPEXT', 0.35, " & CLng(DateSerial(2015, 1, 1)) & ")"
      Call ExecSQL(DbMain, Q1)
       
      
   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 323
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V322 = lUpdOK

End Function
Private Function CorrigeBaseAdm_V321() As Boolean     '6 mar 2015
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 321 -----------------------------------
   If lDbVerAdm = 321 And lUpdOK = True Then
         
      'cambiamos iíndice para que no sea único
      Q1 = "DROP INDEX Imp ON Impuestos "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX Imp ON Impuestos (Impuesto) "
      Rc = ExecSQL(DbMain, Q1, False)
             
   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 322
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V321 = lUpdOK

End Function
Private Function CorrigeBaseAdm_V320() As Boolean     'entregada 11 nov 2014
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 320 -----------------------------------
   If lDbVerAdm = 320 And lUpdOK = True Then
   
      Q1 = "UPDATE TipoValor SET TipoDoc = ',3,4,5,' WHERE Valor IN ('IVA Retenido Parcial', 'IVA Retenido Total')"
      Call ExecSQL(DbMain, Q1)
       
   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 321
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V320 = lUpdOK

End Function
Private Function CorrigeBaseAdm_V319() As Boolean     'entregada 6 nov 2014
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 319 -----------------------------------
   If lDbVerAdm = 319 And lUpdOK = True Then
   
      'Agregamos el campo Orden a la tabla TipoValor

      Set Tbl = DbMain.TableDefs("TipoValor")

      Err.Clear
      Set Fld = Tbl.CreateField("Orden", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.Orden", vbExclamation
         lUpdOK = False

      End If

      Q1 = "UPDATE TipoValor SET Orden = 1 WHERE Codigo = " & LIBCOMPRAS_AFECTO
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 2 WHERE Codigo = " & LIBCOMPRAS_EXENTO
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 3 WHERE Codigo = " & LIBCOMPRAS_TOTAL
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 4 WHERE Codigo = " & LIBCOMPRAS_IVACREDFISC
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 5 WHERE Codigo = " & LIBCOMPRAS_OTROSIMP
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 6 WHERE Codigo = " & LIBCOMPRAS_IVAIRREC
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 7 WHERE Codigo = " & LIBCOMPRAS_IVAACTFIJO
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 8 WHERE Codigo = " & LIBCOMPRAS_IVARETPARC
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 9 WHERE Codigo = " & LIBCOMPRAS_IVARETTOT
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 10 WHERE Codigo = " & LIBCOMPRAS_IMPPISCO
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 11 WHERE Codigo = " & LIBCOMPRAS_IMPVINOS
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 12 WHERE Codigo = " & LIBCOMPRAS_IMPCERVEZA
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 13 WHERE Codigo = " & LIBCOMPRAS_IMPBEBANALC
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 14 WHERE Codigo = " & LIBCOMPRAS_ILABEDANALCAZUCAR
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 15 WHERE Codigo = " & LIBCOMPRAS_ILANOTASDEB
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 16 WHERE Codigo = " & LIBCOMPRAS_ILANOTASCRED
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 17 WHERE Codigo = " & LIBCOMPRAS_IVAANTICIPHARINA
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 18 WHERE Codigo = " & LIBCOMPRAS_IVAANTICIPCARNE
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 19 WHERE Codigo = " & LIBCOMPRAS_IMPESPDIESEL
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 20 WHERE Codigo = " & LIBCOMPRAS_IMPESPPETRGRAL
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 21 WHERE Codigo = " & LIBCOMPRAS_IMPESPDIESELTRANS       'LIBCOMPRAS_IMPESPPETRTRANS
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 22 WHERE Codigo = " & LIBCOMPRAS_IMPESPPETRGENCF
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 23 WHERE Codigo = " & LIBCOMPRAS_IMPESPPETRCARGACF
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 24 WHERE Codigo = " & LIBCOMPRAS_IMPESPPETRGENSINCF
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 25 WHERE Codigo = " & LIBCOMPRAS_IMPESPPETRCARGASINCF
      Call ExecSQL(DbMain, Q1)
       
       
      Q1 = "UPDATE TipoValor SET Orden = 1 WHERE Codigo = " & LIBVENTAS_AFECTO
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 2 WHERE Codigo = " & LIBVENTAS_EXENTO
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 3 WHERE Codigo = " & LIBVENTAS_TOTAL
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 4 WHERE Codigo = " & LIBVENTAS_IVADEBFISC
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 5 WHERE Codigo = " & LIBVENTAS_OTROSIMP
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 6 WHERE Codigo = " & LIBVENTAS_REBAJA65
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 7 WHERE Codigo = " & LIBVENTAS_IVARETPARC
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 8 WHERE Codigo = " & LIBVENTAS_IVARETTOT
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 9 WHERE Codigo = " & LIBVENTAS_RETMARGENCOM
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 10 WHERE Codigo = " & LIBVENTAS_IMPPISCO
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 11 WHERE Codigo = " & LIBVENTAS_IMPVINOS
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 12 WHERE Codigo = " & LIBVENTAS_IMPCERVEZA
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 13 WHERE Codigo = " & LIBVENTAS_IMPBEBANHALC
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 14 WHERE Codigo = " & LIBVENTAS_ILABEDANALCAZUCAR
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 15 WHERE Codigo = " & LIBVENTAS_ILANOTASDEB
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 16 WHERE Codigo = " & LIBVENTAS_ILANOTASCRED
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 17 WHERE Codigo = " & LIBVENTAS_IMPART37E
      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 18 WHERE Codigo = " & LIBVENTAS_IMPART37J
      Call ExecSQL(DbMain, Q1)
'      Q1 = "UPDATE TipoValor SET Orden = 19 WHERE Codigo = " & LIBVENTAS_RETANTCAMBIOSUJHARINA
'      Call ExecSQL(DbMain, Q1)
      Q1 = "UPDATE TipoValor SET Orden = 20 WHERE Codigo = " & LIBVENTAS_RETANTCAMBIOSUJCARNE
      Call ExecSQL(DbMain, Q1)
           
       
   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 320
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V319 = lUpdOK

End Function
Private Function CorrigeBaseAdm_V318() As Boolean     'entregada 26 oct 2014
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 318 -----------------------------------
   If lDbVerAdm = 318 And lUpdOK = True Then
   
      'insertamos nuevos impuestos
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII)"
      Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPESPPETRGENCF & ", 'Imp. Esp. Petr. Diesel Gen. CF', ' ', ' ', -1, 742, ',1,3,4,', 'IEPD','Genral. CF','25' )"
      Call ExecSQL(DbMain, Q1)
   
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII)"
      Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPESPPETRCARGACF & ", 'Imp. Esp. Petr. Diesel Transp. CF', ' ', ' ', -1, 743, ',1,3,4,', 'IEPD','Trans. CF','29' )"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII)"
      Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPESPPETRGENSINCF & ", 'Imp. Esp. Petr. Diesel Gen. s/CF', ' ', ' ', -1, 0, ',1,3,4,', 'IEPD','Genral s/CF',' ' )"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII)"
      Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPESPPETRCARGASINCF & ", 'Imp. Esp. Petr. Diesel Transp. s/CF', ' ', ' ', -1, 0, ',1,3,4,', 'IEPD','Trans. s/CF',' ' )"
      Call ExecSQL(DbMain, Q1)
  
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, CodF29_Adic, TipoDoc, Tit1, Tit2, CodImpSII)"
      Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_ILABEDANALCAZUCAR & ", 'ILA Beb. Analc. c/elevado cont. Azúcar', ' ', ' ', -1, 753, 754, ',1,3,4,', 'ILA Analc.','Alto Azucar',' ' )"
      Call ExecSQL(DbMain, Q1)
  
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII)"
      Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_ILABEDANALCAZUCAR & ", 'ILA Beb. Analc. c/elevado cont. Azúcar', ' ', ' ', -1, 752, ',1,3,4,', 'ILA Analc.','Alto Azucar',' ' )"
      Call ExecSQL(DbMain, Q1)
  
      'actualizamos otros imp. específicos del petróleo
      Q1 = "UPDATE TipoValor SET TipoDoc = ',1,3,4,' "
      Q1 = Q1 & " WHERE Codigo = " & LIBCOMPRAS_IMPESPPETRGRAL
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET TipoDoc = ',1,3,4,' "
      Q1 = Q1 & " WHERE Codigo = " & LIBCOMPRAS_IMPESPDIESELTRANS             'LIBCOMPRAS_IMPESPPETRTRANS
      Call ExecSQL(DbMain, Q1)
         
   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 319
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V318 = lUpdOK

End Function
Private Function CorrigeBaseAdm_V317() As Boolean     'agregada 19 nov 2013
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 317 -----------------------------------
   If lDbVerAdm = 317 And lUpdOK = True Then
   
      'Agregamos el campo CodIFRS en la tabla de Plan Avanzado (ahora ya no está separado el tema de IFRS en dos tablas)

      Set Tbl = DbMain.TableDefs("PlanAvanzado")

      Err.Clear
      Set Fld = Tbl.CreateField("CodIFRS", dbText, 15)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanAvanzado.CodIFRS_EstRes", vbExclamation
         lUpdOK = False

      End If
      
      'Agregamos el campo CodIFRS en la tabla de Plan Intermedio (ahora ya no está separado el tema de IFRS en dos tablas)

      Set Tbl = DbMain.TableDefs("PlanIntermedio")

      Err.Clear
      Set Fld = Tbl.CreateField("CodIFRS", dbText, 15)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanIntermedio.CodIFRS_EstRes", vbExclamation
         lUpdOK = False

      End If
      
      'Agregamos el campo CodIFRS en la tabla de Plan Básico (ahora ya no están separados)
      
      Set Tbl = DbMain.TableDefs("PlanBasico")

      Err.Clear
      Set Fld = Tbl.CreateField("CodIFRS", dbText, 15)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanBasico.CodIFRS_EstRes", vbExclamation
         lUpdOK = False

      End If

      'Insertamos la tabla con los reportes IFRS y asignamos los códigos IFRS a los planes predefinidos

      Call InsertTblIFRS_50
      

      

   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 318
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V317 = lUpdOK

End Function
Private Function CorrigeBaseAdm_V316() As Boolean     'entregada 13 nov 2013
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 316 -----------------------------------
   If lDbVerAdm = 316 And lUpdOK = True Then
   
      'Agregamos campos IdCompAperTrib y NCompAperTrib a EmpresasAno
      Set Tbl = DbMain.TableDefs("EmpresasAno")

      Err.Clear
      Set Fld = Tbl.CreateField("NCompAperTrib", dbLong)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.NCompAperTrib", vbExclamation
         lUpdOK = False

      End If
      
      Err.Clear
      Set Fld = Tbl.CreateField("IdCompAperTrib", dbLong)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.IdCompAperTrib", vbExclamation
         lUpdOK = False

      End If
      
      Set Tbl = Nothing

   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 317
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V316 = lUpdOK

End Function



Private Function CorrigeBaseAdm_V315() As Boolean     'entregada 7 oct 2013
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 315 -----------------------------------
   If lDbVerAdm = 315 And lUpdOK = True Then
   
      'se agrega campos AjustesIFRS y CaclPropIVA a tabla ControlEmpresa
      Set Tbl = DbMain.TableDefs("ControlEmpresa")
      
      Err.Clear
      Set Fld = Tbl.CreateField("AjustesIFRS", dbByte)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
   
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.AjustesIFRS", vbExclamation
         lUpdOK = False
      
      End If
      Err.Clear
      
      Set Fld = Tbl.CreateField("CalcPropIVA", dbByte)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
   
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "ControlEmpresa.CalcPropIVA", vbExclamation
         lUpdOK = False
      
      End If
            
      Call UpdCodDocSII


   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 316
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V315 = lUpdOK

End Function


Private Function CorrigeBaseAdm_V314() As Boolean     'entregada 10 sep 2013
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 314 -----------------------------------
   If lDbVerAdm = 314 And lUpdOK = True Then
   
      'se agrega campo AceptaPropIVA a tabla TipoDocs
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("AceptaPropIVA", dbBoolean)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
   
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.AceptaPropIVA", vbExclamation
         lUpdOK = False
      
      End If
      
      'seteamos el campo para los tipos de docs indicados
      
      Q1 = "UPDATE TipoDocs SET AceptaPropIVA = 1 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo IN ('FAC', 'NCC', 'NDC', 'FCC', 'IMP', 'LFC', 'NCF', 'NDF', 'FIC')"
      
      Call ExecSQL(DbMain, Q1)
      

   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 315
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V314 = lUpdOK

End Function


Private Function CorrigeBaseAdm_V313() As Boolean     '19 ago 2013
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 313 -----------------------------------
   If lDbVerAdm = 313 And lUpdOK = True Then
   
      'se agregan campos CodDocSII y CodDocDTESII a tabla TipoDocs
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodDocSII", dbText, 3)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
   
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodDocSII", vbExclamation
         lUpdOK = False
      
      End If
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodDocDTESII", dbText, 3)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
   
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodDocDTESII", vbExclamation
         lUpdOK = False
      
      End If
      
      'se agrega campo CodImpSII  a tabla TipoValor
      Set Tbl = DbMain.TableDefs("TipoValor")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodImpSII", dbText, 3)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
   
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.CodImpSII", vbExclamation
         lUpdOK = False
      
      End If
           
      Call AlterField(DbMain, "Param", "Codigo", dbLong) ' pam: 19 ago 2013
   
   '--------------------- Actualización Versión -------------------------
   
   
      If lUpdOK Then
         lDbVerAdm = 314
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V313 = lUpdOK

End Function


Private Function CorrigeBaseAdm_V312() As Boolean     '31 jul 2013
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 312 -----------------------------------
   If lDbVerAdm = 312 And lUpdOK = True Then
   
      'modificación nombre de cuenta en plan avanzado
   
      Q1 = "UPDATE PlanAvanzado SET Descripcion = 'Intereses Diferidos Leasing L.P.' "
      Q1 = Q1 & " WHERE Codigo = '1031202'"
      Call ExecSQL(DbMain, Q1)
   
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 313
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V312 = lUpdOK

End Function


Private Function CorrigeBaseAdm_V311() As Boolean     '30 mayo 2013
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 311 -----------------------------------
   If lDbVerAdm = 311 And lUpdOK = True Then
   
      'modificación datos tabla IFRS_EstadoSituacionFinanciera
   
      Q1 = "UPDATE IFRS_EstadoSituacionFinanciera SET IdPadre = 27 "
      Q1 = Q1 & " WHERE Codigo = '2102000'"                             'Pasivos no Corrientes a nivel 3
      Call ExecSQL(DbMain, Q1)
   
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 312
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V311 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V310() As Boolean     '7 mayo 2013
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 310 -----------------------------------
   If lDbVerAdm = 310 And lUpdOK = True Then
   
      'modificaciones en códigos de traspaso F29
   
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 6111 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'BEX' "
      Call ExecSQL(DbMain, Q1)
   
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 311
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V310 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V309() As Boolean     '7 mayo 2013
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 309 -----------------------------------
   If lDbVerAdm = 309 And lUpdOK = True Then
   
      'modificaciones en códigos de traspaso F29
   
      Q1 = "UPDATE TipoDocs SET CodF29Count = 0, CodF29Neto = 0, CodF29ExCount = 586, CodF29Exento = 142 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'BOE'"
      Call ExecSQL(DbMain, Q1)
         
      Q1 = "UPDATE TipoDocs SET CodF29ExCount = 586, CodF29Exento = 142 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'BEX' "
      Call ExecSQL(DbMain, Q1)
   
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 310
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V309 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V308() As Boolean     '6 mayo 2013
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 308 -----------------------------------
   If lDbVerAdm = 308 And lUpdOK = True Then
   
      'modificaciones en códigos de traspaso F29
   
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 0, CodF29Exento = 562 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'FCE'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 0, CodF29Exento = 560 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'IEX'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 0, CodF29Exento = 142 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FVE'"
      Call ExecSQL(DbMain, Q1)
   
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 0, CodF29Exento = 20 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo IN ('EXP','NDE') "
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 0, CodF29Exento = -20 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'NCE' "
      Call ExecSQL(DbMain, Q1)
   
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 309
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V308 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V307() As Boolean     '26 sep 2012
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 307 -----------------------------------
   If lDbVerAdm = 307 And lUpdOK = True Then
   
      'modificaciones en códigos de traspaso F29
   
      'Liquidación Factura de Compra
      Q1 = "UPDATE TipoDocs SET CodF29Count = 519, CodF29IVA = 520 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'LFC'"
      Call ExecSQL(DbMain, Q1)
   
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 308
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V307 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V306() As Boolean   'Versión 4.0.0    26 sep 2012
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field

   

   On Error Resume Next

   If lDbVerAdm = 306 And lUpdOK = True Then
   
      'Insertamos la tabla con los reportes IFRS

      Call InsertTblIFRS
      
      'Agregamos los campos de IFRS en la tabla de Plan Avanzado

      Set Tbl = DbMain.TableDefs("PlanAvanzado")

      Err.Clear
      Set Fld = Tbl.CreateField("CodIFRS_EstRes", dbText, 15)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanAvanzado.CodIFRS_EstRes", vbExclamation
         lUpdOK = False

      End If

      Err.Clear
      Set Fld = Tbl.CreateField("CodIFRS_EstFin", dbText, 15)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanAvanzado.CodIFRS_EstFin", vbExclamation
         lUpdOK = False

      End If
      
      'Agregamos los campos de IFRS en la tabla de Plan Intermedio

      Set Tbl = DbMain.TableDefs("PlanIntermedio")

      Err.Clear
      Set Fld = Tbl.CreateField("CodIFRS_EstRes", dbText, 15)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanIntermedio.CodIFRS_EstRes", vbExclamation
         lUpdOK = False

      End If

      Err.Clear
      Set Fld = Tbl.CreateField("CodIFRS_EstFin", dbText, 15)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanIntermedio.CodIFRS_EstFin", vbExclamation
         lUpdOK = False

      End If
      
      'Agregamos los campos de IFRS en la tabla de Plan Básico

      Set Tbl = DbMain.TableDefs("PlanBasico")

      Err.Clear
      Set Fld = Tbl.CreateField("CodIFRS_EstRes", dbText, 15)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanBasico.CodIFRS_EstRes", vbExclamation
         lUpdOK = False

      End If

      Err.Clear
      Set Fld = Tbl.CreateField("CodIFRS_EstFin", dbText, 15)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh

      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PlanBasico.CodIFRS_EstFin", vbExclamation
         lUpdOK = False

      End If
      
      Call AsignarCodIFRSaPlanesPreDefinidos

      If lUpdOK Then
         lDbVerAdm = 307
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)

      End If
   End If

  CorrigeBaseAdm_V306 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V305() As Boolean   'Version 3.0, 16 abr 2012
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String

   


   On Error Resume Next

   '--------------------- Versión 305 -----------------------------------
   If lDbVerAdm = 305 And lUpdOK = True Then
   
      'eliminamos los códigos F22 637 y 638 de los planes predefinidos
      Q1 = "UPDATE PlanAvanzado SET CodF22 =  0 WHERE CodF22 IN(637, 638)"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanIntermedio SET CodF22 =  0 WHERE CodF22 IN(637, 638)"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanBasico SET CodF22 =  0 WHERE CodF22 IN(637, 638)"
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVerAdm = 306
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)

      End If
      
   End If

  CorrigeBaseAdm_V305 = lUpdOK



End Function

Public Function CorrigeBaseAdm_V304() As Boolean   'Version 3.0, 24 nov 2011
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field

   
   
   On Error Resume Next
   
   If lDbVerAdm = 304 And lUpdOK = True Then
   
      Call UpdateIPCUTM2009_10_11
    
      If lUpdOK Then
         lDbVerAdm = 305
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V304 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V303() As Boolean   'Version 3.0, 23 nov 2011
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field

   
   
   On Error Resume Next
   
   If lDbVerAdm = 303 And lUpdOK = True Then
   
      Set Tbl = DbMain.TableDefs("RazonesFin")
      
      Err.Clear
      Set Fld = Tbl.CreateField("Glosa", dbText, 120)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
   
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "RazonesFin.Glosa", vbExclamation
         lUpdOK = False
      
      End If
     
      If lUpdOK Then
         lDbVerAdm = 304
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V303 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V302() As Boolean   'Version 3.0, 6 sept. 2011
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field

   
   
   On Error Resume Next
   
   If lDbVerAdm = 302 And lUpdOK = True Then
   
      'faltó una h en Whisky
      Q1 = "UPDATE TipoValor SET Valor = 'Impto. Pisco, Licores, Whisky, Aguard.'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'Impto. Pisco, Licores, Wisky, Aguard.'"
      Call ExecSQL(DbMain, Q1)
      
      'se vuelve a agregar el campo RaZonFija para evitar proplemas con testing en Legal Publishing
      Set Tbl = DbMain.TableDefs("RazonesFin")
      
      Err.Clear
      Set Fld = Tbl.CreateField("RazonFija", dbBoolean)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
   
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "RazonesFin.RazonFija", vbExclamation
         lUpdOK = False
      
      End If
     
      If lUpdOK Then
         lDbVerAdm = 303
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V302 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V301() As Boolean   'Version 3.0, 18 agosto 2011
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field

   
   On Error Resume Next
   
   If lDbVerAdm = 301 And lUpdOK = True Then
   
      Err.Clear
      
      Set Tbl = DbMain.TableDefs("TipoValor")
      
      'agregamos campos Tit1, Tit2
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("Tit1", dbText, 25)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.Tit1", vbExclamation
         lUpdOK = False
      End If
   
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("Tit2", dbText, 25)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.Tit2", vbExclamation
         lUpdOK = False
      End If
   
   
      'asignamos campos Tit1 y Tit2
      
      'Libro de Compras

   
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'IVA'"
      Q1 = Q1 & ", Tit2 = 'Irrecuper.'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAIRREC
      
      Call ExecSQL(DbMain, Q1)
   
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'IVA'"
      Q1 = Q1 & ", Tit2 = 'Act. Fijo'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAACTFIJO
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'IVA Ret.'"
      Q1 = Q1 & ", Tit2 = 'Parcial'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVARETPARC
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'IVA Ret.'"
      Q1 = Q1 & ", Tit2 = 'Total'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVARETTOT
      
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Imp. Pisco'"
      Q1 = Q1 & ", Tit2 = 'Licores'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPPISCO
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Imp. Vinos'"
      Q1 = Q1 & ", Tit2 = 'Champaña'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPVINOS
      
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Impto. '"
      Q1 = Q1 & ", Tit2 = 'Cervezas'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPCERVEZA
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Imp. Beb.'"
      Q1 = Q1 & ", Tit2 = 'Analcohól.'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPBEBANALC
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'ILA Notas'"
      Q1 = Q1 & ", Tit2 = 'Débito'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_ILANOTASDEB

      Call ExecSQL(DbMain, Q1)

      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'ILA Notas'"
      Q1 = Q1 & ", Tit2 = 'Crédito'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_ILANOTASCRED

      Call ExecSQL(DbMain, Q1)
      
'      Q1 = "UPDATE TipoValor SET "
'      Q1 = Q1 & "  Tit1 = 'IVA Antic.'"
'      Q1 = Q1 & ", Tit2 = 'Harina'"
'      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAANTICIPHARINA
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'IVA Antic.'"
      Q1 = Q1 & ", Tit2 = 'Carne'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAANTICIPCARNE
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Imp. Esp'"
      Q1 = Q1 & ", Tit2 = 'Petróleo'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPESPDIESEL
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Imp. Esp'"
      Q1 = Q1 & ", Tit2 = 'Petr. Gral.'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPESPPETRGRAL
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Imp. Esp'"
      Q1 = Q1 & ", Tit2 = 'Trans.Carga'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPESPDIESELTRANS            'LIBCOMPRAS_IMPESPPETRTRANS
      
      Call ExecSQL(DbMain, Q1)
      
      'Libro de Ventas
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Rebaja'"
      Q1 = Q1 & ", Tit2 = '65%'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_REBAJA65
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'IVA Ret.'"
      Q1 = Q1 & ", Tit2 = 'Parcial'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IVARETPARC
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'IVA Ret.'"
      Q1 = Q1 & ", Tit2 = 'Total'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IVARETTOT
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Ret. Márgen'"
      Q1 = Q1 & ", Tit2 = 'Comercializ.'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_RETMARGENCOM
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Imp. Pisco'"
      Q1 = Q1 & ", Tit2 = 'Licores'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPPISCO
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Imp. Vinos'"
      Q1 = Q1 & ", Tit2 = 'Champaña'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPVINOS
      
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Impto.'"
      Q1 = Q1 & ", Tit2 = 'Cerveza'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPCERVEZA
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Imp. Beb.'"
      Q1 = Q1 & ", Tit2 = 'Analcohól.'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPBEBANHALC
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'ILA Notas'"
      Q1 = Q1 & ", Tit2 = 'Débito'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_ILANOTASDEB
      
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'ILA Notas'"
      Q1 = Q1 & ", Tit2 = 'Crédito'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_ILANOTASCRED
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Imp. Art.37'"
      Q1 = Q1 & ", Tit2 = 'e) h) i) l)'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPART37E
   
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Imp. Art.37'"
      Q1 = Q1 & ", Tit2 = 'j)'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IMPART37J
     
      Call ExecSQL(DbMain, Q1)
      
'      Q1 = "UPDATE TipoValor SET "
'      Q1 = Q1 & "  Tit1 = 'Ret. Antic.'"
'      Q1 = Q1 & ", Tit2 = 'Harina'"
'      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_RETANTCAMBIOSUJHARINA
'
'      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET "
      Q1 = Q1 & "  Tit1 = 'Ret. Antic.'"
      Q1 = Q1 & ", Tit2 = 'Carne'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_RETANTCAMBIOSUJCARNE
   
      Call ExecSQL(DbMain, Q1)
     
      If lUpdOK Then
         lDbVerAdm = 302
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V301 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V300() As Boolean   'Version 3.0, 10 agosto 2011
   Dim Q1 As String
   Dim Rs As Recordset
   Dim MaxTipoDoc As Long

   
   On Error Resume Next
   
   If lDbVerAdm = 300 And lUpdOK = True Then
   
      'agregamos tipo doc Maquina registradora en Libro Ventas
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = '" & TDOC_MAQREGISTRADORA & "'"
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
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Máquina Registradora', '" & TDOC_MAQREGISTRADORA & "', 'ACTIVO', -1, -1, -1, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
   
   
       If lUpdOK Then
         lDbVerAdm = 301
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V300 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V299() As Boolean   'Version 3.0, 13 julio 2011
   Dim Q1 As String

   
   On Error Resume Next
   
   If lDbVerAdm <= 299 And lUpdOK = True Then
   
      Call CrearTblRazFin
      Call DefTiposRazFin
      Call DefRazonesFin
   
       If lUpdOK Then
         lDbVerAdm = 300
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
  
  CorrigeBaseAdm_V299 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V48() As Boolean   '15 abril 2011
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim StrCod As String

   On Error Resume Next

   If lDbVerAdm = 48 And lUpdOK = True Then

      'eliminamos algunos códigos de exportación a Form 22, de acuerdo a lo solicitado por Victor Morales
      'en reporte 54-B, 7 abr 2011
      
      StrCod = "239, 240, 778, 779, 816, 817, 857, 858, 861"
      
      Q1 = "UPDATE PlanBasico SET CodF22 = 0 WHERE CodF22 IN(" & StrCod & ")"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanIntermedio SET CodF22 = 0 WHERE CodF22 IN(" & StrCod & ")"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanAvanzado SET CodF22 = 0 WHERE CodF22 IN(" & StrCod & ")"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVerAdm = 49
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V48 = lUpdOK

End Function


Public Function CorrigeBaseAdm_V47() As Boolean   '9 mar 2011
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim MaxTipoDoc As Long

   On Error Resume Next

   

   If lDbVerAdm = 47 And lUpdOK = True Then   '9 mar 2011

      'agregamos tipo doc documento "FIC" Factura de Inicio a libro de Compras
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'FIC'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)

         'obtenemos el máximo para Compras
         Q1 = "SELECT Max(TipoDoc) FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         Call CloseRs(Rs)

         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, CodF29Count, CodF29Neto, CodF29IVA, CodF29AFCount, CodF29AFIVA, TieneAfecto, TieneExento, ExigeRUT)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & MaxTipoDoc & ", 'Factura de Inicio', 'FIC', 'ACTIVO', -1, 5191, 5201, 0, 524, 525, -1, -1, -1)"
         Call ExecSQL(DbMain, Q1)

      Else
         Call CloseRs(Rs)

      End If


      'agregamos tipo doc documento "FIV" Factura de Inicio a libro de Ventas
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FIV'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)

         'obtenemos el máximo para Ventas
         Q1 = "SELECT Max(TipoDoc) FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         Call CloseRs(Rs)

         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, CodF29Count, CodF29Neto, CodF29IVA, TieneAfecto, TieneExento, ExigeRUT)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Factura de Inicio', 'FIV', 'ACTIVO', -1, 515, 587, 5871, -1, -1, -1)"
         Call ExecSQL(DbMain, Q1)

      Else
         Call CloseRs(Rs)

      End If

      If lUpdOK Then
         lDbVerAdm = 48
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V47 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V46() As Boolean     '1 sept. 2010 (v 2.0.8)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 46 -----------------------------------
   If lDbVerAdm = 46 And lUpdOK = True Then
   
      'modificaciones en códigos de traspaso F29
   
      'Nota de crédito de ventas No Giro
      Q1 = "UPDATE TipoDocs SET CodF29CountNoGiro = 733, CodF29IVANoGiro = 734, CodF29NetoNoGiro = 5734 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'NCV'"
      Call ExecSQL(DbMain, Q1)
       
      'Nota de Crédito Factura de Compra (Retenido Parcial)
      Q1 = "UPDATE TipoDocs SET CodF29CountRetParcial = 0, CodF29NetoRetParcial = 0, CodF29DifIVARetParcial = 736 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'NCF'"
      Call ExecSQL(DbMain, Q1)
       
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 47
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V46 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V45() As Boolean     '31 ago 2010 (v 2.0.8)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next
   
   
   '--------------------- Versión 45 -----------------------------------
   If lDbVerAdm = 45 And lUpdOK = True Then
   
   'agregamos cuenta Seguro de Cesantía por Pagar a planes Básico, Intermedio y Avanzado
   
      Q1 = "SELECT IdCuenta FROM PlanBasico WHERE Codigo='2011200'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         IdCtaPadre = vFld(Rs("IdCuenta"))

         Q1 = "INSERT INTO PlanBasico "
         Q1 = Q1 & " (idPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22, Atrib" & ATRIB_CAPITALPROPIO & ")"
         Q1 = Q1 & " VALUES(" & IdCtaPadre & ", '2011206', 'SCESANTIA', 'Seguro de Cesantía por Pagar',"
         Q1 = Q1 & " ' ', 4, " & ECTA_ACTIVA & ", " & CLASCTA_PASIVO & ", 0, 0, 0, " & CAPPROPIO_PASIVO_EXIGIBLE & ", 0, 1)"
         Call ExecSQL(DbMain, Q1)
      End If
      
      Call CloseRs(Rs)
      
      Q1 = "SELECT IdCuenta FROM PlanIntermedio WHERE Codigo='2011200'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         IdCtaPadre = vFld(Rs("IdCuenta"))

         Q1 = "INSERT INTO PlanIntermedio "
         Q1 = Q1 & " (idPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22, Atrib" & ATRIB_CAPITALPROPIO & ")"
         Q1 = Q1 & " VALUES(" & IdCtaPadre & ", '2011206', 'SCESANTIA', 'Seguro de Cesantía por Pagar',"
         Q1 = Q1 & " ' ', 4, " & ECTA_ACTIVA & ", " & CLASCTA_PASIVO & ", 0, 0, 0, " & CAPPROPIO_PASIVO_EXIGIBLE & ", 0, 1)"
         Call ExecSQL(DbMain, Q1)
      End If
      
      Call CloseRs(Rs)
      
      Q1 = "SELECT IdCuenta FROM PlanAvanzado WHERE Codigo='2011200'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         IdCtaPadre = vFld(Rs("IdCuenta"))

         Q1 = "INSERT INTO PlanAvanzado "
         Q1 = Q1 & " (idPadre, Codigo, Nombre, Descripcion, CodFECU, Nivel, Estado, Clasificacion, Debe, Haber, MarcaApertura, TipoCapPropio, CodF22, Atrib" & ATRIB_CAPITALPROPIO & ")"
         Q1 = Q1 & " VALUES(" & IdCtaPadre & ", '2011206', 'SCESANTIA', 'Seguro de Cesantía por Pagar',"
         Q1 = Q1 & " ' ', 4, " & ECTA_ACTIVA & ", " & CLASCTA_PASIVO & ", 0, 0, 0, " & CAPPROPIO_PASIVO_EXIGIBLE & ", 0, 1)"
         Call ExecSQL(DbMain, Q1)
      End If
      
      Call CloseRs(Rs)
       
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 46
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V45 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V44() As Boolean     '19 ago 2010 (v 2.0.8)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer


   On Error Resume Next
   
   
   '--------------------- Versión 44 -----------------------------------
   If lDbVerAdm = 44 And lUpdOK = True Then
   
      'agregamos traspaso de Neto para diversos documentos, por problema de redondeo, a IVA Estándar
         
      Err.Clear
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 5501 WHERE Diminutivo = 'LFV'"
      Call ExecSQL(DbMain, Q1)
   
      Err.Clear
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 6111 WHERE Diminutivo = 'BOV'"
      Call ExecSQL(DbMain, Q1)
   
      Err.Clear
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 5510, CodF29NetoNoGiro = 5717 WHERE Diminutivo = 'NCV'"
      Call ExecSQL(DbMain, Q1)
   
      Err.Clear
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 5513, CodF29NetoNoGiro = 5717 WHERE Diminutivo = 'NDV'"
      Call ExecSQL(DbMain, Q1)
   
      Err.Clear
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 5502, CodF29NetoNoGiro = 5717 WHERE Diminutivo = 'FAV'"
      Call ExecSQL(DbMain, Q1)
  
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 45
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V44 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V43() As Boolean     '28 ene 2010 (v 2.0.7)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer


   On Error Resume Next
   
   
   '--------------------- Versión 43 -----------------------------------
   If lDbVerAdm = 43 And lUpdOK = True Then
   
      'agregamos traspaso de Retención de Boletas a Terceros (BRT) IVA Estándar
         
      Err.Clear
      Q1 = "UPDATE TipoDocs SET CodF29RetHon = 151 WHERE Diminutivo = 'BRT'"
      Call ExecSQL(DbMain, Q1)
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 44
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V43 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V42() As Boolean     '30 dic.2009 (v 2.0.7)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer


   On Error Resume Next
   
   
   '--------------------- Versión 42 -----------------------------------
   If lDbVerAdm = 42 And lUpdOK = True Then
         
      Err.Clear
      
      Set Tbl = DbMain.TableDefs("Empresas")
      
      'agregamos campo RutDisp para almacenar el Rut que se muestra en el membrete de los reportes, con el fin de permitir más de una empresa con el mismo RUT, sólo para la Asoc. de AFP
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("RutDisp", dbText, 20)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "Empresas.RutDisp", vbExclamation
         lUpdOK = False
      End If
      
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 43
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V42 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V41() As Boolean     ' 4 dic 2009
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   Dim IdMoneda As Long
   Dim TDocsCompras As String

   On Error Resume Next
   
   
   
   '--------------------- Versión 41 -----------------------------------
   If lDbVerAdm = 41 And lUpdOK = True Then
      
      'obtenemos el tipo de documento para FAC para asignarlo al campo TipoDoc de la tabla TipoValor
      TDocsCompras = ","
      Q1 = "SELECT TipoDoc FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS & " AND Diminutivo IN( 'FAC') ORDER BY TipoDoc"
      Set Rs = OpenRs(DbMain, Q1)
      Do While Rs.EOF = False
         TDocsCompras = TDocsCompras & vFld(Rs(0)) & ","
         Rs.MoveNext
      Loop
      Call CloseRs(Rs)
      
      'agregamos un registro a tabla TipoValor para libro Compras - Impuesto Petróleo Diesel General
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'Imp. Esp. al Petr. Diesel General'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPESPPETRGRAL & ", 'Imp. Esp. al Petr. Diesel General', ' ', ' ', -1, 127, ',1,')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
                                
      'agregamos un registro a tabla TipoValor para libro Compras - Impuesto Petróleo Transportista Carga
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'Imp. Esp. al Petr. Transp. Carga'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPESPDIESELTRANS & ", 'Imp. Esp. al Petr. Transp. Carga', ' ', ' ', -1, 544, ',1,')"      'LIBCOMPRAS_IMPESPPETRTRANS
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 42
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V41 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V40() As Boolean     ' 21 sep 2009
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   Dim IdMoneda As Long

   On Error Resume Next
   
   
   
   '--------------------- Versión 40 -----------------------------------
   If lDbVerAdm = 40 And lUpdOK = True Then
      
      Q1 = "DROP INDEX Pc ON PcUsr"
      Call ExecSQL(DbMain, Q1)
      
      Err.Clear
      Set Tbl = DbMain.TableDefs("PcUsr")
      
      Err.Clear
      Set Fld = Tbl.CreateField("Pid", dbLong)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "PcUsr.Pid", vbExclamation
         lUpdOK = False
      
      End If
      
      Q1 = "CREATE UNIQUE INDEX Pc ON PcUsr( Pc, Usr )"
      Call ExecSQL(DbMain, Q1)
                                
                           
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 41
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V40 = lUpdOK

End Function

Private Function CorrigeBaseAdm_V39() As Boolean     ' 21 ago 2009
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   Dim IdMoneda As Long

   On Error Resume Next
   
   
   
   '--------------------- Versión 39 -----------------------------------
   If lDbVerAdm = 39 And lUpdOK = True Then
      
      'insertamos nuevos puntos de IPC
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2009, 4, 1)) & ", 99.11, -0.2, 1.000)"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2009, 5, 1)) & ", 98.86, -0.3, 1.001)"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2009, 6, 1)) & ", 99.20, 0.3, 1.003)"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2009, 7, 1)) & ", 98.77, -0.4, 1.000)"
      Call ExecSQL(DbMain, Q1)
         
      'insertamos valores UTM
      Q1 = "SELECT IdMoneda FROM Monedas WHERE Simbolo='UTM'"
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         IdMoneda = vFld(Rs("IdMoneda"))
      End If
      Call CloseRs(Rs)
      
      If IdMoneda > 0 Then
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2009, 6, 1)) & ", 36792)"
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2009, 7, 1)) & ", 36682)"
         Call ExecSQL(DbMain, Q1)
      
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2009, 8, 1)) & ", 36792)"
         Call ExecSQL(DbMain, Q1)
      
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2009, 9, 1)) & ", 36645)"
         Call ExecSQL(DbMain, Q1)
      End If
                           
                           
                           
                           
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 40
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V39 = lUpdOK

End Function
Private Function CorrigeBaseAdm_V38() As Boolean     ' 8 jul 2009
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer


   On Error Resume Next
   
   
   
   '--------------------- Versión 38 -----------------------------------
   If lDbVerAdm = 38 And lUpdOK = True Then
      
      ' Tabla para evitar que dos usuarios se conecten al mismo tiempo en un mismo PC
      Q1 = "CREATE TABLE PcUsr ( PC char(30), Usr char(30), CONSTRAINT Pc PRIMARY KEY (Pc) )"
      Call ExecSQL(DbMain, Q1)
                           
      DbMain.TableDefs.Refresh
                          
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVerAdm = 39
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBaseAdm_V38 = lUpdOK

End Function










Public Function CorrigeBaseAdm_V37() As Boolean   '12 may 2009
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset

   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim MaxTipoDoc As Long
   
   On Error Resume Next

   
   
   If lDbVerAdm = 37 And lUpdOK = True Then   '12 may 2009
   
      'agregamos tipo doc documento "BEX" Boleta de Venta con Exento
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'BEX'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo para Compras
         Q1 = "SELECT Max(TipoDoc) FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         Call CloseRs(Rs)
         
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, CodF29Count, CodF29IVA, TieneAfecto, TieneExento, ExigeRUT)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Boleta de Venta con Exento', 'BEX', 'ACTIVO', -1, 110, 111, -1, -1, 0)"
         Call ExecSQL(DbMain, Q1)
      
      Else
         Call CloseRs(Rs)
      
      End If
      
            
      If lUpdOK Then
         lDbVerAdm = 38
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
            
   End If
     
   CorrigeBaseAdm_V37 = lUpdOK
   
End Function

Public Function CorrigeBaseAdm_V36() As Boolean   '8 may 2009
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset

   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   
   On Error Resume Next

   
   
   If lDbVerAdm = 36 And lUpdOK = True Then   '8 may 2009
      Call UpdateIPCUTM2009
            
      If lUpdOK Then
         lDbVerAdm = 37
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
            
   End If
     
   CorrigeBaseAdm_V36 = lUpdOK
   
End Function
Public Function CorrigeBaseAdm_V35() As Boolean   '19 dic 2008
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset

   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   
   On Error Resume Next

   
   
   If lDbVerAdm = 35 And lUpdOK = True Then   '19 dic 2008
      
      Q1 = "UPDATE TipoDocs SET CodF29CountDTE = 5111 WHERE CodF29CountDTE = 511"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET CodF29IVADTE = 511 WHERE CodF29IVADTE = 5111"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET CodF29IVADTE = -511 WHERE CodF29IVADTE = -5111"
      Call ExecSQL(DbMain, Q1)
            
      If lUpdOK Then
         lDbVerAdm = 36
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
            
   End If
     
   CorrigeBaseAdm_V35 = lUpdOK
   
End Function
Public Function CorrigeBaseAdm_V34() As Boolean   '15 oct 2008
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset

   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim TDocFCompra As String
   Dim MaxTipoDoc As Long
   Dim TDocsCompras As String
   Dim IdMoneda As Long
   
   On Error Resume Next

   

   Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      lDbVerAdm = 0
   Else
      lDbVerAdm = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)
   
   If lDbVerAdm = 34 And lUpdOK = True Then   '15 oct 2008
      
      Call CorrigeRegComunas
      
      Call CodActEconomicas
      
      If lUpdOK Then
         lDbVerAdm = 35
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
            
   End If
     
   CorrigeBaseAdm_V34 = lUpdOK
   
End Function
Public Function CorrigeBaseAdm_V33() As Boolean   '9 oct 2008
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim TDocFCompra As String
   Dim MaxTipoDoc As Long
   Dim TDocsCompras As String
   Dim IdMoneda As Long
   
   On Error Resume Next

   

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
   
   If DbVer = 33 And lUpdOK = True Then   '9 oct 2008
      
      'insertamos nuevos puntos de IPC
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2008, 7, 1)) & ", 141.28, 0.011, 1.021)"
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2008, 8, 1)) & ", 142.59, 0.009, 1.009)"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2008, 9, 1)) & ", 144.11, 0.011, 1)"
      Call ExecSQL(DbMain, Q1)
         
      'insertamos valores UTM
      Q1 = "SELECT IdMoneda FROM Monedas WHERE Simbolo='UTM'"
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         IdMoneda = vFld(Rs("IdMoneda"))
      End If
      Call CloseRs(Rs)
      
      If IdMoneda > 0 Then
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2008, 1, 1)) & ", 34496)"
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2008, 2, 1)) & ", 34668)"
         Call ExecSQL(DbMain, Q1)
      
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2008, 3, 1)) & ", 34668)"
         Call ExecSQL(DbMain, Q1)
      
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2008, 4, 1)) & ", 34807)"
         Call ExecSQL(DbMain, Q1)
      
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2008, 5, 1)) & ", 35085)"
         Call ExecSQL(DbMain, Q1)
      
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2008, 6, 1)) & ", 35225)"
         Call ExecSQL(DbMain, Q1)
      
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2008, 7, 1)) & ", 35648)"
         Call ExecSQL(DbMain, Q1)
      
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2008, 8, 1)) & ", 36183)"
         Call ExecSQL(DbMain, Q1)
      
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2008, 9, 1)) & ", 36581)"
         Call ExecSQL(DbMain, Q1)
      
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2008, 10, 1)) & ", 36910)"
         Call ExecSQL(DbMain, Q1)
      
         Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
         Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(2008, 11, 1)) & ", 37316)"
         Call ExecSQL(DbMain, Q1)
      
         'ponemos un decimal más a la UTM
         Q1 = "UPDATE Monedas SET DecInf = 3 WHERE IdMoneda=" & IdMoneda
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      If lUpdOK Then
         DbVer = DbVer + 1
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
            
   End If
     
   CorrigeBaseAdm_V33 = True
   
End Function
Public Function CorrigeBaseAdm_V32() As Boolean   '14 Jul 2008
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim TDocFCompra As String
   Dim MaxTipoDoc As Long
   Dim TDocsCompras As String
   
   On Error Resume Next

   

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
   
   If DbVer = 32 And lUpdOK = True Then   '14 Jul 2008
         
      'agregamos campo CodF29CountDTE a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29CountDTE", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29CountDTE", vbExclamation
         lUpdOK = False
      
      End If
      
      'agregamos campo CodF29NetoDTE a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29NetoDTE", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29NetoDTE", vbExclamation
         lUpdOK = False
      
      End If

      'agregamos campo CodF29IVAIrrecDTE a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29IVAIrrecDTE", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29IVAIrrecDTE", vbExclamation
         lUpdOK = False
      
      End If

      'agregamos campo CodF29CountIVAIrrec a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29CountIVAIrrec", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29CountIVAIrrec", vbExclamation
         lUpdOK = False
      
      End If

      'agregamos campo CodF29NetoIVAIrrec a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29NetoIVAIrrec", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29NetoIVAIrrec", vbExclamation
         lUpdOK = False
      
      End If

      'actualizamos campos anteriores
      
      'DTE
      Q1 = "UPDATE TipoDocs SET CodF29IVADTE = 5111, CodF29CountDTE = 511, CodF29NetoDTE = 514, CodF29IVAIrrecDTE = 5141 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo IN( 'FAC', 'NCC', 'NDC', 'FCC', 'IMP', 'NDF')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoDocs SET CodF29IVADTE = -5111, CodF29CountDTE = 511, CodF29NetoDTE = 514, CodF29IVAIrrecDTE = 5141 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo IN( 'NCF')"
      Call ExecSQL(DbMain, Q1)

      'IVA Irrecuperable
      Q1 = "UPDATE TipoDocs SET CodF29CountIVAIrrec = 564, CodF29NetoIVAIrrec = 521 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo IN( 'FAC', 'NCC', 'NDC')"
      Call ExecSQL(DbMain, Q1)

      Q1 = "UPDATE TipoDocs SET CodF29CountIVAIrrec = 566, CodF29NetoIVAIrrec = 560 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo IN( 'IMP')"
      Call ExecSQL(DbMain, Q1)

      'eliminamos IVA Irrecuperable de ventas
      Q1 = "DELETE * FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IVAIRREC
      Call ExecSQL(DbMain, Q1)
      
      'eliminamos IVA Irrecuperable Acitvo Fijo
      Q1 = "DELETE * FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAIRRACTFIJO
      Call ExecSQL(DbMain, Q1)
      
      'obtenemos los tipos de documentos para FAC, NCC, NDC Y LFC para asignarlos al campo TipoDoc de la tabla TipoValor
      TDocsCompras = ","
      Q1 = "SELECT TipoDoc FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS & " AND Diminutivo IN( 'FAC', 'NCC', 'NDC', 'IMP' )"
      Set Rs = OpenRs(DbMain, Q1)
      Do While Rs.EOF = False
         TDocsCompras = TDocsCompras & vFld(Rs(0)) & ","
         Rs.MoveNext
      Loop
      Call CloseRs(Rs)

      'asignamos estos documentos a TipoValor IVAIrrecuperable
      Q1 = "UPDATE TipoValor SET TipoDoc = '" & TDocsCompras & "' WHERE TipoLib=" & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAIRREC
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         DbVer = DbVer + 1
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
            
   End If
     
   CorrigeBaseAdm_V32 = True
   
End Function
Public Function CorrigeBaseAdm_V31() As Boolean   '10 Jul 2008
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim TDocFCompra As String
   Dim MaxTipoDoc As Long
   Dim TDocsCompras As String
   
   On Error Resume Next

   

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
   
   If DbVer = 31 And lUpdOK = True Then   '10 Jul 2008
         
      'insertamos IPC dic 2007 a junio 2008
   
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2007, 12, 1)) & ", 133.95, 0, 0)"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2008, 1, 1)) & ", 133.90, 0, 1.028)"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2008, 2, 1)) & ", 134.44, 0.004, 1.028)"
      Call ExecSQL(DbMain, Q1)
   
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2008, 3, 1)) & ", 135.56, 0.008, 1.024)"
      Call ExecSQL(DbMain, Q1)
   
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2008, 4, 1)) & ", 136.08, 0.004, 1.015)"
      Call ExecSQL(DbMain, Q1)
   
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2008, 5, 1)) & ", 137.65, 0.012, 1.012)"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
      Q1 = Q1 & " VALUES(" & CLng(DateSerial(2008, 6, 1)) & ", 139.7, 0.015, 1)"
      Call ExecSQL(DbMain, Q1)
   
      'asignamos código a campos nuevos para Factura de Compras del Libro de Ventas
      Q1 = "UPDATE TipoDocs SET CodF29Count = 5191, CodF29Neto = 5201, CodF29IVA = 0, CodF29CountRetParcial = 0, CodF29NetoRetParcial = 0, CodF29DifIVARetParcial = 0 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo IN( 'FCC', 'NDF')"
      Call ExecSQL(DbMain, Q1)

      Q1 = "UPDATE TipoDocs SET CodF29Count = 5191, CodF29Neto = -5201, CodF29IVA = 0, CodF29CountRetParcial = 0, CodF29NetoRetParcial = 0, CodF29DifIVARetParcial = 0 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'NCF'"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         DbVer = DbVer + 1
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
     
   CorrigeBaseAdm_V31 = True
   
End Function
Public Function CorrigeBaseAdm_V30() As Boolean   '8 Jul 2008
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim TDocFCompra As String
   Dim MaxTipoDoc As Long
   Dim TDocsCompras As String
   
   On Error Resume Next

   

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
   
   If DbVer = 30 And lUpdOK = True Then   ' 8 Jul 2008
   
      'obtenemos el tipo de documento para FCC (Factura de Compra), NCF y NDF (N. Cred y Deb de Factura de Compra)
      TDocFCompra = ","
      Q1 = "SELECT TipoDoc FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS & " AND Diminutivo IN( 'FCC', 'NCF', 'NDF') ORDER BY TipoDoc"
      Set Rs = OpenRs(DbMain, Q1)
      Do While Not Rs.EOF
         TDocFCompra = TDocFCompra & vFld(Rs(0)) & ","
         Rs.MoveNext
      Loop
      Call CloseRs(Rs)
   
      'modificamos documentos asociados a FCC (Factura de Compra), NCF y NDF (N. Cred y Deb de Factura de Compra)
      Q1 = "UPDATE TipoValor SET TipoDoc = '" & TDocFCompra & "'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA Retenido Parcial'"
      Call ExecSQL(DbMain, Q1)

      Q1 = "UPDATE TipoValor SET TipoDoc = '" & TDocFCompra & "'"
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA Retenido Total'"
      Call ExecSQL(DbMain, Q1)
   
      'agregamos impuesto específio Diesel
   
      'obtenemos los tipos de documentos para FAC, NCC y NDC para asignarlos al campo TipoDoc de la tabla TipoValor
      TDocsCompras = ","
      Q1 = "SELECT TipoDoc FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS & " AND Diminutivo IN( 'FAC', 'NCC', 'NDC' ) ORDER BY TipoDoc"
      Set Rs = OpenRs(DbMain, Q1)
      Do While Rs.EOF = False
         TDocsCompras = TDocsCompras & vFld(Rs(0)) & ","
         Rs.MoveNext
      Loop
      Call CloseRs(Rs)
      
      'agregamos un registro a tabla TipoValor para libro Compras - Impuesto Petróleo Diesel
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'Imp. Esp. al Petróleo Diesel'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPESPDIESEL & ", 'Imp. Esp. al Petróleo Diesel', ' ', ' ', -1, 0, '" & TDocsCompras & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
   
      'agregamos tipo doc documento "CIT" Crédito Impto. Timbre y Estampilla
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'CIT'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo para Compras
         Q1 = "SELECT Max(TipoDoc) FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, CodF29Count, CodF29IVA, TieneAfecto, TieneExento, ExigeRUT)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & MaxTipoDoc & ", 'Crédito Impto. Timbre y Estampilla', 'CIT', 'ACTIVO', -1, 0, 0, 0, -1, 0)"
         Call ExecSQL(DbMain, Q1)
      
      Else
         Call CloseRs(Rs)
         
      End If
      
   
   
      If lUpdOK Then
         DbVer = DbVer + 1
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
     
   CorrigeBaseAdm_V30 = True
   
End Function
Public Function CorrigeBaseAdm_V29() As Boolean   '3 Jul 2008
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim MaxTipoDoc As Integer
   Dim MaxCod As Integer
   Dim TDocFCompra As Integer
   Dim TDocFCompraV As Integer
   Dim TDocsCompras As String
   Dim TDocsComprasSLiq As String
   Dim TDocsVentas As String
   Dim TDocsVentasSLiq As String

   On Error Resume Next

   

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
   
   If DbVer = 29 And lUpdOK = True Then   ' 3 Jul 2008
   
      MsgBox1 "Esta base de datos " & BD_COMUN & " está dañada, utilice un archivo nuevo.", vbExclamation
   
      Call CloseDb(DbMain)
      End
   
   End If
     
   CorrigeBaseAdm_V29 = True
   
End Function
Public Function CorrigeBaseAdm_V28() As Boolean   '16 may 2008
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim MaxTipoDoc As Integer
   Dim MaxCod As Integer
   Dim TDocFCompra As Integer
   Dim TDocFCompraV As Integer
   Dim TDocsCompras As String
   Dim TDocsComprasSLiq As String
   Dim TDocsVentas As String
   Dim TDocsVentasSLiq As String

   On Error Resume Next

   

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
   
   
   If DbVer = 28 And lUpdOK = True Then   ' 16 may 2008
   
      '--------------------------- Nuevos Documentos --------------------------------
      
   
      'agregamos tipo doc Liquidación Factura en Libro Compras (nuevamente)
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'LFC'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc) FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, CodF29Count, CodF29IVA, TieneAfecto, TieneExento, ExigeRUT)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & MaxTipoDoc & ", 'Liquidación Factura', 'LFC', 'ACTIVO', -1, 500, 501, -1, -1, -1)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
   
      'agregamos tipo doc Nota de Crédito por Factura de Compra en Libro Compras (nuevamente)
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'NCF'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc) FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, CodF29Count, CodF29IVA, CodF29IVADTE, TieneAfecto, TieneExento, ExigeRUT, EsRebaja)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & MaxTipoDoc & ", 'Nota de Crédito Fac. Compra', 'NCF', 'ACTIVO', -1, 519, -520, -511, -1, -1, -1, -1)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
   
      'agregamos tipo doc Nota de Débito por Factura de Compra en Libro Compras (nuevamente)
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'NDF'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = True Then  'no existe, lo agregamos
         
         Call CloseRs(Rs)
         
         'obtenemos el máximo
         Q1 = "SELECT Max(TipoDoc) FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS
         Set Rs = OpenRs(DbMain, Q1)
         MaxTipoDoc = 1
         If Rs.EOF = False Then
            MaxTipoDoc = vFld(Rs(0)) + 1
         End If
         
         'insertamos
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo,  CodF29Count, CodF29IVA, CodF29IVADTE, TieneAfecto, TieneExento, ExigeRUT, EsRebaja)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & MaxTipoDoc & ", 'Nota de Débito Fac. Compra', 'NDF', 'ACTIVO', -1, 519, 520, 511, -1, -1, -1, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
     
   
      '----------------------- Campos nuevos en tabla TipoDocs ------------------------------------
   
   
      'agregamos campo CodF29CountNoGiro a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29CountNoGiro", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29CountNoGiro", vbExclamation
         lUpdOK = False
      
      End If
   
      'agregamos campo CodF29NetoNoGiro a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29NetoNoGiro", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29NetoNoGiro", vbExclamation
         lUpdOK = False
      
      End If
   
      'agregamos campo CodF29IVANoGiro a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29IVANoGiro", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29IVANoGiro", vbExclamation
         lUpdOK = False
      
      End If
      
      'agregamos campo CodF29ExCountNoGiro a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29ExCountNoGiro", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29ExCountNoGiro", vbExclamation
         lUpdOK = False
      
      End If
      
      'agregamos campo CodF29ExentoNoGiro a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29ExentoNoGiro", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29ExentoNoGiro", vbExclamation
         lUpdOK = False
      
      End If
      
      'asignamos código a campos nuevos
      Q1 = "UPDATE TipoDocs SET CodF29CountNoGiro = 714, CodF29NetoNoGiro = 715, CodF29IVANoGiro = 0, "
      Q1 = Q1 & " CodF29ExCountNoGiro = 0, CodF29ExentoNoGiro = 0 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FVE'"
      Call ExecSQL(DbMain, Q1)

      Q1 = "UPDATE TipoDocs SET CodF29CountNoGiro = 716, CodF29NetoNoGiro = 0, CodF29IVANoGiro = 717, "
      Q1 = Q1 & " CodF29ExCountNoGiro = 714, CodF29ExentoNoGiro = 715 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo IN ( 'FAV', 'NCV', 'NDV' )"
      Call ExecSQL(DbMain, Q1)
   
   
      'agregamos campo CodF29CountRetParcial a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29CountRetParcial", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29CountRetParcial", vbExclamation
         lUpdOK = False
      
      End If
   
      'agregamos campo CodF29NetoRetParcial a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29NetoRetParcial", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29NetoRetParcial", vbExclamation
         lUpdOK = False
      
      End If
   
      'agregamos campo CodF29DifIVARetParcial a tabla TipoDocs
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDocs")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29DifIVARetParcial", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29DifIVARetParcial", vbExclamation
         lUpdOK = False
      
      End If
      
      'asignamos código a campos nuevos para Factura de Compras del Libro de Ventas
      Q1 = "UPDATE TipoDocs SET CodF29CountRetParcial = 516, CodF29NetoRetParcial = 720, CodF29DifIVARetParcial = 5171 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FCV'"
      Call ExecSQL(DbMain, Q1)
      
   
      '--------------------------- Campos Nuevos Tabla TipoValor --------------------------------
   
   
     
      'agregamos campo CodF29 a tabla TipoValor
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoValor")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.CodF29", vbExclamation
         lUpdOK = False
      
      End If
   
      'agregamos campo CodF29_Adic a tabla TipoValor (para impuestos que se traspasan a dos códigos del F29)
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoValor")
      
      Err.Clear
      Set Fld = Tbl.CreateField("CodF29_Adic", dbInteger)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.CodF29_Adic", vbExclamation
         lUpdOK = False
      
      End If
   
      'agregamos campo TipoDoc a tabla TipoValor
      Err.Clear
      Set Tbl = DbMain.TableDefs("TipoDoc")
      
      Err.Clear
      Set Fld = Tbl.CreateField("TipoDoc", dbText, 30)
      Tbl.Fields.Append Fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.TipoDoc", vbExclamation
         lUpdOK = False
      
      End If
   
   
      '--------------------------- Registros Nuevos en tabla TipoValor --------------------------------

      'agregamos varios registros a tabla TipoValor
      
      'LIBRO COMPRAS
      
      'obtenemos el máximo para Compras
      Q1 = "SELECT Max(Codigo) FROM TipoValor WHERE TipoLib=" & LIB_COMPRAS
      Set Rs = OpenRs(DbMain, Q1)
      MaxCod = 1
      If Rs.EOF = False Then
         MaxCod = vFld(Rs(0)) + 1
      End If
      Call CloseRs(Rs)

      'obtenemos el tipo de documento para FCC (Factura de Compra)
      Q1 = "SELECT TipoDoc FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS & " AND Diminutivo = 'FCC'"
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         TDocFCompra = vFld(Rs(0))
      End If
      Call CloseRs(Rs)

      
      'obtenemos los tipos de documentos para FAC, NCC, NDC Y LFC para asignarlos al campo TipoDoc de la tabla TipoValor
      TDocsCompras = ","
      Q1 = "SELECT TipoDoc FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS & " AND Diminutivo IN( 'FAC', 'NCC', 'NDC', 'LFC' )"
      Set Rs = OpenRs(DbMain, Q1)
      Do While Rs.EOF = False
         TDocsCompras = TDocsCompras & vFld(Rs(0)) & ","
         Rs.MoveNext
      Loop
      Call CloseRs(Rs)

      'obtenemos los tipos de documentos para FAC, NCC y NDC para asignarlos al campo TipoDoc de la tabla TipoValor
      TDocsComprasSLiq = ","
      Q1 = "SELECT TipoDoc FROM TipoDocs WHERE TipoLib=" & LIB_COMPRAS & " AND Diminutivo IN( 'FAC', 'NCC', 'NDC' )"
      Set Rs = OpenRs(DbMain, Q1)
      Do While Rs.EOF = False
         TDocsComprasSLiq = TDocsComprasSLiq & vFld(Rs(0)) & ","
         Rs.MoveNext
      Loop
      Call CloseRs(Rs)
      
      'agregamos registros a tabla TipoValor para libro Compras - Factura Compra
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA Irrec. Act. Fijo'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAIRRACTFIJO & ", 'IVA Irrec. Act. Fijo', ' ', ' ', -1, 0, '" & TDocsCompras & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA Retenido Parcial'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETPARC & ", 'IVA Retenido Parcial', ' ', ' ', -1, 554, '," & TDocFCompra & ",')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA Retenido Total'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETTOT & ", 'IVA Retenido Total', ' ', ' ', -1, 39, '," & TDocFCompra & ",')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA anticipado del periodo Harina'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAANTICIPHARINA & ", 'IVA anticipado del periodo Harina', ' ', ' ', -1, 5561, '" & TDocsComprasSLiq & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA anticipado del periodo Carne'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAANTICIPCARNE & ", 'IVA anticipado del periodo Carne', ' ', ' ', -1, 5562, '" & TDocsComprasSLiq & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      'agregamos TipoValor de libro Compras - cualquier doc
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'Impto. Pisco, Licores, Wisky, Aguard.'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos 2 códigos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, CodF29_Adic, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPPISCO & ", 'Impto. Pisco, Licores, Wisky, Aguard.', ' ', ' ', -1, 575, 576, '" & TDocsCompras & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
     
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'Impto. Vinos, Champaña, Chichas'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, CodF29_Adic, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPVINOS & ", 'Impto. Vinos, Champaña, Chichas', ' ', ' ', -1, 574, 33, '" & TDocsCompras & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
     
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'Impto. Cervezas'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, CodF29_Adic, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPCERVEZA & ", 'Impto. Cervezas', ' ', ' ', -1, 580, 149, '" & TDocsCompras & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'Impto. Bebidas Analcohólicas'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, CodF29_Adic, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPBEBANALC & ", 'Impto. Bebidas Analcohólicas', ' ', ' ', -1, 582, 85, '" & TDocsCompras & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
0      'Se elimina IVA retenido
      Q1 = "DELETE * FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVARETENIDO
      Call ExecSQL(DbMain, Q1)
      
      'Se elimina Anticipos
      Q1 = "DELETE * FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_ANTICIPOS
      Call ExecSQL(DbMain, Q1)
      
      'LIBRO VENTAS
      
      'obtenemos el máximo para Ventas
      Q1 = "SELECT Max(Codigo) FROM TipoValor WHERE TipoLib=" & LIB_VENTAS
      Set Rs = OpenRs(DbMain, Q1)
      MaxCod = 1
      If Rs.EOF = False Then
         MaxCod = vFld(Rs(0)) + 1
      End If
      Call CloseRs(Rs)
      
      'obtenemos el tipo de documento para FCV (Factura de Compra ventas)
      Q1 = "SELECT TipoDoc FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS & " AND Diminutivo = 'FCV'"
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         TDocFCompraV = vFld(Rs(0))
      End If
      Call CloseRs(Rs)

      TDocsVentas = ","
      
      'obtenemos los tipos de documentos para FAC, NCC, NDC Y LFV para asignarlos al campo TipoDoc de la tabla TipoValor
      Q1 = "SELECT TipoDoc FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS & " AND Diminutivo IN( 'FAV', 'NCV', 'NDV', 'LFV') "
      Set Rs = OpenRs(DbMain, Q1)
      Do While Rs.EOF = False
         TDocsVentas = TDocsVentas & vFld(Rs(0)) & ","
         Rs.MoveNext
      Loop
      Call CloseRs(Rs)
     
      'obtenemos los tipos de documentos para FAC, NCC y NDC para asignarlos al campo TipoDoc de la tabla TipoValor
      Q1 = "SELECT TipoDoc FROM TipoDocs WHERE TipoLib=" & LIB_VENTAS & " AND Diminutivo IN( 'FAV', 'NCV', 'NDV') "
      Set Rs = OpenRs(DbMain, Q1)
      Do While Rs.EOF = False
         TDocsVentasSLiq = TDocsComprasSLiq & vFld(Rs(0)) & ","
         Rs.MoveNext
      Loop
      Call CloseRs(Rs)
     
      'agregamos TipoValor de libro Ventas - Factura de Compra
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'IVA Retenido Parcial'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, CodF29_Adic, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IVARETPARC & ", 'IVA Retenido Parcial', ' ', ' ', -1, 517, 0, '," & TDocFCompraV & ",')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'IVA Retenido Total'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, CodF29_Adic, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IVARETTOT & ", 'IVA Retenido Total', ' ', ' ', -1, 5871, 0, '," & TDocFCompraV & ",')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
            
      'agregamos TipoValor de libro Ventas - factura de compra
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'Retención márgen de comercialización'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_RETMARGENCOM & ", 'Retención márgen de comercialización', ' ', ' ', -1, 597, '," & TDocFCompraV & ",')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
     
'      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'Ret. anticipo cambio sujeto Harina'"
'      Set Rs = OpenRs(DbMain, Q1)
'
'      If Rs.EOF = True Then  'no existe, lo agregamos
'
'         Call CloseRs(Rs)
'
'         'insertamos
'         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
'         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_RETANTCAMBIOSUJHARINA & ", 'Ret. anticipo cambio sujeto Harina', ' ', ' ', -1, 5551, '" & TDocsVentasSLiq & "')"
'         Call ExecSQL(DbMain, Q1)
'
'      End If
'
'      Call CloseRs(Rs)
'
'      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'Ret. anticipo cambio sujeto Carne'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_RETANTCAMBIOSUJCARNE & ", 'Ret. anticipo cambio sujeto Carne', ' ', ' ', -1, 5552, '" & TDocsVentasSLiq & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      'agregamos TipoValor de libro Ventas - cualquier doc
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'Impto. Pisco, Licores, Wisky, Aguard.'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IMPPISCO & ", 'Impto. Pisco, Licores, Wisky, Aguard.', ' ', ' ', -1, 577,  '" & TDocsVentas & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'Impto. Vinos, Champaña, Chichas'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IMPVINOS & ", 'Impto. Vinos, Champaña, Chichas', ' ', ' ', -1, 32, '" & TDocsVentas & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'Impto. Cervezas'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IMPCERVEZA & ", 'Impto. Cervezas', ' ', ' ', -1, 150, '" & TDocsVentas & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'Impto. Bebidas Analcohólicas'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IMPBEBANHALC & ", 'Impto. Bebidas Analcohólicas', ' ', ' ', -1, 146, '" & TDocsVentas & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'ILA por Notas de Débito emitidas'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_ILANOTASDEB & ", 'ILA por Notas de Débito emitidas', ' ', ' ', -1, 545, '" & TDocsVentas & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'ILA por Notas de Crédito emitidas'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_ILANOTASCRED & ", 'ILA por Notas de Crédito emitidas', ' ', ' ', -1, 546, '" & TDocsVentas & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
      
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'Impto. Adicional Art.37  e) h) i) l)'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IMPART37E & ", 'Impto. Adicional Art.37  e) h) i) l)', ' ', ' ', -1, 522, '" & TDocsVentas & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      MaxCod = MaxCod + 1
     
      Q1 = "SELECT IdTValor FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'Impto. Adicional Art.37  j)'"
      Set Rs = OpenRs(DbMain, Q1)

      If Rs.EOF = True Then  'no existe, lo agregamos

         Call CloseRs(Rs)
         
         'insertamos
         Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IMPART37J & ", 'Impto. Adicional Art.37  j)', ' ', ' ', -1, 526, '" & TDocsVentas & "')"
         Call ExecSQL(DbMain, Q1)

      End If
      
      Call CloseRs(Rs)
      
      '----------------------------- Diversos Ajustes -------------------------------------------------
      
      'asignamos código a rebaja 65%
      Q1 = "UPDATE TipoValor SET CodF29 = 126 WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_REBAJA65
      Call ExecSQL(DbMain, Q1)

     
      'Se elimina IVA retenido
      Q1 = "DELETE * FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_IVARETENIDO
      Call ExecSQL(DbMain, Q1)
      
      'Se elimina Retenciones
      Q1 = "DELETE * FROM TipoValor WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_RETENCIONES
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         DbVer = DbVer + 2 ' 30 *** caso especial, la 29 no corre
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
     
   CorrigeBaseAdm_V28 = True
   
End Function
Public Function CorrigeBaseAdm_V27() As Boolean   '11 mar 2008
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim MaxTipoDoc As Integer

   On Error Resume Next

   

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
   If DbVer = 27 And lUpdOK = True Then   ' 11 mar 2008
      Call CorrigeCodF22_2("PlanAvanzado")
      Call CorrigeCodF22_2("PlanIntermedio")
      Call CorrigeCodF22_2("PlanBasico")
     
     If lUpdOK Then
         DbVer = DbVer + 1
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
     
   CorrigeBaseAdm_V27 = True
   
End Function
Public Function CorrigeBaseAdm_V26() As Boolean   '25 Ago 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim MaxTipoDoc As Integer

   On Error Resume Next

   

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
   If DbVer = 26 And lUpdOK = True Then   ' 25 Ago 2006
     
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
         Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Venta sin Documento', 'VSD', 'ACTIVO', -1, 0, 0, 0, 0)"
         Call ExecSQL(DbMain, Q1)
         
      End If
      
      Call CloseRs(Rs)
     
     If lUpdOK Then
         DbVer = DbVer + 1
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
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long

   On Error Resume Next

   

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
   If DbVer = 25 And lUpdOK = True Then   ' 24 Abr 2006
      Call CorrigeTipoCapPropio_1("PlanAvanzado")
      Call CorrigeTipoCapPropio_1("PlanIntermedio")
      Call CorrigeTipoCapPropio_1("PlanBasico")
     
     If lUpdOK Then
         DbVer = DbVer + 1
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
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long

   On Error Resume Next

   

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
   If DbVer = 24 And lUpdOK = True Then   ' 30 Marzo 2006
      Call CorrigeCodF22_1("PlanAvanzado")
      Call CorrigeCodF22_1("PlanIntermedio")
      Call CorrigeCodF22_1("PlanBasico")
     
     If lUpdOK Then
         DbVer = DbVer + 1
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
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long

   On Error Resume Next

   

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
   If DbVer = 23 And lUpdOK = True Then   ' 17 Marzo 2006
     Call AlterField(DbMain, "CodActiv", "Codigo", dbText, 8)
     
     If lUpdOK Then
         DbVer = DbVer + 1
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
     
   CorrigeBaseAdm_V23 = lUpdOK
   
End Function

Public Function CorrigeBaseAdm_V22() As Boolean   '23 Enero 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long

   On Error Resume Next

   

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
   
   If DbVer = 22 And lUpdOK = True Then   ' 23 Ene 2006
   
   
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
      
      
      If lUpdOK Then
         DbVer = DbVer + 1
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
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim DbActMDB As String, ConnStr As String

   On Error Resume Next

   

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
   
   If (DbVer = 20 Or DbVer = 21) And lUpdOK = True Then   ' 30 Sept. 2005
   
   
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
         lUpdOK = False
      End If
      
      'agregamos campo CodF29Exento    (Cod29 para Total Exento de Docs que tienen Exento y Afecto)
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29Exento", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29Exento", vbExclamation
         lUpdOK = False
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
         lUpdOK = False
         MsgBox1 "No se encontró archivo " & DbActMDB & "." & vbNewLine & vbNewLine & "No se actualizaron los nuevos Códigos de Actividad Económica.", vbExclamation + vbOKOnly
         MsgBox1 "ATENCIÓN: La actualización ha sido interrumpida. " & vbLf & vbLf & "Solicite el archivo a personal de soporte y vuelva a ejectuar la aplicación.", vbExclamation + vbOKOnly
      
         Call CloseDb(DbMain)
         
         End
      
      End If
      
      '***
      
      If lUpdOK Then
         DbVer = 22
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
         
     
   CorrigeBaseAdm_V21 = lUpdOK
   
End Function
   
' Para hacer manteciones a ciertas tablas de la db LexContab, con manejo de versión
Public Function CorrigeBaseAdm_2005_01() As Boolean
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DbVer As Integer, UpdOK As Boolean
   Dim id As Long
   Dim i As Integer
   Dim Rc As Long
   Dim MaxTipoDoc As Integer
   Dim IdCtaPadre As Long

   On Error Resume Next

   Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs Is Nothing Then
      MsgBox1 "La base de datos está corrupta o es muy antigua.", vbCritical
      Call CloseDb(DbMain)
      End
   End If
      
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

   If DbVer = 0 And lUpdOK = True Then
   
      '--------------------- EmpresasAno -----------------------------------

      Set Tbl = DbMain.TableDefs("EmpresasAno")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("IdCompAper", dbLong)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.IdCompAper", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("NumLastCompUnico", dbLong)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.NumLastCompUnico", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("NumLastCompA", dbLong)  'número último comp. de apertura

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.NumLastA", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("NumLastCompE", dbLong)  'número último comp. de egreso

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.NumLastCompE", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("NumLastCompI", dbLong)  'número último comp. de ingreso

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.NumLastCompI", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("NumLastCompT", dbLong)  'número último comp. de traspaso

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "EmpresasAno.NumLastCompT", vbExclamation
         lUpdOK = False
      End If

      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         DbVer = 1
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 1 -----------------------------------

   If DbVer = 1 And lUpdOK = True Then
   
      Q1 = "CREATE TABLE LParam (Codigo SMALLINT, Valor char(255) NULL )"
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX Codigo ON LParam (Codigo) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "INSERT INTO LParam (Codigo, Valor) VALUES (1, NULL )"
      Rc = ExecSQL(DbMain, Q1, False)


      If lUpdOK Then
         DbVer = 2
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 2 -----------------------------------

   If DbVer = 2 And lUpdOK = True Then

      Q1 = "DROP TABLE IPC"
      Rc = ExecSQL(DbMain, Q1)

      Q1 = "CREATE TABLE IPC ( AnoMes long, pIPC float, vIPC float, fCM float)"
      Rc = ExecSQL(DbMain, Q1)

      Q1 = "CREATE UNIQUE INDEX AnoMes ON IPC (AnoMes) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1)

      If lUpdOK Then
         DbVer = 3
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 3 -----------------------------------

   If DbVer = 3 And lUpdOK = True Then ' 13 DIC 2004
   
      Q1 = "CREATE UNIQUE INDEX Perfil ON Perfiles (Nombre)"
      Call ExecSQL(DbMain, Q1)

      Q1 = "SELECT Max(idPerfil) as M FROM Perfiles"
      Set Rs = OpenRs(DbMain, Q1)
      id = vFld(Rs("M")) + 1
      Call CloseRs(Rs)

      Q1 = "INSERT INTO Perfiles (idPerfil, Nombre, Privilegios, idApp)"
      Q1 = Q1 & " VALUES (" & id & ",'(todo)', 65535, 0)"
      Rc = ExecSQL(DbMain, Q1)

      Q1 = "SELECT Max(idPerfil) as M FROM Perfiles"
      Set Rs = OpenRs(DbMain, Q1)
      id = vFld(Rs("M"))
      Call CloseRs(Rs)

      Q1 = "UPDATE UsuarioEmpresa INNER JOIN Usuarios ON UsuarioEmpresa.idUsuario = Usuarios.IdUsuario"
      Q1 = Q1 & " SET UsuarioEmpresa.idPerfil=" & id
      Q1 = Q1 & " WHERE Usuarios.Usuario IN ('Usuario1', 'Usuario2', 'Usuario3')"
      Rc = ExecSQL(DbMain, Q1)

      If lUpdOK Then
         DbVer = 4
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 4 -----------------------------------

   If DbVer = 4 And lUpdOK = True Then ' 21 Ene 2005
   
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
         lUpdOK = False
      
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
         lUpdOK = False
      End If

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29Neto", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29Neto", vbExclamation
         lUpdOK = False
      End If

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29IVA", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29IVA", vbExclamation
         lUpdOK = False
      End If

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29IVADTE", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29IVADTE", vbExclamation
         lUpdOK = False
      End If

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29AFCount", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29AFCount", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29AFIVA", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29AFIVA", vbExclamation
         lUpdOK = False
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
         lUpdOK = False
      End If
  
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29RetDieta", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29RetDieta", vbExclamation
         lUpdOK = False
      End If
   
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29IVARet3ro", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.CodF29IVARet3ro", vbExclamation
         lUpdOK = False
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

      If lUpdOK Then
         DbVer = 5
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If
   
      '--------------------- Versión 5 -----------------------------------

   If DbVer = 5 And lUpdOK = True Then ' 25 Ene 2005
   
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

      If lUpdOK Then
         DbVer = 6
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If
      
      '--------------------- Versión 6 -----------------------------------
      
   If DbVer = 6 And lUpdOK = True Then ' 31 Ene 2005
   
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
         lUpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("TieneExento", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocs.TieneExento", vbExclamation
         lUpdOK = False
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
         lUpdOK = False
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
         lUpdOK = False
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


      If lUpdOK Then
         DbVer = 7
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If
      
      '--------------------- Versión 7 -----------------------------------

   If DbVer = 7 And lUpdOK = True Then ' 9 Mar 2005
         
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
     
      If lUpdOK Then
         DbVer = 8
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If
   
   
      '--------------------- Versión 8 -----------------------------------

   If DbVer = 8 And lUpdOK = True Then ' 11 Mar 2005
   
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
         lUpdOK = False
      
      End If
                                 
      'ponemos en True este campo para los tipos de docs que corresponde
      Call ExecSQL(DbMain, "UPDATE TipoValor SET Multiple=-1 WHERE Valor = 'Afecto' OR Valor = 'Exento' OR Valor = 'Bruto' OR Valor = 'Honorarios sin Retención'")
      
      If lUpdOK Then
         DbVer = 9
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
    
      '--------------------- Versión 9 -----------------------------------

   If DbVer = 9 And lUpdOK = True Then ' 24 Mar 2005
   
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
      
      If lUpdOK Then
         DbVer = 10
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
      '--------------------- Versión 10 -----------------------------------

   If DbVer = 10 And lUpdOK = True Then ' 29 Mar 2005
   
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
         lUpdOK = False
      
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
         lUpdOK = False
      
      End If
      
      'seteamos el campo para los docs de boletas
      Q1 = "UPDATE TipoDocs SET DocBoletas=1 WHERE Diminutivo IN('BOV', 'DVB', 'BOE', 'VEM')"
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         DbVer = 11
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
      '--------------------- Versión 11 -----------------------------------

   If DbVer = 11 And lUpdOK = True Then ' 15 Abr 2005
   
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
      
      
      If lUpdOK Then
         DbVer = 12
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
  
   End If
      
    '--------------------- Versión 12 -----------------------------------
   
   If DbVer = 12 And lUpdOK = True Then   '28 Abril 2005
   
      Q1 = "UPDATE PlanAvanzado SET Atrib" & ATRIB_CAJA & "=1 WHERE " & GenLike(DbMain, "Caja", "Descripcion", GL_RWILD)
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanIntermedio SET Atrib" & ATRIB_CAJA & "=1 WHERE " & GenLike(DbMain, "Caja", "Descripcion", GL_RWILD)
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanBasico SET Atrib" & ATRIB_CAJA & "=1 WHERE " & GenLike(DbMain, "Caja", "Descripcion", GL_RWILD)
      Call ExecSQL(DbMain, Q1)
      
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         DbVer = 13
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
    '--------------------- Versión 13 -----------------------------------
   
   If DbVer = 13 And lUpdOK = True Then   '29 Abril 2005
   
      Q1 = "UPDATE PlanAvanzado SET TipoCapPropio=" & CAPPROPIO_PASIVO_NOEXIGIBLE & " WHERE Codigo IN('2030101','2030201', '2030301', '2030401','2030501','2031101','2031201','2031301')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanIntermedio SET TipoCapPropio=" & CAPPROPIO_PASIVO_NOEXIGIBLE & " WHERE Codigo IN('2030101','2030201', '2030301', '2030401','2030501','2031101','2031201','2031301')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE PlanBasico SET TipoCapPropio=" & CAPPROPIO_PASIVO_NOEXIGIBLE & " WHERE Codigo IN('2030101','2030201', '2030301', '2030401','2030501','2031101','2031201','2031301')"
      Call ExecSQL(DbMain, Q1)
      
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         DbVer = 14
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
      
    '--------------------- Versión 14 -----------------------------------
   
   If DbVer = 14 And lUpdOK = True Then   ' 13 mayo 2005
   
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
           
      If lUpdOK Then
         DbVer = 15
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
    '--------------------- Versión 15 -----------------------------------
    
   If DbVer = 15 And lUpdOK = True Then   ' 14 julio 2005
   
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
              
      If lUpdOK Then
         DbVer = 16
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   '--------------------- Versión 16 -----------------------------------
   
   If DbVer = 16 And lUpdOK = True Then   ' 14 julio 2005
   
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor ) "
      Q1 = Q1 & " VALUES ('VER', 3, '0' )" ' RUT
      Rc = ExecSQL(DbMain, Q1)
                            
      Q1 = "DROP INDEX PC_MAC ON Equipos"
      Rc = ExecSQL(DbMain, Q1)
                            
      Q1 = "CREATE UNIQUE INDEX PC_MAC_COD ON Equipos (PC, MAC, CodPC ) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1)

      If lUpdOK Then
         DbVer = 17
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   '--------------------- Versión 17 -----------------------------------
   
   If DbVer = 17 And lUpdOK = True Then   ' 21 julio 2005
      
      Call CrearTblContEmpresa
      
      If lUpdOK Then
         DbVer = 18
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
     
   '--------------------- Versión 18 -----------------------------------
   
   If DbVer = 18 And lUpdOK = True Then   ' 22 julio 2005
      
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
           
      If lUpdOK Then
         DbVer = 19
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
      '--------------------- Versión 19 -----------------------------------
   
   If DbVer = 19 And lUpdOK = True Then   ' 18 Agosto 2005

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

      If lUpdOK Then
         DbVer = 20
         Q1 = "UPDATE Param SET Valor=" & DbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   
   CorrigeBaseAdm_2005_01 = lUpdOK
      
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
   Set Fld = Tbl.CreateField("RazonFija", dbBoolean)
   Tbl.Fields.Append Fld

   If Err = 0 Then
      Tbl.Fields.Refresh
   
   ElseIf Err <> 3191 Then ' ya existe
      MsgBeep vbExclamation
      MsgBox "Error " & Err & ", " & Error & vbLf & "RazonesFin.RazonFija", vbExclamation
      CrearTblRazFin = False
   
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
   Rc = ExecSQL(DbMain, Q1, False)
   
   Q1 = "CREATE UNIQUE INDEX Nombre ON RazonesFin (Nombre)"
   Rc = ExecSQL(DbMain, Q1, False)
   
End Function
'Define tipos de razones financieras
Private Sub DefTiposRazFin()
   Dim Q1 As String
   
   On Error Resume Next
   
   'agregamos tipos de Razones Financieras
   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', " & RF_ENDEUDAMIENTO & ", 'Endeudamiento')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', " & RF_LIQUIDEZ & ", 'Liquidez')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', " & RF_RENTABILIDAD & ", 'Rentabilidad')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', " & RF_ROTACIONES & ", 'Rotaciones')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', " & RF_CONSOLIDACION & ", 'Consolidación')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', " & RF_OBSOLESCENCIA & ", 'Obsolescencia')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Param "
   Q1 = Q1 & " (Tipo, Codigo, Valor)"
   Q1 = Q1 & " VALUES( 'TIPORAZFIN', " & RF_OTROS & ", 'Otros')"
   Call ExecSQL(DbMain, Q1)

End Sub

Private Sub DefRazonesFin()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tipo As Integer
   
   On Error Resume Next
   
   '1.- Endeudamiento
   
   Tipo = RF_ENDEUDAMIENTO
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Deuda/Patrimonio E', 'Veces', 'Deuda Total', 'Patrimonio', '/', 1)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", '" & ParaSQL("Deuda/Patrimonio E'") & "', 'Veces', 'Deuda Total', 'Patrimonio', '/', 1)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Total Activos/Total Deuda', 'Veces', 'Total Activos', 'Total Deuda', '/', 1)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Cobertura de Intereses', 'Veces', 'Utilidad Operacional', 'Gastos Financieros', '/', 1)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Endeudamiento Corto Plazo', 'Veces', 'Deuda Corto Plazo', 'Patrimonio', '/', 1)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Endeudamiento Largo Plazo', 'Veces', 'Deuda Largo Plazo', 'Patrimonio', '/', 1)"
   Call ExecSQL(DbMain, Q1)

   '2.- Liquidez
   
   Tipo = RF_LIQUIDEZ
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Razón Circulante', 'Veces', 'Activos Circulantes', 'Pasivos Circulantes', '/', 1)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Razón Ácida', 'Veces', 'Activos Circulantes - Existencias', 'Pasivos Circulantes', '/', 1)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Razón de Tesorería', 'Veces', 'Efectivo + Efectivo Equivalente', 'Pasivos Circulantes', '/', 1)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Márgen de Maniobra', '$', 'Activos Circulantes - Pasivos Circulantes', ' ', '/', 1)"
   Call ExecSQL(DbMain, Q1)

   '3.- Rentabilidad
   
   Tipo = RF_RENTABILIDAD
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Utilidades sobre Venta', '%', 'Utilidades', 'Ventas', '/', 1)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Utilidades sobre Activos', '%', 'Utilidades', 'Activos', '/', 1)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Utilidades sobre Patrimonio', '%', 'Utilidades', 'Patrimonio', '/', 1)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'ROA', '%', 'Utilidades antes de Impuestos', 'Activo Total Neto', '/', 1)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'ROE', '%', 'Utilidad Líquida', 'Patrimonio', '/', 1)"
   Call ExecSQL(DbMain, Q1)
    
   '4.- Rotaciones
   
   Tipo = RF_ROTACIONES
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Rotación de Cuentas por Cobrar', 'Veces', 'Ventas Brutas', 'Total Cuentas por Cobrar', '/', 1)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Rotación de Cuentas por Pagar', 'Veces', 'Costo de Venta', 'Total Cuentas por Pagar', '/', 1)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Rotación de Inventario', 'Veces', 'Costo de Venta', 'Inventario', '/', 1)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Periodo Medio de Cobro', 'Días', 'Total Cuentas por Cobrar', 'Ventas Brutas/Cantidad de Días', '/', 1)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Periodo Medio de Pago', 'Días', 'Total Cuentas por Pagar', 'Costo de Ventas Bruto/Cantidad de Días', '/', 1)"
   Call ExecSQL(DbMain, Q1)
  
   '5.- Consolidación
   
   Tipo = RF_CONSOLIDACION
  
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Activos Consolidados/Activo Neto', '%', 'Valor Activos Consolidados', 'Total Activo Neto', '/', 1)"
   Call ExecSQL(DbMain, Q1)
  
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Activos Consolidados/Activo Neto - Activo Circ.', '%', 'Valor Activos Consolidados', 'Total Activo Neto - Activo Circulante', '/', 1)"
   Call ExecSQL(DbMain, Q1)
  
   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Activos Consolidados/Activo Neto - Activo Circ.', '%', 'Valor Activos Consolidados', 'Total Activo Neto - Activo Circulante', '/', 1)"
   Call ExecSQL(DbMain, Q1)
  
   '6.- Obsolescencia
   
   Tipo = RF_OBSOLESCENCIA

   Q1 = "INSERT INTO RazonesFin "
   Q1 = Q1 & "( Tipo, Nombre, UnidadRes, TxtNumerador, TxtDenominador, Operador, RazonFija  )"
   Q1 = Q1 & " VALUES (" & Tipo & ", 'Obsolescencia', '%', 'Depreciación Acumulada', 'Total Activo Fijo Bruto', '/', 1)"
   Call ExecSQL(DbMain, Q1)

End Sub
'define las cuentas que intervienen en cada razón financiera
Private Sub DefCuentasRazFin()
   On Error Resume Next


End Sub

Private Sub CorrigeRegComunas()
   Dim Q1 As String
   Dim Rs As Recordset
   
   'modificamos nombres con Ñ y otros
   Q1 = "UPDATE Regiones SET Comuna = 'CAMIÑA' WHERE Comuna = 'CAMI-A'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Comuna = 'CHAÑARAL' WHERE Comuna = 'CHA-ARAL'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Comuna = 'VICUÑA' WHERE Comuna = 'VICU-A'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Comuna = 'VIÑA DEL MAR' WHERE Comuna = 'VI-A DEL MAR'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Comuna = 'CON-CON' WHERE Comuna = 'CONCON'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE Regiones SET Comuna = 'CAÑETE' WHERE Comuna = 'CA-ETE'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE Regiones SET Comuna = 'RIO IBAÑEZ' WHERE Comuna = 'RIO IBA-EZ'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE Regiones SET Comuna = 'SAN JOSE DE MAIPO' WHERE Comuna = 'SAN JOSE MAIPO'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE Regiones SET Comuna = 'ÑIQUEN' WHERE Comuna = 'SAN GREGORIO DE ÑIQUEN'"
   Call ExecSQL(DbMain, Q1)

   'comunas descontinuadas
   Q1 = "UPDATE Regiones SET Comuna = 'SANTIAGO (*)' WHERE Comuna = 'SANTIAGO'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE Regiones SET Comuna = 'ANTARTICA (*)' WHERE Comuna = 'ANTARTICA'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE Regiones SET Comuna = 'NAVARINO (*)' WHERE Comuna = 'NAVARINO'"
   Call ExecSQL(DbMain, Q1)

   'Agregamos nuevas comunas
   Q1 = "INSERT INTO Regiones(Codigo, Comuna) VALUES('01', 'ALTO HOSPICIO')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO Regiones(Codigo, Comuna) VALUES('08', 'HUALPEN')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Regiones(Codigo, Comuna) VALUES('08', 'ALTO BIOBIO')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Regiones(Codigo, Comuna) VALUES('09', 'CHOLCHOL')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Regiones(Codigo, Comuna) VALUES(12, 'CABO DE HORNOS')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Regiones(Codigo, Comuna) VALUES(13, 'SANTIAGO CENTRO')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO Regiones(Codigo, Comuna) VALUES(13, 'SANTIAGO OESTE')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO Regiones(Codigo, Comuna) VALUES(13, 'SANTIAGO SUR')"
   Call ExecSQL(DbMain, Q1)
   
   'actualizamos regiones de algunas comunas
   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'VALDIVIA'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'MARIQUINA'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'LANCO'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'LOS LAGOS'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'FUTRONO'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'CORRAL'"
   Call ExecSQL(DbMain, Q1)
      
   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'MAFIL'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'PANGUIPULLI'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'LA UNION'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'PAILLACO'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'RIO BUENO'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 14 WHERE Comuna = 'LAGO RANCO'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 15 WHERE Comuna = 'ARICA'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE Regiones SET Codigo = 15 WHERE Comuna = 'CAMARONES'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 15 WHERE Comuna = 'PUTRE'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE Regiones SET Codigo = 15 WHERE Comuna = 'GENERAL LAGOS'"
   Call ExecSQL(DbMain, Q1)

End Sub

Private Sub CodActEconomicas()
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "UPDATE CodActiv SET Version=1 WHERE Codigo IN( '721000', '723000', '742120', '742130', '742140', '74300', '92200', '95001')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO CodActiv (Codigo, Descrip, Version) "
   Q1 = Q1 & "VALUES( '742121','EMPRESA DE SERVICIOS GEOLOGICOS Y DE PROSPECCION',2)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO CodActiv (Codigo, Descrip, Version) "
   Q1 = Q1 & "VALUES( '742122','SERVICIOS PROFESIONALES EN GEOLOGIA Y PROSPECCION',2)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO CodActiv (Codigo, Descrip, Version) "
   Q1 = Q1 & "VALUES( '742131','EMPRESA DE SERVICIOS DE TOPOGRAFIA Y AGRIMENSURA',2)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO CodActiv (Codigo, Descrip, Version) "
   Q1 = Q1 & "VALUES( '742132','SERVICIOS PROFESIONALES DE TOPOGRAFIA Y AGRIMENSURA',2)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO CodActiv (Codigo, Descrip, Version) "
   Q1 = Q1 & "VALUES( '742141','SERVICIOS DE INGENIERIA PRESTADOS POR EMPRESAS N.C.P.',2)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO CodActiv (Codigo, Descrip, Version) "
   Q1 = Q1 & "VALUES( '742142','SERVICIOS DE INGENIERIA PRESTADOS POR PROFESIONALES N.C.P.',2)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO CodActiv (Codigo, Descrip, Version) "
   Q1 = Q1 & "VALUES( '743001','EMPRESAS DE PUBLICIDAD',2)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO CodActiv (Codigo, Descrip, Version) "
   Q1 = Q1 & "VALUES( '743002','SERVICIOS PERSONALES EN PUBLICIDAD',2)"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO CodActiv (Codigo, Descrip, Version) "
   Q1 = Q1 & "VALUES( '922001','AGENCIAS DE NOTICIAS',2)"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO CodActiv (Codigo, Descrip, Version) "
   Q1 = Q1 & "VALUES( '922002','SERVICIOS PERIODISTICOS PRESTADO POR PROFESIONALES',2)"
   Call ExecSQL(DbMain, Q1)



   Q1 = "UPDATE CodActiv SET Descrip='RECICLAMIENTO DE OTROS DESPERDICIOS Y DESECHOS N.C.P.' "
   Q1 = Q1 & " WHERE Codigo = '372090'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='ALMACENES MEDIANOS (VENTA DE ALIMENTOS); SUPERMERCADOS, MINIMARKETS' "
   Q1 = Q1 & " WHERE Codigo = '521112'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE CodActiv SET Descrip='ALMACENES PEQUENOS (VENTA DE ALIMENTOS)' "
   Q1 = Q1 & " WHERE Codigo = '521120'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='GRANDES TIENDAS - PRODUCTOS DE FERRETERIA Y PARA EL HOGAR' "
   Q1 = Q1 & " WHERE Codigo = '521200'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='GRANDES TIENDAS -  VESTUARIO Y PRODUCTOS PARA EL HOGAR' "
   Q1 = Q1 & " WHERE Codigo = '521300'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='AGENTES Y LIQUIDADORES DE SEGUROS' "
   Q1 = Q1 & " WHERE Codigo = '672020'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='CORREDORES DE PROPIEDADES' "
   Q1 = Q1 & " WHERE Codigo = '702000'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='ASESORES Y CONSULTORES EN INFORMATICA (SOFTWARE)' "
   Q1 = Q1 & " WHERE Codigo = '722000'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='PROCESAMIENTO DE DATOS Y ACTIVIDADES RELACIONADAS CON BASES DE DATOS' "
   Q1 = Q1 & " WHERE Codigo = '724000'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='EMPRESA DE SERVICIOS INTEGRALES DE INFORMATICA' "
   Q1 = Q1 & " WHERE Codigo = '726000'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='OTROS SERVICIOS DESARROLLADOS POR PROFESIONALES' "
   Q1 = Q1 & " WHERE Codigo = '742190'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE CodActiv SET Descrip='SERVICIOS SUMINISTRO DE PERSONAL; EMPRESAS SERVICIOS TRANSITORIOS' "
   Q1 = Q1 & " WHERE Codigo = '749110'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='SERVICIOS PERSONALES RELACIONADOS CON SEGURIDAD' "
   Q1 = Q1 & " WHERE Codigo = '749229'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='EMPRESAS DE LIMPIEZA DE EDIFICIOS RESIDENCIALES Y NO RESIDENCIALES' "
   Q1 = Q1 & " WHERE Codigo = '749310'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='EMPRESAS DE LIMPIEZA DE EDIFICIOS RESIDENCIALES Y NO RESIDENCIALES' "
   Q1 = Q1 & " WHERE Codigo = '749310'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='OTRAS ACTIVIDADES DE FOTOGRAFIA' "
   Q1 = Q1 & " WHERE Codigo = '749409'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='SERVICIOS PERSONALES DE TRADUCCION, INTERPRETACION Y LABORES DE OFICINA' "
   Q1 = Q1 & " WHERE Codigo = '749932'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='EMPRESA DE TRADUCCION E INTERPRETACION' "
   Q1 = Q1 & " WHERE Codigo = '749933'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='ESTABLECIMIENTO ENSENANZA PREESCOLAR' "
   Q1 = Q1 & " WHERE Codigo = '801010'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE CodActiv SET Descrip='ESTABLECIENTO ENSENANZA PRIMARIA' "
   Q1 = Q1 & " WHERE Codigo = '801020'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='ESTABLECIMIENTO ENSENANZA SECUNDARIA DE FORMACION GENERAL' "
   Q1 = Q1 & " WHERE Codigo = '802100'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='ESTABLECIMIENTO ENSENANZA SECUNDARIA DE FORMACION TECNICA Y PROFESIONAL' "
   Q1 = Q1 & " WHERE Codigo = '802200'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='UNIVERSIDADES' "
   Q1 = Q1 & " WHERE Codigo = '803010'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='INSTITUTOS PROFESIONALES' "
   Q1 = Q1 & " WHERE Codigo = '803020'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='CENTROS DE FORMACION TECNICA' "
   Q1 = Q1 & " WHERE Codigo = '803030'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='ESTABLECIMIENTO ENSENANZA PRIMARIA Y SECUNDARIA PARA ADULTOS' "
   Q1 = Q1 & " WHERE Codigo = '809010'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='ESTABLECIEMIENTO ENSENANZA PREUNIVERSITARIA' "
   Q1 = Q1 & " WHERE Codigo = '809020'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='EDUCACION A DISTANCIA (INTERNET, CORRESPONDENCIA, OTRAS)' "
   Q1 = Q1 & " WHERE Codigo = '809041'"
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE CodActiv SET Descrip='SERVICIOS DE AGENCIAS DE NOTICIAS' "
   Q1 = Q1 & " WHERE Codigo = '922000'"
   Call ExecSQL(DbMain, Q1)

End Sub


Private Sub UpdateIPCUTM2009()
   Dim i As Integer, a As Integer, Ano As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim PtoIPC(9, 12) As Double
   Dim VarIpc(9, 12) As Double
   Dim IPCAcum(9, 12) As Double
   Dim Factor(9, 12) As Double
   Dim UTM(9, 12) As Double
   Dim IdMoneda As Long
 
   PtoIPC(5, 12) = 84.434
   
   PtoIPC(6, 1) = 84.503
   PtoIPC(6, 2) = 84.427
   PtoIPC(6, 3) = 84.922
   PtoIPC(6, 4) = 85.465
   PtoIPC(6, 5) = 85.674
   PtoIPC(6, 6) = 86.176
   PtoIPC(6, 7) = 86.643
   PtoIPC(6, 8) = 86.873
   PtoIPC(6, 9) = 86.887
   PtoIPC(6, 10) = 86.664
   PtoIPC(6, 11) = 86.518
   PtoIPC(6, 12) = 86.602
      
   PtoIPC(7, 1) = 86.867
   PtoIPC(7, 2) = 86.72
   PtoIPC(7, 3) = 87.09
   PtoIPC(7, 4) = 87.591
   PtoIPC(7, 5) = 88.135
   PtoIPC(7, 6) = 88.958
   PtoIPC(7, 7) = 89.962
   PtoIPC(7, 8) = 90.938
   PtoIPC(7, 9) = 91.969
   PtoIPC(7, 10) = 92.255
   PtoIPC(7, 11) = 92.952
   PtoIPC(7, 12) = 93.377

   PtoIPC(8, 1) = 93.343
   PtoIPC(8, 2) = 93.719
   PtoIPC(8, 3) = 94.5
   PtoIPC(8, 4) = 94.862
   PtoIPC(8, 5) = 95.957
   PtoIPC(8, 6) = 97.386
   PtoIPC(8, 7) = 98.487
   PtoIPC(8, 8) = 99.4
   PtoIPC(8, 9) = 100.46
   PtoIPC(8, 10) = 101.345
   PtoIPC(8, 11) = 101.213
   PtoIPC(8, 12) = 100

   PtoIPC(9, 1) = 99.24
   PtoIPC(9, 2) = 98.88
   PtoIPC(9, 3) = 99.26
   PtoIPC(9, 4) = 99.11

   VarIpc(6, 1) = 0.001
   VarIpc(6, 2) = -0.001
   VarIpc(6, 3) = 0.006
   VarIpc(6, 4) = 0.006
   VarIpc(6, 5) = 0.002
   VarIpc(6, 6) = 0.006
   VarIpc(6, 7) = 0.005
   VarIpc(6, 8) = 0.003
   VarIpc(6, 9) = 0
   VarIpc(6, 10) = -0.003
   VarIpc(6, 11) = -0.002
   VarIpc(6, 12) = 0.001
   
   VarIpc(7, 1) = 0.003
   VarIpc(7, 2) = -0.002
   VarIpc(7, 3) = 0.004
   VarIpc(7, 4) = 0.006
   VarIpc(7, 5) = 0.006
   VarIpc(7, 6) = 0.009
   VarIpc(7, 7) = 0.011
   VarIpc(7, 8) = 0.011
   VarIpc(7, 9) = 0.011
   VarIpc(7, 10) = 0.003
   VarIpc(7, 11) = 0.008
   VarIpc(7, 12) = 0.005

   VarIpc(8, 1) = 0
   VarIpc(8, 2) = 0.004
   VarIpc(8, 3) = 0.008
   VarIpc(8, 4) = 0.004
   VarIpc(8, 5) = 0.012
   VarIpc(8, 6) = 0.015
   VarIpc(8, 7) = 0.011
   VarIpc(8, 8) = 0.009
   VarIpc(8, 9) = 0.011
   VarIpc(8, 10) = 0.009
   VarIpc(8, 11) = -0.001
   VarIpc(8, 12) = -0.012

   VarIpc(9, 1) = -0.008
   VarIpc(9, 2) = -0.004
   VarIpc(9, 3) = 0.004

   IPCAcum(7, 1) = 0.3
   IPCAcum(7, 2) = 0.1
   IPCAcum(7, 3) = 0.6
   IPCAcum(7, 4) = 1.1
   IPCAcum(7, 5) = 1.8
   IPCAcum(7, 6) = 2.7
   IPCAcum(7, 7) = 3.9
   IPCAcum(7, 8) = 5
   IPCAcum(7, 9) = 6.2
   IPCAcum(7, 10) = 6.5
   IPCAcum(7, 11) = 7.3
   IPCAcum(7, 12) = 7.8

   IPCAcum(8, 1) = 0
   IPCAcum(8, 2) = 0.4
   IPCAcum(8, 3) = 1.2
   IPCAcum(8, 4) = 1.6
   IPCAcum(8, 5) = 2.8
   IPCAcum(8, 6) = 4.3
   IPCAcum(8, 7) = 5.5
   IPCAcum(8, 8) = 6.5
   IPCAcum(8, 9) = 7.6
   IPCAcum(8, 10) = 8.5
   IPCAcum(8, 11) = 8.4
   IPCAcum(8, 12) = 7.1
   
   IPCAcum(9, 1) = -0.8
   IPCAcum(9, 2) = -1.1
   IPCAcum(9, 3) = -0.7

   Factor(7, 1) = 1.073
   Factor(7, 2) = 1.07
   Factor(7, 3) = 1.072
   Factor(7, 4) = 1.067
   Factor(7, 5) = 1.061
   Factor(7, 6) = 1.055
   Factor(7, 7) = 1.045
   Factor(7, 8) = 1.033
   Factor(7, 9) = 1.022
   Factor(7, 10) = 1.011
   Factor(7, 11) = 1.008
   Factor(7, 12) = 1

   Factor(8, 1) = 1.084
   Factor(8, 2) = 1.084
   Factor(8, 3) = 1.08
   Factor(8, 4) = 1.071
   Factor(8, 5) = 1.067
   Factor(8, 6) = 1.055
   Factor(8, 7) = 1.039
   Factor(8, 8) = 1.028
   Factor(8, 9) = 1.018
   Factor(8, 10) = 1.007
   Factor(8, 11) = 1
   Factor(8, 12) = 1
   
   UTM(6, 1) = 31508
   UTM(6, 2) = 31413
   UTM(6, 3) = 31444
   UTM(6, 4) = 31413
   UTM(6, 5) = 31601
   UTM(6, 6) = 31791
   UTM(6, 7) = 31855
   UTM(6, 8) = 32046
   UTM(6, 9) = 32206
   UTM(6, 10) = 32303
   UTM(6, 11) = 32303
   UTM(6, 12) = 32206

   UTM(7, 1) = 32142
   UTM(7, 2) = 32174
   UTM(7, 3) = 32271
   UTM(7, 4) = 32206
   UTM(7, 5) = 32335
   UTM(7, 6) = 32529
   UTM(7, 7) = 32724
   UTM(7, 8) = 33019
   UTM(7, 9) = 33382
   UTM(7, 10) = 33749
   UTM(7, 11) = 34120
   UTM(7, 12) = 34222

   UTM(8, 1) = 34496
   UTM(8, 2) = 34668
   UTM(8, 3) = 34668
   UTM(8, 4) = 34807
   UTM(8, 5) = 35085
   UTM(8, 6) = 35225
   UTM(8, 7) = 35648
   UTM(8, 8) = 36183
   UTM(8, 9) = 36581
   UTM(8, 10) = 36910
   UTM(8, 11) = 37316
   UTM(8, 12) = 37652
   
   UTM(9, 1) = 37614
   UTM(9, 2) = 37163
   UTM(9, 3) = 36866
   UTM(9, 4) = 36719
   UTM(9, 5) = 36866

   'primero el IPC
   
   For a = 5 To 9
   
      For i = 1 To 12
      
         If a = 9 And i > 3 Then
            Exit For
         End If
      
         If PtoIPC(a, i) <> 0 Then
            Ano = 2000 + a
            
            Q1 = "SELECT pIPC FROM IPC WHERE AnoMes = " & CLng(DateSerial(Ano, i, 1))
            Set Rs = OpenRs(DbMain, Q1)
            
            If Rs.EOF Then
               Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
               Q1 = Q1 & " VALUES(" & CLng(DateSerial(Ano, i, 1)) & "," & str(PtoIPC(a, i)) & "," & str(VarIpc(a, i)) & "," & str(Factor(a, i)) & ")"
   
            Else
               Q1 = "UPDATE IPC SET "
               Q1 = Q1 & "  pIPC = " & str(PtoIPC(a, i))
               Q1 = Q1 & ", vIPC = " & str(VarIpc(a, i))
               Q1 = Q1 & ", fCM = " & str(Factor(a, i))
               Q1 = Q1 & " WHERE AnoMes = " & CLng(DateSerial(Ano, i, 1))
            End If
            
            Call CloseRs(Rs)
            
            Call ExecSQL(DbMain, Q1)
         End If
         
      Next i
      
   Next a
   
   'ahora la UTM
   
   Q1 = "SELECT IdMoneda FROM Monedas WHERE Simbolo = 'UTM'"
   Set Rs = OpenRs(DbMain, Q1)
   
   IdMoneda = 0
   If Rs.EOF = False Then
      IdMoneda = vFld(Rs("IdMoneda"))
   End If
   
   Call CloseRs(Rs)
   
   If IdMoneda > 0 Then
   
      For a = 6 To 9
      
         For i = 1 To 12
         
            If a = 9 And i > 5 Then
               Exit For
            End If
         
            Ano = 2000 + a
            
            Q1 = "SELECT Valor FROM Equivalencia WHERE IdMoneda = " & IdMoneda & " AND Fecha = " & CLng(DateSerial(Ano, i, 1))
            Set Rs = OpenRs(DbMain, Q1)
            
            If Rs.EOF Then
               Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
               Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(Ano, i, 1)) & "," & str(UTM(a, i)) & ")"
   
            Else
               Q1 = "UPDATE Equivalencia SET "
               Q1 = Q1 & "  Valor = " & str(UTM(a, i))
               Q1 = Q1 & " WHERE IdMoneda = " & IdMoneda & " AND Fecha = " & CLng(DateSerial(Ano, i, 1))
            End If
            
            Call CloseRs(Rs)
            
            Call ExecSQL(DbMain, Q1)
            
         Next i
         
      Next a
   End If
   
End Sub



Private Sub UpdateIPCUTM2009_10_11()
   Dim i As Integer, a As Integer, Ano As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim PtoIPC(12, 12) As Double
   Dim VarIpc(12, 12) As Double
   Dim IPCAcum(12, 12) As Double
   Dim Factor(12, 12) As Double
   Dim UTM(12, 12) As Double
   Dim IdMoneda As Long
    
   PtoIPC(9, 8) = 98.41
   PtoIPC(9, 9) = 99.38
   PtoIPC(9, 10) = 99.38
   PtoIPC(9, 11) = 98.92
   PtoIPC(9, 12) = 98.62
      
   PtoIPC(10, 1) = 100.03
   PtoIPC(10, 2) = 100.31
   PtoIPC(10, 3) = 100.39
   PtoIPC(10, 4) = 100.86
   PtoIPC(10, 5) = 101.22
   PtoIPC(10, 6) = 101.22
   PtoIPC(10, 7) = 101.87
   PtoIPC(10, 8) = 101.77
   PtoIPC(10, 9) = 102.18
   PtoIPC(10, 10) = 102.28
   PtoIPC(10, 11) = 102.35
   PtoIPC(10, 12) = 102.47
      
   PtoIPC(11, 1) = 102.76
   PtoIPC(11, 2) = 102.98
   PtoIPC(11, 3) = 103.77
   PtoIPC(11, 4) = 104.1
   PtoIPC(11, 5) = 104.52
   PtoIPC(11, 6) = 104.7
   PtoIPC(11, 7) = 104.83
   PtoIPC(11, 8) = 105
   PtoIPC(11, 9) = 105.52
   PtoIPC(11, 10) = 106.03
'   PtoIPC(11, 11) = 0
'   PtoIPC(11, 12) = 0
      

   VarIpc(9, 1) = -0.008
   VarIpc(9, 2) = -0.004
   VarIpc(9, 3) = 0.004
   VarIpc(9, 4) = -0.002
   VarIpc(9, 5) = -0.003
   VarIpc(9, 6) = 0.003
   VarIpc(9, 7) = -0.004
   VarIpc(9, 8) = -0.004
   VarIpc(9, 9) = 1
   VarIpc(9, 10) = 0
   VarIpc(9, 11) = -0.005
   VarIpc(9, 12) = -0.003
   
   VarIpc(10, 1) = 0.005
   VarIpc(10, 2) = 0.003
   VarIpc(10, 3) = 0.001
   VarIpc(10, 4) = 0.005
   VarIpc(10, 5) = 0.004
   VarIpc(10, 6) = 0
   VarIpc(10, 7) = 0.006
   VarIpc(10, 8) = -0.001
   VarIpc(10, 9) = 0.004
   VarIpc(10, 10) = 0.001
   VarIpc(10, 11) = 0.001
   VarIpc(10, 12) = 0.001
   
   VarIpc(11, 1) = 0.003
   VarIpc(11, 2) = 0.002
   VarIpc(11, 3) = 0.008
   VarIpc(11, 4) = 0.003
   VarIpc(11, 5) = 0.004
   VarIpc(11, 6) = 0.002
   VarIpc(11, 7) = 0.001
   VarIpc(11, 8) = 0.002
   VarIpc(11, 9) = 0.005
   VarIpc(11, 10) = 0.005
'   VarIPC(11, 11) = 0
'   VarIPC(11, 12) = 0
   

   IPCAcum(9, 8) = -0.4
   IPCAcum(9, 9) = 1
   IPCAcum(9, 10) = 0
   IPCAcum(9, 11) = -0.5
   IPCAcum(9, 12) = -0.3

   IPCAcum(10, 1) = 0.5
   IPCAcum(10, 2) = 0.8
   IPCAcum(10, 3) = 0.9
   IPCAcum(10, 4) = 1.4
   IPCAcum(10, 5) = 1.7
   IPCAcum(10, 6) = 1.7
   IPCAcum(10, 7) = 2.4
   IPCAcum(10, 8) = 2.3
   IPCAcum(10, 9) = 2.7
   IPCAcum(10, 10) = 2.8
   IPCAcum(10, 11) = 2.9
   IPCAcum(10, 12) = 3#
   
   IPCAcum(11, 1) = 0.3
   IPCAcum(11, 2) = 0.2
   IPCAcum(11, 3) = 0.8
   IPCAcum(11, 4) = 0.3
   IPCAcum(11, 5) = 0.4
   IPCAcum(11, 6) = 0.2
   IPCAcum(11, 7) = 0.1
   IPCAcum(11, 8) = 0.2
   IPCAcum(11, 9) = 0.5
   IPCAcum(11, 10) = 0.5
'   IPCAcum(11, 11) = 0
'   IPCAcum(11, 12) = 0

   Factor(9, 1) = -1.011
   Factor(9, 2) = -1.003
   Factor(9, 3) = 0
   Factor(9, 4) = -1.003
   Factor(9, 5) = -1.002
   Factor(9, 6) = 1.001
   Factor(9, 7) = -1.003
   Factor(9, 8) = 1.002
   Factor(9, 9) = 1.005
   Factor(9, 10) = -1.005
   Factor(9, 11) = -1.005
   Factor(9, 12) = 1
   
   Factor(10, 1) = 1.029
   Factor(10, 2) = 1.023
   Factor(10, 3) = 1.02
   Factor(10, 4) = 1.02
   Factor(10, 5) = 1.015
   Factor(10, 6) = 1.011
   Factor(10, 7) = 1.011
   Factor(10, 8) = 1.005
   Factor(10, 9) = 1.006
   Factor(10, 10) = 1.002
   Factor(10, 11) = 1.001
   Factor(10, 12) = 1
   
   Factor(11, 1) = 1.035
   Factor(11, 2) = 1.032
   Factor(11, 3) = 1.03
   Factor(11, 4) = 1.022
   Factor(11, 5) = 1.019
   Factor(11, 6) = 1.014
   Factor(11, 7) = 1.013
   Factor(11, 8) = 1.011
   Factor(11, 9) = 1.01
   Factor(11, 10) = 1.005
   Factor(11, 11) = 1
   Factor(11, 12) = 1
   
   UTM(9, 1) = 37614
   UTM(9, 2) = 37163
   UTM(9, 3) = 36866
   UTM(9, 4) = 36719
   UTM(9, 5) = 36866
   UTM(9, 6) = 36792
   UTM(9, 7) = 36682
   UTM(9, 8) = 36792
   UTM(9, 9) = 36645
   UTM(9, 10) = 36498
   UTM(9, 11) = 36863
   UTM(9, 12) = 36863

   UTM(10, 1) = 36679
   UTM(10, 2) = 36569
   UTM(10, 3) = 36752
   UTM(10, 4) = 36862
   UTM(10, 5) = 36899
   UTM(10, 6) = 37083
   UTM(10, 7) = 37231
   UTM(10, 8) = 37231
   UTM(10, 9) = 37454
   UTM(10, 10) = 37417
   UTM(10, 11) = 37567
   UTM(10, 12) = 37605

   UTM(11, 1) = 37643
   UTM(11, 2) = 37681
   UTM(11, 3) = 37794
   UTM(11, 4) = 37870
   UTM(11, 5) = 38173
   UTM(11, 6) = 38288
   UTM(11, 7) = 38441
   UTM(11, 8) = 38518
   UTM(11, 9) = 38557
   UTM(11, 10) = 38634
   UTM(11, 11) = 38827
   UTM(11, 12) = 39021
   

   'primero el IPC
   
   For a = 9 To 11
   
      For i = 1 To 12
      
         If (a = 9 And i >= 8) Or a > 9 Then
      
            If PtoIPC(a, i) <> 0 Then
               Ano = 2000 + a
               
               Q1 = "SELECT pIPC FROM IPC WHERE AnoMes = " & CLng(DateSerial(Ano, i, 1))
               Set Rs = OpenRs(DbMain, Q1)
               
               If Rs.EOF Then
                  Q1 = "INSERT INTO IPC (AnoMes, pIPC, vIPC, fCM)"
                  Q1 = Q1 & " VALUES(" & CLng(DateSerial(Ano, i, 1)) & "," & str(PtoIPC(a, i)) & "," & str(VarIpc(a, i)) & "," & str(Factor(a, i)) & ")"
               Else
                  Q1 = "UPDATE IPC SET "
                  Q1 = Q1 & "  pIPC = " & str(PtoIPC(a, i))
                  Q1 = Q1 & ", vIPC = " & str(VarIpc(a, i))
                  Q1 = Q1 & ", fCM = " & str(Factor(a, i))
                  Q1 = Q1 & " WHERE AnoMes = " & CLng(DateSerial(Ano, i, 1))
               End If
               
               Call CloseRs(Rs)
               
               Call ExecSQL(DbMain, Q1)
            End If
         End If
      Next i
      
   Next a
   
   'ahora la UTM
   
   Q1 = "SELECT IdMoneda FROM Monedas WHERE Simbolo = 'UTM'"
   Set Rs = OpenRs(DbMain, Q1)
   
   IdMoneda = 0
   If Rs.EOF = False Then
      IdMoneda = vFld(Rs("IdMoneda"))
   End If
   
   Call CloseRs(Rs)
   
   If IdMoneda > 0 Then
   
      For a = 9 To 11
      
         For i = 1 To 12
         
            If (a = 9 And i >= 8) Or a > 9 Then
            
               Ano = 2000 + a
               
               Q1 = "SELECT Valor FROM Equivalencia WHERE IdMoneda = " & IdMoneda & " AND Fecha = " & CLng(DateSerial(Ano, i, 1))
               Set Rs = OpenRs(DbMain, Q1)
               
               If Rs.EOF Then
                  Q1 = "INSERT INTO Equivalencia (IdMoneda, Fecha, Valor)"
                  Q1 = Q1 & " VALUES(" & IdMoneda & "," & CLng(DateSerial(Ano, i, 1)) & "," & str(UTM(a, i)) & ")"
      
               Else
                  Q1 = "UPDATE Equivalencia SET "
                  Q1 = Q1 & "  Valor = " & str(UTM(a, i))
                  Q1 = Q1 & " WHERE IdMoneda = " & IdMoneda & " AND Fecha = " & CLng(DateSerial(Ano, i, 1))
               End If
               
               Call CloseRs(Rs)
               
               Call ExecSQL(DbMain, Q1)
               
            End If
            
         Next i
            
      Next a
   End If
   
End Sub

Private Sub UpdCodDocSII()
   Dim Q1 As String

   'asignamos los códigos de tipo de documento de acuerdo al documento del SII
   
   'Ventas
   Q1 = "UPDATE TipoDocs SET CodDocSII = '030', CodDocDTESII='033' WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FAV'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '032', CodDocDTESII='034' WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FVE'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '035', CodDocDTESII='039' WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'BOV'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '038', CodDocDTESII='041' WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'BOE'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '040', CodDocDTESII='043' WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'LFV'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '045', CodDocDTESII='046' WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'FCV'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '055', CodDocDTESII='056' WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'NDV'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '060', CodDocDTESII='061' WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'NCV'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '101', CodDocDTESII='000' WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'EXP'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '104', CodDocDTESII='111' WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'NDE'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '106', CodDocDTESII='112' WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'NCE'"
   Call ExecSQL(DbMain, Q1)
      
   'Compras
   Q1 = "UPDATE TipoDocs SET CodDocSII = '030', CodDocDTESII='033' WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'FAC'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '032', CodDocDTESII='034' WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'FCE'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '040', CodDocDTESII='043' WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'LFC'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '045', CodDocDTESII='046' WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'FCC'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '055', CodDocDTESII='056' WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'NDC'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoDocs SET CodDocSII = '060', CodDocDTESII='061' WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'NCC'"
   Call ExecSQL(DbMain, Q1)
      
   Q1 = "UPDATE TipoDocs SET CodDocSII = '914', CodDocDTESII='000' WHERE TipoLib = " & LIB_COMPRAS & " AND Diminutivo = 'IMP'"
   Call ExecSQL(DbMain, Q1)
      
   
   'asignamos los códigos de tipo impuesto de acuerdo al documento del SII
   
   'Ventas
   Q1 = "UPDATE TipoValor SET CodImpSII = '39' WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'IVA Retenido Total'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '42' WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'IVA Retenido Parcial'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '126' WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'Rebaja 65%'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '127' WHERE TipoLib = " & LIB_VENTAS & " AND Valor = 'Otros Impuestos'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '113' WHERE TipoLib = " & LIB_VENTAS & " AND " & GenLike(DbMain, "Impto. Adicional", "Valor")
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '148' WHERE TipoLib = " & LIB_VENTAS & " AND " & GenLike(DbMain, "Licores", "Valor")
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '32' WHERE TipoLib = " & LIB_VENTAS & " AND " & GenLike(DbMain, "Vinos", "Valor")
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '150' WHERE TipoLib = " & LIB_VENTAS & " AND " & GenLike(DbMain, "Cervezas", "Valor")
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '146' WHERE TipoLib = " & LIB_VENTAS & " AND " & GenLike(DbMain, "Analcohólicas", "Valor")
   Call ExecSQL(DbMain, Q1)
   
      'por ahora que no nos han dado el código para estos impuestos, asimilándolo a Otros Impuestos
      Q1 = "UPDATE TipoValor SET CodImpSII = '127' WHERE TipoLib = " & LIB_VENTAS & " AND " & GenLike(DbMain, "ILA", "Valor")
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET CodImpSII = '127' WHERE TipoLib = " & LIB_VENTAS & " AND " & GenLike(DbMain, "Ret. anticipo", "Valor")
      Call ExecSQL(DbMain, Q1)
   


   'Compras
   Q1 = "UPDATE TipoValor SET CodImpSII = '14' WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA Irrecuperable'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '39' WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA Retenido Total'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '42' WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA Retenido Parcial'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '127' WHERE TipoLib = " & LIB_COMPRAS & " AND Valor IN( 'Imp. Esp. al Petróleo Diesel', 'Imp. Esp. al Petr. Diesel General', 'Imp. Esp. al Petr. Transp. Carga' )"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '28' WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'Otros Impuestos'"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '147' WHERE TipoLib = " & LIB_COMPRAS & " AND " & GenLike(DbMain, "Licores", "Valor")
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '33' WHERE TipoLib = " & LIB_COMPRAS & " AND " & GenLike(DbMain, "Vinos", "Valor")
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '149' WHERE TipoLib = " & LIB_COMPRAS & " AND " & GenLike(DbMain, "Cervezas", "Valor")
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '85' WHERE TipoLib = " & LIB_COMPRAS & " AND " & GenLike(DbMain, "Analcohólicas", "Valor")
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '160' WHERE TipoLib = " & LIB_COMPRAS & " AND " & GenLike(DbMain, "Harina Carne", "Valor", GL_OR + GL_WILD)
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodImpSII = '180' WHERE TipoLib = " & LIB_COMPRAS & " AND Valor = 'IVA Activo Fijo'"
   Call ExecSQL(DbMain, Q1)
   
      'por ahora que no nos han dado el código para estos impuestos, asimilándolo a Otros Impuestos
      Q1 = "UPDATE TipoValor SET CodImpSII = '28' WHERE TipoLib = " & LIB_COMPRAS & " AND " & GenLike(DbMain, "ILA", "Valor")
      Call ExecSQL(DbMain, Q1)
   
   
End Sub

Private Sub ActImpAdicionales2016()
   Dim Q1 As String

   'Compras
   
   'IVA Irrecuperable
   
   'dejamos sin código SII al iva Irrecuperable sin calificación
   Q1 = "UPDATE TipoValor SET CodSIIDTE = '', Valor = Valor & '(*)', Tit2 = Tit2 & '(*)', TitCompleto = 'IVA Irrecuperable (*)' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAIRREC
   Call ExecSQL(DbMain, Q1)

   'agregamos los distintos IVA irrecuperables
   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAIRREC1 & ", 'IVA Irrec.1: Compras dest. oper. exentas', ' ', ' ', 0, ' ', ',1,3,4,6,', 'IVA Irrec.1','Dest. oper. exentas',' ', 6, 0, 0, '1', 'IVA Irrec. 1: Compras destinadas a generar oper. no gravadas o exentas')"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAIRREC2 & ", 'IVA Irrec.2: Fact. proveedores fuera de plazo', ' ', ' ', 0, ' ', ',1,3,4,6,', 'IVA Irrec.2','Fact. prov. fuera de plazo',' ', 6, 0, 0, '2', 'IVA Irrec. 2: Facturas de proveedores regist. fuera de plazo')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAIRREC3 & ", 'IVA Irrec.3: Gastos rechazados', ' ', ' ', 0, ' ', ',1,3,4,6,', 'IVA Irrec.3','Gastos rechazados',' ', 6, 0, 0, '3', 'IVA Irrec. 3: Gastos rechazados')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAIRREC4 & ", 'IVA Irrec.4: Entregas gratuitas', ' ', ' ', 0, ' ', ',1,3,4,6,', 'IVA Irrec.4','Entregas gratuitas',' ', 6, 0, 0, '4', 'IVA Irrec. 4: Entregas gratuitas(premios, bonif., etc.) recibidas')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAIRREC9 & ", 'IVA Irrec.9: Otros', ' ', ' ', 0, ' ', ',1,3,4,6,', 'IVA Irrec.9','Otros',' ', 6, 0, 0, '9', 'IVA Irrec. 9: Otros')"
   Call ExecSQL(DbMain, Q1)

   'IVA Retenido Parcial
   
   'dejamos sin código SII al IVA Ret Parcial sin calificación
   Q1 = "UPDATE TipoValor SET Atributo = 'SINUSO', CodSIIDTE = '', Valor = 'IVA Ret. Parcial (*)', Tit1 = 'IVA Ret.', Tit2 = 'Parcial (*)', TitCompleto = 'IVA Retenido Parcial (*)' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVARETPARC
   Call ExecSQL(DbMain, Q1)

   'agregamos los distintos IVA Ret Parcial
   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETPARCTRIGO & ", 'IVA Ret. Parcial Trigo', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Parcial','Trigo',' ', 8, 0, -1, '34', 'IVA Retenido Parcial Trigo')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETPARCMADERA & ", 'IVA Ret. Parcial Madera', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Parcial','Madera',' ', 8, 0, -1, '33', 'IVA Retenido Parcial Madera')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETPARCGANADO & ", 'IVA Ret. Parcial Ganado', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Parcial','Ganado',' ', 8, 0, -1, '32', 'IVA Retenido Parcial Ganado')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETPARCLEGUMBRES & ", 'IVA Ret. Parcial Legumbres', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Parcial','Legumbres',' ', 8, 0, -1, '30', 'IVA Retenido Parcial Legumbres')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETPARCARROZ & ", 'IVA Ret. Parcial Arroz', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Parcial','Arroz',' ', 8, 0, -1, '36', 'IVA Retenido Parcial Arroz')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETPARCSILVESTRES & ", 'IVA Ret. Parcial Silvestres', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Parcial','Silvestres',' ', 8, 0, -1, '31', 'IVA Retenido Parcial Silvestres')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETPARCHIDROBIO & ", 'IVA Ret. Parcial Hidrobiol.', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Parcial','Hidrobiol.',' ', 8, 0, -1, '37', 'IVA Retenido Parcial Hidrobiológicas')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETPARCFAMBPASAS & ", 'IVA Ret. Parcial Framb. Pasas', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Parcial','Framb. Pasas',' ', 8, 14, -1, '48', 'IVA Retenido Parcial Frambuezas y Pasas')"
   Call ExecSQL(DbMain, Q1)

   
   'IVA Retenido Total
   
   'asignamos código SII al IVA Ret Total sin calificación
   Q1 = "UPDATE TipoValor SET CodSIIDTE = '15', Tasa = 100, EsRecuperable = -1, TitCompleto = 'IVA Retenido Total' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVARETTOT
   Call ExecSQL(DbMain, Q1)

   'agregamos los distintos IVA Ret Total
   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETTOTCHATARRA & ", 'IVA Ret. Total Chatarra', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Total','Chatarra',' ', 9, 100, -1, '38', 'IVA Retenido Total Chatarra')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETTOTPPA & ", 'IVA Ret. Total PPA', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Total','PPA',' ', 9, 100, -1, '39', 'IVA Retenido Total PPA')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETTOTCONSTR & ", 'IVA Ret. Total Constr.', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Total','Constr.',' ', 9, 100, -1, '41', 'IVA Retenido Total Construcción')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETTOTCARTONES & ", 'IVA Ret. Total Cartones', ' ', ' ', 0, ' ', ',3,4,5,10,11,', 'IVA Ret.Total','Cartones',' ', 9, 100, -1, '47', 'IVA Retenido Total Cartones')"
   Call ExecSQL(DbMain, Q1)


   'IVA Activo Fijo
   Q1 = "UPDATE TipoValor SET EsRecuperable = -1, TipoDoc = ',1,2,3,4,6,13,', TitCompleto = 'IVA Activo Fijo' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAACTFIJO
   Call ExecSQL(DbMain, Q1)
   

   'Ya no se usa IVA por Adq. o Const. Inmuebles, se reemplaza por IVA Cred Fiscal
   Q1 = "UPDATE TipoValor SET Atributo = 'SINUSO', CodSIIDTE = '', Valor = 'IVA por Adq. o Const. Inmuebles(*)', Tit1 = 'IVA Adq./Const.', Tit2 = 'Inmuebles(*)', TitCompleto = 'IVA por Adq. o Const. Inmuebles(*)' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAADQCONSTINMUEBLES
   Call ExecSQL(DbMain, Q1)

   'Otros impuestos
   Q1 = "UPDATE TipoValor SET CodSIIDTE = '24', Valor = 'Impto. Licores, Piscos, Destilados', Tit1 = 'Impto. Licores', Tit2 = 'Piscos, Destilados', Tasa = 31.5, EsRecuperable = -1, TitCompleto = 'Impto. Licores, Piscos, Destilados, etc.' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPPISCO
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodSIIDTE = '25', Tasa = 20.5, EsRecuperable = -1, TitCompleto = 'Impto. Vinos, Champaña, Chichas' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPVINOS
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE TipoValor SET CodSIIDTE = '26', Tasa = 20.5, EsRecuperable = -1, Valor = 'Impto. Cervezas y Bebidas Alcohólicas', Tit1 = 'Impto. Cervezas', Tit2 = 'Bebidas Alcohólicas', TitCompleto = 'Impto. Cervezas y Bebidas Alcohólicas' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPCERVEZA
   Call ExecSQL(DbMain, Q1)

   Q1 = "UPDATE TipoValor SET CodSIIDTE = '27', Tasa = 10, EsRecuperable = -1, Valor = 'Impto. Beb. Analc. y min. c/edulc.', Tit1 = 'Impto. Bebidas', Tit2 = 'Analc. y min. c/edulc.', TitCompleto = 'Impto. Beb. Analcohólicas y minerales con edulcorante' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPBEBANALC
   Call ExecSQL(DbMain, Q1)

   'ya no se usa ILA, se reemplazan por los 4 anteriores
   Q1 = "UPDATE TipoValor SET Atributo = 'SINUSO', CodSIIDTE = '', Valor = 'ILA por Notas Déb. (*)', Tit1 = 'ILA por Notas Déb.', Tit2 = '(*)', TitCompleto = 'ILA por Notas Débito recibidas (*)' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_ILANOTASDEB
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET Atributo = 'SINUSO', CodSIIDTE = '', Valor = 'ILA por Notas Créd. (*)', Tit1 = 'ILA por Notas Créd.', Tit2 = '(*)', TitCompleto = 'ILA por Notas Crédito recibidas (*)' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_ILANOTASCRED
   Call ExecSQL(DbMain, Q1)
   

   'Harina y Carne
   Q1 = "UPDATE TipoValor SET CodSIIDTE = '19', Valor = 'IVA Anticipado Harina', Tasa = 12, EsRecuperable = -1, TitCompleto = 'IVA Anticipado Harina' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAANTICIPHARINA
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET CodSIIDTE = '18', Valor = 'IVA Anticipado Carne', Tasa = 5, EsRecuperable = -1, TitCompleto = 'IVA Anticipado Carne' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAANTICIPCARNE
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAFAENACARNE & ", 'IVA Antic. Faenamiento Carne', ' ', ' ', 0, ' ', ',1,3,4,', 'IVA Antic.','Faenam. Carne',' ', 18, 5, 0, '17', 'IVA Anticipado Faenamiento Carne')"
   Call ExecSQL(DbMain, Q1)

   'Diesel
   Q1 = "UPDATE TipoValor SET CodSIIDTE = '28', Valor = 'Impto. Esp. Diesel', Tit1 = 'Impto. Esp.', Tit2 = 'Diesel', EsRecuperable = -1, Orden = 14, TitCompleto = 'Impuesto Específico Diesel' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPESPDIESEL
   Call ExecSQL(DbMain, Q1)

   'Transporte
   Q1 = "UPDATE TipoValor SET CodSIIDTE = '29', Valor = 'Impto. Esp. Diesel Transportista', Tit1 = 'Impto. Esp.', Tit2 = 'Diesel Transp.', EsRecuperable = -1, TitCompleto = 'Impuesto Específico Diesel Transportista' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IMPESPDIESELTRANS
   Call ExecSQL(DbMain, Q1)
 
   
   
   'Eliminamos todos los otros impuestos asociados a Diesel y Transporte que ya no se usan
'   Q1 = "DELETE * FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo IN(" & LIBCOMPRAS_IMPESPPETRGRAL & "," & LIBCOMPRAS_IMPESPPETRGENCF & "," & LIBCOMPRAS_IMPESPPETRGENSINCF & "," & LIBCOMPRAS_IMPESPPETRCARGACF & "," & LIBCOMPRAS_IMPESPPETRCARGASINCF & ")"
'   Call ExecSQL(DbMain, Q1)
   
   'Dejamos SIN USO los otros impuestos asociados a Diesel y Transporte que ya no se usan
   Q1 = "UPDATE TipoValor SET Atributo = 'SINUSO', CodSIIDTE = '', Valor = Valor  & ' (*)', Tit2 = Tit2 & '(*)', TitCompleto = Valor & ' (*)' "
   Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo IN(" & LIBCOMPRAS_IMPESPPETRGRAL & "," & LIBCOMPRAS_IMPESPPETRGENCF & "," & LIBCOMPRAS_IMPESPPETRGENSINCF & "," & LIBCOMPRAS_IMPESPPETRCARGACF & "," & LIBCOMPRAS_IMPESPPETRCARGASINCF & ")"
   Call ExecSQL(DbMain, Q1)
   
   
   'Eliminamos ILA por Bebidas Analcoholicas con elevado cont. Azúcar
'   Q1 = "DELETE * FROM TipoValor WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_ILABEDANALCAZUCAR
'   Call ExecSQL(DbMain, Q1)

   'Dejamos SIN USO ILA por Bebidas Analcoholicas con elevado cont. Azúcar
   Q1 = "UPDATE TipoValor SET Atributo = 'SINUSO', CodSIIDTE = '', Valor = 'ILA por Beb. Analc. Azúcar (*)', Tit1 = 'ILA Beb. Analc.', Tit2 = ' Azúcar(*)', TitCompleto = 'ILA por Bebidas Analcoholicas con elevado cont. Azúcar (*)' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_ILABEDANALCAZUCAR
   Call ExecSQL(DbMain, Q1)
   
   
   'Agregamos Otros Impuestos
   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPPIEDRASPREC & ", 'Impto. Joyas, Piedras Prec., Pieles', ' ', ' ', 0, ' ', ',1,3,4,', 'Impto. Joyas', 'Piedras Prec.', ' ', 30, 15, -1, '23', 'Impto. Joyas, Piedras Preciosas, Pieles Finas')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPALFOMBRAS & ", 'Imp. Adic. (alfombras, tapices)', ' ', ' ', 0, ' ', ',1,3,4,', 'Imp. Adic.', '(alfombras, tapices)', ' ', 31, 15, -1, '44', 'Imp. Adicional (alfombras, tapices, casas rodantes, caviar)')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVARETORO & ", 'IVA Retenido Oro', ' ', ' ', 0, ' ', ',1,3,4,', 'IVA Ret.', 'Oro' ,' ', 32, 100, -1, '46', 'IVA Retenido Oro')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPPIROTECNIA & ", 'Imp. Adic. (Pirotecnia)', ' ', ' ', 0, ' ', ',1,3,4,', 'Imp. Adic.', '(Pirotecnia)', ' ', 33, 50, -1, '45', 'Impuesto Adicional (Pirotecnia)')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAMARGCOM & ", 'IVA Margen Comer.', ' ', ' ', 0, ' ', ',1,3,4,', 'IVA Margen', 'Comer.', ' ', 34, 0, -1, '14', 'IVA de Margen de Comercialización')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPGASOLINA & ", 'Impto. Esp. Gasolina', ' ', ' ', 0, ' ', ',1,3,4,', 'Impto. Esp.', 'Gasolina' ,' ', 35, 0, -1, '35', 'Impto. Específico Gasolina')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IVAMARGCOMPREPAGO & ", 'IVA Margen Comer. Inst. Prepago', ' ', ' ', 0, ' ', ',1,3,4,', 'IVA Margen Comer.', 'Inst. Prepago', ' ', 36, 0, -1, '50', 'IVA de Margen de Comer. de Inst. de Prepago')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPGASNATURAL & ", 'Impto. Gas Natural Comprimido', ' ', ' ', 0, ' ', ',1,3,4,', 'Impto. Gas', 'Nat. Comp.', ' ', 37, 0, -1, '51', 'Impuesto Gas Natural Comprimido')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPGASLIQ & ", 'Impto. Gas Licuado', ' ', ' ', 0, ' ', ',1,3,4,', 'Impto. Gas', 'Licuado', ' ', 38, 0, -1, '52', 'Impuesto Gas Licuado de Petróleo')"
   Call ExecSQL(DbMain, Q1)

   Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
   Q1 = Q1 & " VALUES(" & LIB_COMPRAS & "," & LIBCOMPRAS_IMPSUPLEMENTEROS & ", 'Imp. Ret. Suplementeros', ' ', ' ', 0, ' ', ',1,3,4,', 'Impto. Ret.', 'Suplementeros', ' ', 39, 0.5, -1, '53', 'Imp. Retenido Suplementeros Art. 74 n° 5, LIR')"
   Call ExecSQL(DbMain, Q1)

   'asignamos un atributo para reconocer IVA Irrecuperable e IVA Activo Fijo
   Q1 = "UPDATE TipoValor SET Atributo = 'IVAIRREC' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo IN ( " & LIBCOMPRAS_IVAIRREC & "," & LIBCOMPRAS_IVAIRREC1 & "," & LIBCOMPRAS_IVAIRREC2 & "," & LIBCOMPRAS_IVAIRREC3 & "," & LIBCOMPRAS_IVAIRREC4 & "," & LIBCOMPRAS_IVAIRREC9 & ")"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET Atributo = 'IVAACTFIJO' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo = " & LIBCOMPRAS_IVAACTFIJO
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET Atributo = 'IVARETTOT' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo IN ( " & LIBCOMPRAS_IVARETTOT & "," & LIBCOMPRAS_IVARETTOTCHATARRA & "," & LIBCOMPRAS_IVARETTOTPPA & "," & LIBCOMPRAS_IVARETTOTCONSTR & "," & LIBCOMPRAS_IVARETTOTCARTONES & "," & LIBCOMPRAS_IVARETORO & ")"
   Call ExecSQL(DbMain, Q1)
   
   Q1 = "UPDATE TipoValor SET Atributo = 'IVARETPAR' WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo IN ( " & LIBCOMPRAS_IVARETPARCTRIGO & "," & LIBCOMPRAS_IVARETPARCMADERA & "," & LIBCOMPRAS_IVARETPARCGANADO & "," & LIBCOMPRAS_IVARETPARCLEGUMBRES & "," & LIBCOMPRAS_IVARETPARCARROZ & "," & LIBCOMPRAS_IVARETPARCSILVESTRES & "," & LIBCOMPRAS_IVARETPARCHIDROBIO & "," & LIBCOMPRAS_IVARETPARCFAMBPASAS & "," & LIBCOMPRAS_IMPSUPLEMENTEROS & ")"
   Call ExecSQL(DbMain, Q1)
   
   'Un documento de compra no puede tener más de un IVA activo fijo, más de in IVA Cred. Fiscal o más de un IVA Irrecuperable genérico
   Q1 = "UPDATE TipoValor SET Multiple = 0 WHERE TipoLib = " & LIB_COMPRAS & " AND Codigo IN ( " & LIBCOMPRAS_IVAACTFIJO & "," & LIBCOMPRAS_IVAIRREC & "," & LIBCOMPRAS_IVACREDFISC & ")"
   Call ExecSQL(DbMain, Q1)
   
   'eliminamos el 0 a la izquierda en el código SII de algunos documentos
   Q1 = "UPDATE TipoDocs SET CodDocSII = Val(CodDocSII), CodDocDTESII = Val(CodDocDTESII)"
   Call ExecSQL(DbMain, Q1)

End Sub


Public Function UpdateTipoPartidaCtas(ByVal TblName As String)
   Dim Q1 As String
   
   Q1 = "UPDATE " & TblName & " SET TipoPartida =1 WHERE Codigo ='3010101' AND Descripcion = 'Ingresos por Ventas de Productos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =1 WHERE Codigo ='3010102' AND Descripcion = 'Ingresos por Prestación de Servicios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =3 WHERE Codigo ='3020101' AND Descripcion = 'Intereses Cobrados a Clientes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =3 WHERE Codigo ='3020102' AND Descripcion = 'Intereses Percibidos por Prestamos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =3 WHERE Codigo ='3020105' AND Descripcion = 'Intereses por Depósitos a Plazo'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =4 WHERE Codigo ='3020106' AND Descripcion = 'Utilidad en Venta de Acciones'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =4 WHERE Codigo ='3020301' AND Descripcion = 'Venta de Activos Fijos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =4 WHERE Codigo ='3020302' AND Descripcion = 'Arriendos Obtenidos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010201' AND Descripcion = 'Costo de Venta de Productos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010202' AND Descripcion = 'Costo Directos de Prestac. De Servicios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010301' AND Descripcion = 'Sueldo Base'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010302' AND Descripcion = 'Comisiones por Ventas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010303' AND Descripcion = 'Bonos de Producción'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010304' AND Descripcion = 'Horas Extraordinarias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010305' AND Descripcion = 'Bono de Responsabilidad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010306' AND Descripcion = 'Gratificación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010307' AND Descripcion = 'Aguinaldos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010308' AND Descripcion = 'Otros Bonos Imponibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010309' AND Descripcion = 'Asignación de Movilización'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010310' AND Descripcion = 'Asignación de Colación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010311' AND Descripcion = 'Asig de Pérdida de Caja'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010312' AND Descripcion = 'Asig. Desgaste de Herramientas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010313' AND Descripcion = 'Viáticos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010314' AND Descripcion = 'Aporte Patronal'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010315' AND Descripcion = 'Otro Conceptos No Imponibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010401' AND Descripcion = 'Gastos de Uniformes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010402' AND Descripcion = 'Gastos de Capacitación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010403' AND Descripcion = 'Becas de Estudio'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010404' AND Descripcion = 'Implementos de Seguridad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010405' AND Descripcion = 'Gasto de Sala Cuna'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010406' AND Descripcion = 'Gastos de Importaciones'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010407' AND Descripcion = 'Honorarios Servicios de Terceros'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010408' AND Descripcion = 'Gastos de Aseo'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010409' AND Descripcion = 'Gastos de Mantención de Edificios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010410' AND Descripcion = 'Gastos de Remodelación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010411' AND Descripcion = 'Servicio de Mantención de Maquinarias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010412' AND Descripcion = 'Servicios Técnicos Computacionales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010413' AND Descripcion = 'Gastos de Fotocopias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010414' AND Descripcion = 'Gastos de Imprenta - Formularios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010415' AND Descripcion = 'Gastos Implementos Computacionales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010416' AND Descripcion = 'Gastos de Internet y Transmisión de Datos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010417' AND Descripcion = 'Fletes y Embalajes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010418' AND Descripcion = 'Seguros'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010419' AND Descripcion = 'Arriendos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010420' AND Descripcion = 'Gastos de Electricidad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010421' AND Descripcion = 'Gastos de Calefacción - Gas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010422' AND Descripcion = 'Gastos de Agua Potable'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010423' AND Descripcion = 'Gastos de Teléfono'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010424' AND Descripcion = 'Gastos Básicos del Mes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010425' AND Descripcion = 'Gastos de Movilización'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010426' AND Descripcion = 'Gastos de Combustibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010427' AND Descripcion = 'Gastos de Viajes (Pasajes- Hotel - Auto)'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010428' AND Descripcion = 'Gastos de Coktail'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010429' AND Descripcion = 'Gastos de Cafetería y Similares'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010430' AND Descripcion = 'Gasto Patente Municipal'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010431' AND Descripcion = 'Gasto de IVA  C.Fiscal No Utilizado'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010432' AND Descripcion = 'Depreciacion'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010433' AND Descripcion = 'Otros Gastos Directos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010442' AND Descripcion = 'I.V.A. No Recuperable'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010501' AND Descripcion = 'Sueldo Base'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010502' AND Descripcion = 'Comisiones por Ventas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010503' AND Descripcion = 'Bonos de Producción'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010504' AND Descripcion = 'Horas Extraordinarias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010505' AND Descripcion = 'Bono de Responsabilidad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010506' AND Descripcion = 'Gratificación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010507' AND Descripcion = 'Aguinaldos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010508' AND Descripcion = 'Otros Bonos Imponibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010509' AND Descripcion = 'Asignación de Movilización'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010510' AND Descripcion = 'Asignación de Colación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010511' AND Descripcion = 'Asig de Pérdida de Caja'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010512' AND Descripcion = 'Asig. Desgaste de Herramientas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010513' AND Descripcion = 'Viáticos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010514' AND Descripcion = 'Aporte Patronal'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010515' AND Descripcion = 'Otro Conceptos No Imponibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010607' AND Descripcion = 'Dieta Directorio'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010608' AND Descripcion = 'Honorarios Profesionales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010609' AND Descripcion = 'Honorarios Servicios de Terceros'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =7 WHERE Codigo ='3010701' AND Descripcion = 'Depreciación del Ejercicio'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3040101' AND Descripcion = 'Impuesto de Primera Categoría del Ejercicio'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010601' AND Descripcion = 'Gastos de Uniformes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010602' AND Descripcion = 'Gastos de Capacitación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010603' AND Descripcion = 'Becas de Estudio'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010604' AND Descripcion = 'Implementos de Seguridad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010605' AND Descripcion = 'Gasto de Sala Cuna'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010606' AND Descripcion = 'Gastos de Importaciones'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010610' AND Descripcion = 'Gastos Legales Abogados'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010611' AND Descripcion = 'Gastos Notariales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010612' AND Descripcion = 'Gastos de Suscripción'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010613' AND Descripcion = 'Gastos de Aseo'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010614' AND Descripcion = 'Gastos de Mantención de Edificios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010615' AND Descripcion = 'Gastos de Remodelación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010616' AND Descripcion = 'Servicio de Mantención de Maquinarias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010617' AND Descripcion = 'Servicios Técnicos Computacionales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010618' AND Descripcion = 'Artículos de Oficina'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010619' AND Descripcion = 'Gastos de Fotocopias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010620' AND Descripcion = 'Gastos de Imprenta - Formularios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010621' AND Descripcion = 'Gastos de Publicidad - Papelería'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010622' AND Descripcion = 'Gastos de Publicidad - Otros'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010623' AND Descripcion = 'Gastos de Regalos a Clientes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010624' AND Descripcion = 'Gastos Implementos Computacionales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010625' AND Descripcion = 'Gastos de Internet y Transmisión de Datos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010626' AND Descripcion = 'Fletes y Embalajes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010627' AND Descripcion = 'Seguros'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010628' AND Descripcion = 'Arriendos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010629' AND Descripcion = 'Gastos de Electricidad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010630' AND Descripcion = 'Gastos de Calefacción - Gas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010631' AND Descripcion = 'Gastos de Agua Potable'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010632' AND Descripcion = 'Gastos de Teléfono'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010633' AND Descripcion = 'Gastos Básicos del Mes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010634' AND Descripcion = 'Gastos de Movilización'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010635' AND Descripcion = 'Gastos de Combustibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010636' AND Descripcion = 'Gastos de Viajes (Pasajes- Hotel - Auto)'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010637' AND Descripcion = 'Gastos de Coktail'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010638' AND Descripcion = 'Gastos de Cafetería y Similares'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010639' AND Descripcion = 'Gasto Patente Municipal'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010640' AND Descripcion = 'Gasto de IVA  C.Fiscal No Utilizado'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010642' AND Descripcion = 'Impuestos de Timbres - Letras'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010643' AND Descripcion = 'Gasto de Mercaderías Obsoletas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010644' AND Descripcion = 'Estimación de Deudores Incobrables'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010645' AND Descripcion = 'Otros Gastos de Administración'"
   Call ExecSQL(DbMain, Q1)

End Function

Public Function UpdateTipoPartidaIFRS(ByVal TblName As String)
   Dim Q1 As String

   Q1 = "UPDATE " & TblName & " SET TipoPartida =1 WHERE Codigo ='3010101' AND Descripcion = 'Ingresos por Ventas de Productos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =1 WHERE Codigo ='3010102' AND Descripcion = 'Ingresos por Prestación de Servicios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =3 WHERE Codigo ='3011001' AND Descripcion = 'Intereses Cobrados a Clientes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =3 WHERE Codigo ='3011002' AND Descripcion = 'Intereses Percibidos por Préstamos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =3 WHERE Codigo ='3011003' AND Descripcion = 'Descuento por Pronto Pago Obtenido'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =4 WHERE Codigo ='3011006' AND Descripcion = 'Utilidad en Venta de Acciones'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =4 WHERE Codigo ='3010701' AND Descripcion = 'Venta de Activos Fijos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =4 WHERE Codigo ='3010702' AND Descripcion = 'Arriendos Obtenidos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010201' AND Descripcion = 'Costo de Venta de Productos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010202' AND Descripcion = 'Costo Directos de Prestación de Servicios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010203' AND Descripcion = 'Sueldo Base'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010204' AND Descripcion = 'Comisiones por Ventas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010205' AND Descripcion = 'Bonos de Producción'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010206' AND Descripcion = 'Horas Extraordinarias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010207' AND Descripcion = 'Bono de Responsabilidad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010208' AND Descripcion = 'Gratificación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010209' AND Descripcion = 'Aguinaldos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010210' AND Descripcion = 'Otros Bonos Imponibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010211' AND Descripcion = 'Asignación de Movilización'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010212' AND Descripcion = 'Asignación de Colación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010213' AND Descripcion = 'Asignacion de Pérdida de Caja'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010214' AND Descripcion = 'Asig. Desgaste de Herramientas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010215' AND Descripcion = 'Viáticos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010216' AND Descripcion = 'Aporte Patronal'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010217' AND Descripcion = 'Otro Conceptos No Imponibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010218' AND Descripcion = 'Gastos de Uniformes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010219' AND Descripcion = 'Gastos de Capacitación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010220' AND Descripcion = 'Becas de Estudio'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010221' AND Descripcion = 'Implementos de Seguridad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010222' AND Descripcion = 'Gasto de Sala Cuna'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010223' AND Descripcion = 'Gastos de Importaciones'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010224' AND Descripcion = 'Honorarios Servicios de Terceros'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010225' AND Descripcion = 'Gastos de Aseo'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010226' AND Descripcion = 'Gastos de Mantención de Edificios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010227' AND Descripcion = 'Gastos de Remodelación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010228' AND Descripcion = 'Servicio de Mantención de Maquinarias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010229' AND Descripcion = 'Servicios Técnicos Computacionales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010230' AND Descripcion = 'Gastos de Fotocopias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010231' AND Descripcion = 'Gastos de Imprenta  Formularios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010232' AND Descripcion = 'Gastos Implementos Computacionales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010233' AND Descripcion = 'Gastos de Internet y Transmisión de Datos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010234' AND Descripcion = 'Fletes y Embalajes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010235' AND Descripcion = 'Seguros'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010236' AND Descripcion = 'Arriendos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010237' AND Descripcion = 'Gastos de Electricidad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010238' AND Descripcion = 'Gastos de Calefacción  Gas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010239' AND Descripcion = 'Gastos de Agua Potable'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010240' AND Descripcion = 'Gastos de Teléfono'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010241' AND Descripcion = 'Gastos Básicos del Mes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010242' AND Descripcion = 'Gastos de Movilización'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010243' AND Descripcion = 'Gastos de Combustibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010244' AND Descripcion = 'Gastos de Viajes (Pasajes Hotel  Auto )'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010245' AND Descripcion = 'Gastos de Cocktail'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010246' AND Descripcion = 'Gastos de Cafetería y Similares'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010247' AND Descripcion = 'Gasto Patente Municipal'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010248' AND Descripcion = 'Gasto de IVA  Crédito Fiscal No Utilizado'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010249' AND Descripcion = 'Depreciación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010250' AND Descripcion = 'Deterioro de Valor'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010251' AND Descripcion = 'Otros Gastos Directos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =5 WHERE Codigo ='3010252' AND Descripcion = 'I.V.A. No Recuperable'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010501' AND Descripcion = 'Sueldo Base'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010502' AND Descripcion = 'Comisiones por Ventas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010503' AND Descripcion = 'Bonos de Producción'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010504' AND Descripcion = 'Horas Extraordinarias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010505' AND Descripcion = 'Bono de Responsabilidad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010506' AND Descripcion = 'Gratificación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010507' AND Descripcion = 'Aguinaldos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010508' AND Descripcion = 'Otros Bonos Imponibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010509' AND Descripcion = 'Asignación de Movilización'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010510' AND Descripcion = 'Asignación de Colación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010511' AND Descripcion = 'Asig de Pérdida de Caja'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010512' AND Descripcion = 'Asig. Desgaste de Herramientas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010513' AND Descripcion = 'Viáticos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010514' AND Descripcion = 'Aporte Patronal'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010515' AND Descripcion = 'Otro Conceptos No Imponibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010522' AND Descripcion = 'Dieta Directorio'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010523' AND Descripcion = 'Honorarios Profesionales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =6 WHERE Codigo ='3010524' AND Descripcion = 'Honorarios Servicios de Terceros'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =7 WHERE Codigo ='3010561' AND Descripcion = 'Depreciación del Ejercicio'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3011701' AND Descripcion = 'Impuesto de Primera Categoría del Ejercicio'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010516' AND Descripcion = 'Gastos de Uniformes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010517' AND Descripcion = 'Gastos de Capacitación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010518' AND Descripcion = 'Becas de Estudio'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010519' AND Descripcion = 'Implementos de Seguridad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010520' AND Descripcion = 'Gasto de Sala Cuna'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010521' AND Descripcion = 'Gastos de Importaciones'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010525' AND Descripcion = 'Gastos Legales Abogados'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010526' AND Descripcion = 'Gastos Notariales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010527' AND Descripcion = 'Gastos de Suscripción'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010528' AND Descripcion = 'Gastos de Aseo'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010529' AND Descripcion = 'Gastos de Mantención de Edificios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010530' AND Descripcion = 'Gastos de Remodelación'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010531' AND Descripcion = 'Servicio de Mantención de Maquinarias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010532' AND Descripcion = 'Servicios Técnicos Computacionales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010533' AND Descripcion = 'Artículos de Oficina'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010534' AND Descripcion = 'Gastos de Fotocopias'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010535' AND Descripcion = 'Gastos de Imprenta  Formularios'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010536' AND Descripcion = 'Gastos de Publicidad  Papelería'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010537' AND Descripcion = 'Gastos de Publicidad  Otros'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =14 WHERE Codigo ='3010538' AND Descripcion = 'Gastos de Regalos a Clientes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010539' AND Descripcion = 'Gastos Implementos Computacionales'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010540' AND Descripcion = 'Gastos de Internet y Transmisión de Datos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010541' AND Descripcion = 'Fletes y Embalajes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010542' AND Descripcion = 'Seguros'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010543' AND Descripcion = 'Arriendos'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010544' AND Descripcion = 'Gastos de Electricidad'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010545' AND Descripcion = 'Gastos de Calefacción  Gas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010546' AND Descripcion = 'Gastos de Agua Potable'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010547' AND Descripcion = 'Gastos de Teléfono'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010548' AND Descripcion = 'Gastos Básicos del Mes'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010549' AND Descripcion = 'Gastos de Movilización'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010550' AND Descripcion = 'Gastos de Combustibles'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010551' AND Descripcion = 'Gastos de Viajes (Pasajes Hotel  Auto )'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010552' AND Descripcion = 'Gastos de Cocktail'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010553' AND Descripcion = 'Gastos de Cafetería y Similares'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010554' AND Descripcion = 'Gasto Patente Municipal'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010555' AND Descripcion = 'Gasto de IVA  C.Fiscal No Utilizado'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010557' AND Descripcion = 'Impuestos de Timbres  Letras'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010558' AND Descripcion = 'Gasto de Mercaderías Obsoletas'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010559' AND Descripcion = 'Estimación de Deudores Incobrables'"
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE " & TblName & " SET TipoPartida =16 WHERE Codigo ='3010560' AND Descripcion = 'Otros Gastos de Administración'"
   Call ExecSQL(DbMain, Q1)

End Function

Private Function CrearTblPlanCuentasSII()
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   CrearTblPlanCuentasSII = True
      
   '--------------------- Crear tabla PlanCuentasSII -----------------
   
   Set Tbl = New TableDef
   Tbl.Name = "PlanCuentasSII"
   
   Err.Clear
   Set Fld = Tbl.CreateField("IdPlanCuentasSII", dbLong)
   Fld.Attributes = dbAutoIncrField
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "PlanCuentasSII.IdPlanCuentasSII", vbExclamation
      CrearTblPlanCuentasSII = False
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("CodigoSII", dbText, 10)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "PlanCuentasSII.CodigoSII", vbExclamation
      CrearTblPlanCuentasSII = False
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("DescripSII", dbText, 130)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "PlanCuentasSII.DescripSII", vbExclamation
      CrearTblPlanCuentasSII = False
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("FmtCodigoSII", dbText, 15)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "PlanCuentasSII.FmtCodigoSII", vbExclamation
      CrearTblPlanCuentasSII = False
   End If
   
   Err.Clear
   Set Fld = Tbl.CreateField("Clasificacion", dbByte)
   Tbl.Fields.Append Fld
   
   If Err = 0 Then
      Tbl.Fields.Refresh
   ElseIf Err <> 3191 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "PlanCuentasSII.Clasificacion", vbExclamation
      CrearTblPlanCuentasSII = False
   End If
   
   
   DbMain.TableDefs.Append Tbl
   If Err = 0 Then
      DbMain.TableDefs.Refresh
      
      Q1 = "CREATE UNIQUE INDEX IdxId ON PlanCuentasSII (IdPlanCuentasSII) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE UNIQUE INDEX IdxCod ON PlanCuentasSII (CodigoSII)"
      Rc = ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE INDEX IdClasif ON PlanCuentasSII (Clasificacion)"
      Rc = ExecSQL(DbMain, Q1)
      
   ElseIf Err <> 3010 Then ' ya existe
      MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla PlanCuentasSII", vbExclamation
      CrearTblPlanCuentasSII = False
      
   End If
   
   Set Tbl = Nothing
End Function

Private Sub FillPlanCuentasSII()
   Dim Q1 As String
   Dim QBase As String
   Dim QEnd As String
   
   QBase = "INSERT INTO PlanCuentasSII (CodigoSII, DescripSII, FmtCodigoSII, Clasificacion) VALUES("
   
   'agregamos primero las cuentas de Activo
   
   QEnd = "," & CLASCTA_ACTIVO & ")"
   
   Q1 = "'1010100', 'Disponible', '1.01.01.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1010300', 'Depósitos a plazo', '1.01.03.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1010500', 'Valores negociables', '1.01.05.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1010700', 'Instrumentos derivados', '1.01.07.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1010900', 'Pactos Retrocompra- Retroventa', '1.01.09.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1011500', 'Inversiones en el Exterior', '1.01.15.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1012000', 'Deudores por venta, neto (excluye deudores por leasing)', '1.01.20.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1012100', 'Deudores por Leasing', '1.01.21.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1012500', 'Documentos por cobrar', '1.01.25.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1013000', 'Deudores varios', '1.01.30.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1014000', 'Documentos y cuentas por cobrar empresas relacionadas situadas en Chile (cuenta corriente mercantil)', '1.01.40.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1014100', 'Documentos y cuentas por cobrar empresas relacionadas situadas en el Extranjero (cuenta corriente mercantil)', '1.01.41.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1015000', 'Existencias, neto', '1.01.50.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1015100', 'Activos Biológicos, neto', '1.01.51.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1015500', 'Existencias en Tránsito', '1.01.55.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1015900', 'IVA Crédito Fiscal', '1.01.59.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1016000', 'Impuestos por recuperar', '1.01.60.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1016100', 'Créditos por Donaciones', '1.01.61.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1016200', 'Otros Créditos por recuperar', '1.01.62.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1017000', 'Bienes entregados en leasing', '1.01.70.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1019900', 'Otros activos corrientes', '1.01.99.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1021000', 'Propiedad Planta y Equipos y Otros  (excepto bienes entregados en Leasing)', '1.02.10.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1021100', 'Terrenos', '1.02.11.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1021200', 'Construcción y obras de infraestructura', '1.02.12.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1021300', 'Maquinarias y equipos', '1.02.13.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1021400', 'Muebles y utiles', '1.02.14.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1021500', 'Equipos Computacionales y similares', '1.02.15.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1021600', 'Automóviles', '1.02.16.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1021700', 'Vehículos', '1.02.17.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1021800', 'Barcos y Aviones', '1.02.18.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1021900', 'Propiedades de Inversion', '1.02.19.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1022500', 'Software', '1.02.25.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1022600', 'Concesiones', '1.02.26.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1022700', 'Obras en Ejecución', '1.02.27.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1022800', 'Obras en Ejecución', '1.02.28.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1023000', 'Activos en Leasing', '1.02.30.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1029000', 'Depreciación Acumulada (excepto Automoviles y Activos en Leasing)', '1.02.90.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1029200', 'Depreciación Acumulada Automóviles', '1.02.92.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1029500', 'Depreciación Acumulada Activos en Leasing', '1.02.95.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1029900', 'Otros Bienes Propiedad Planta y Equipo', '1.02.99.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1030100', 'Inversiones en empresas relacionadas', '1.03.01.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1030300', 'Menor valor de inversiones (Plusvalias, Goodwill)', '1.03.03.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1030400', 'Mayor valor de inversiones (Minusvalias, Badwill)', '1.03.04.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1030500', 'Cuenta Particular Socio', '1.03.05.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1031000', 'Inversiones en otras sociedades en Chile', '1.03.10.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1031000', 'Inversiones en otras sociedades en el extranjero', '1.03.10.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1031500', 'Cuenta en participacion', '1.03.15.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1031600', 'Inversion en Agencias', '1.03.16.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1032000', 'Deudores a largo plazo', '1.03.20.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1032400', 'Anticipo y préstamos a los empleados', '1.03.24.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1032500', 'Anticipo a proveedores', '1.03.25.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1033000', 'Gastos pagados por anticipado', '1.03.30.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1033100', 'Gastos de Investigación y Desarrollo', '1.03.31.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1033200', 'Gastos Diferidos', '1.03.32.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1033300', 'Menor Valor en Colocacion de bonos', '1.03.33.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1034000', 'Intereses Diferidos por Leasing', '1.03.40.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1034100', 'Otros Intereses Diferidos', '1.03.41.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1034500', 'Garantias', '1.03.45.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1035000', 'Impuestos diferidos', '1.03.50.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1036000', 'Intangibles distintos a la Plusvalia (neto)', '1.03.60.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1037000', 'Posicion de Cambio', '1.03.70.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1037100', 'Intereses Suspendidos', '1.03.71.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1039800', 'Cuentas de Orden de Activos', '1.03.98.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'1039900', 'Otros Activos No Corrientes', '1.03.99.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   'ahora las cuentas de pasivo
   
   QEnd = "," & CLASCTA_PASIVO & ")"
  
   Q1 = "'2010100', 'Obligaciones con bancos e instituciones financieras', '2.01.01.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2010300', 'Obligaciones con el público (Bonos Emitidos)', '2.01.03.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2010400', 'Obligaciones por Leasing', '2.01.04.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2010700', 'Instrumentos derivados', '2.01.07.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2010800', 'Fondo Opcion de Compra por Pagar (Leasing)', '2.01.08.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2011000', 'Cuentas por pagar', '2.01.10.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2011100', 'Proveedores por Pagar', '2.01.11.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2011200', 'Acreedores varios', '2.01.12.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2011400', 'Documentos por pagar', '2.01.14.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2012000', 'Dividendos por pagar', '2.01.20.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2014000', 'Documentos y cuentas por pagar empresas relacionadas situadas en Chile (cuenta corriente mercantil)', '2.01.40.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2014100', 'Documentos y cuentas por pagar empresas relacionadas situadas en el Extranjero (cuenta corriente mercantil)', '2.01.41.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2015000', 'Provision de Indemnización', '2.01.50.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2015100', 'Provisiones por Vacaciones, por Bonos y por otros Beneficios a los Empleados', '2.01.51.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2015400', 'Otras Provisiones', '2.01.54.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2015500', 'Retenciones por Pagar', '2.01.55.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2015900', 'IVA Débito Fiscal', '2.01.59.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2016000', 'Impuesto a la renta por Pagar', '2.01.60.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2016100', 'Otros Impuestos por Pagar', '2.01.61.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2016200', 'Ingresos percibidos por adelantado', '2.01.62.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2017000', 'Anticipo de Clientes', '2.01.70.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2019900', 'Otros pasivos Corrientes', '2.01.99.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2020100', 'Obligaciones con bancos e instituciones financieras', '2.02.01.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2020200', 'Obligaciones con el público (Bonos Emitidos)', '2.02.02.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2020300', 'Documentos por pagar largo plazo', '2.02.03.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2020400', 'Acreedores varios largo plazo', '2.02.04.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2020600', 'Provisiones', '2.02.06.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2020700', 'Impuestos diferidos', '2.02.07.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2029800', 'Cuentas de Orden de Pasivos', '2.02.98.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2029900', 'Otros pasivos NO Corrientes', '2.02.99.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2030100', 'Capital pagado', '2.03.01.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2030200', 'Reserva revalorización capital', '2.03.02.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2030300', 'Sobreprecio en venta de acciones propias', '2.03.03.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2030400', 'Otras reservas', '2.03.04.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2030500', 'Reservas futuros dividendos', '2.03.05.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2030600', 'Utilidades acumuladas', '2.03.06.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2030700', 'Pérdidas acumuladas', '2.03.07.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2030800', 'Dividendos provisorios', '2.03.08.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2030900', 'Cuenta Obligada Socio', '2.03.09.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2032000', 'Reserva Ajuste IFRS por 1a Aplicación', '2.03.20.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2032100', 'Reserva Ajuste IFRS', '2.03.21.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2033000', 'Valor Mercado Intrumentos Derivados de Cobertura acogidos Ley 20.544', '2.03.30.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2033100', 'Valor Mercado Intrumentos Derivados de Cobertura No acogidos Ley 20.544', '2.03.31.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'2039900', 'Otros ajustes patrimoniales', '2.03.99.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
      
   'finalmente las cuentas de resultado
   
   QEnd = "," & CLASCTA_RESULTADO & ")"


   Q1 = "'3010100', 'Ingresos de explotación', '3.01.01.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3010200', 'Costos de explotación', '3.01.02.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3010300', 'Gastos de administración y ventas', '3.01.03.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3011100', 'Ingresos de explotación con partes relacionadas del exterior', '3.01.11.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3011200', 'Costos de explotación con relacionados del exterior', '3.01.12.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3011300', 'Gastos de administración y ventas con relacionados del exterior', '3.01.13.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3020100', 'Ingresos financieros', '3.02.01.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3020200', 'Utilidad (pérdida) inversiones empresas relacionadas', '3.02.02.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3020300', 'Otros ingresos fuera de la explotación', '3.02.03.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3020600', 'Gastos financieros con empresas relacionadas', '3.02.06.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3020700', 'Gastos financieros con empresas no relacionadas', '3.02.07.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3020800', 'Resultado por Instrumentos Derivados', '3.02.08.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3021100', 'Ingresos financieros con partes relacionadas del exterior', '3.02.11.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3021600', 'Gastos financieros con partes relacionadas del exterior', '3.02.16.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3023000', 'Intereses percibidos o devengados con partes relacionadas del exterior', '3.02.30.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3023100', 'Intereses pagados o adeudados con partes relacionadas del exterior', '3.02.31.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3030100', 'Depreciacion', '3.03.01.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3030200', 'Deterioros', '3.03.02.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3030300', 'Amortización Intangibles distintos a las Plusvalias', '3.03.03.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3030400', 'Amortización menor valor de inversiones (Goodwill)', '3.03.04.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3030500', 'Amortización mayor valor de inversiones', '3.03.05.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3030600', 'Castigos', '3.03.06.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3040100', 'Valor Mercado Instrumentos Derivados acogidos Ley 20.544', '3.04.01.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3040200', 'Valor Mercado Instrumentos Derivados NO acogidos Ley 20.544', '3.04.02.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3040300', 'Ajuste Valor Mercado Existencias (VNR) y Activos Biologicos', '3.04.03.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3040400', 'Ajuste Valor Mercado Propiedad Planta y Equipo, Propiedad Inversion, Activos No Corrientes Mantenidos para la Venta', '3.04.04.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3040500', 'Ajuste Valor Mercado Fondos Mutuos', '3.04.05.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3040600', 'Ajuste Valor Mercado Valores Negociables', '3.04.06.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3040700', 'Ajuste Valor Neto Realización', '3.04.07.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3049900', 'Otros Ajustes a Valor Mercado', '3.04.99.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3050100', 'Resultado por la Enajenacion de Inversiones Permanentes', '3.05.01.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3050200', 'Resultado por la Enajenacion de Inversiones en otras sociedades', '3.05.02.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3050500', 'Resultado por la Enajenacion de Inversiones en Valores Negociables', '3.05.05.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3050700', 'Resultado por la Enajenacion Propiedad Planta y Equipos,Propiedad de Inversion y Activos no Corrientes mantenidos para la venta', '3.05.07.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3050800', 'Resultado enajenación Activo Fijo', '3.05.08.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3051000', 'Otros egresos fuera de la explotación', '3.05.10.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3051100', 'Corrección monetaria', '3.05.11.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3051200', 'Diferencias de cambio', '3.05.12.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3051300', 'Donaciones', '3.05.13.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3051500', 'Intereses,Multas y Reajustes', '3.05.15.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3051600', 'Patentes Municipales', '3.05.16.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3051700', 'Otros Impuestos', '3.05.17.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3060100', 'Impuesto a La Renta', '3.06.01.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)
   
   Q1 = "'3060200', 'Impuesto Diferido', '3.06.02.00'"
   Call ExecSQL(DbMain, QBase & Q1 & QEnd)

End Sub
Public Function UpdateCtaPlanSII(ByVal TblName As String)
   Dim Q1 As String
   Dim QBase As String
   
   QBase = "UPDATE " & TblName & " "
   
   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010101' AND Descripcion = 'Caja Principal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010102' AND Descripcion = 'Caja Chica - Fondo Fijo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010103' AND Descripcion = 'Depósitos en Transito'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010104' AND Descripcion = 'Banco'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010201' AND Descripcion = 'Depósitos a Plazo en $'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010202' AND Descripcion = 'Depósitos a Plazo en U.F.'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010203' AND Descripcion = 'Fondos Mutuos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010500' WHERE Codigo = '1010301' AND Descripcion = 'Acciones de S.A. en Bolsa'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010500' WHERE Codigo = '1010302' AND Descripcion = 'Acciones de S.A. No Bolsa'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010500' WHERE Codigo = '1010303' AND Descripcion = 'Mayor Valor Bursátil Acciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012000' WHERE Codigo = '1010401' AND Descripcion = 'Clientes Ventas a Crédito'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012000' WHERE Codigo = '1010402' AND Descripcion = 'Clientes Ventas Al Contado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012000' WHERE Codigo = '1010403' AND Descripcion = 'Clientes Extranjeros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012000' WHERE Codigo = '1010404' AND Descripcion = 'Estimac. Deudores Incobrables'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010501' AND Descripcion = 'Cheques a Fecha por Cobrar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010502' AND Descripcion = 'Letras por Cobrar en Cartera'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010503' AND Descripcion = 'Letras por Cobrar en  Cob. Bancaria'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010504' AND Descripcion = 'Documentos Protestados en Cartera'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010505' AND Descripcion = 'Documentos en Cobranza Judicial'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010506' AND Descripcion = 'Tarjetas de Crédito por Cobrar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010507' AND Descripcion = 'Cheques Enviados a Factoring'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010508' AND Descripcion = 'Letras Enviados a Factoring'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010509' AND Descripcion = 'Estimac. Doctos. Incobrables'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010599' AND Descripcion = 'Otros Documentos por Cobrar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1030500' WHERE Codigo = '1010601' AND Descripcion = 'Cuenta Corriente Socio A'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1030500' WHERE Codigo = '1010602' AND Descripcion = 'Cuenta Corriente Socio B'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010603' AND Descripcion = 'Cuenta Corriente Merc. Empleados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010604' AND Descripcion = 'Anticipos de Sueldo por Descontar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010605' AND Descripcion = 'Anticipo a Proveedores'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010606' AND Descripcion = 'Anticipo de Honorarios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010607' AND Descripcion = 'Fondos por Rendir'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010608' AND Descripcion = 'Anticipos de Gratificación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010609' AND Descripcion = 'Asig. Familiar por Recuperar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010610' AND Descripcion = 'Prestamos Socios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010611' AND Descripcion = 'Arriendos en Garantía por Recuperar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010699' AND Descripcion = 'Otros Deudores por Cobrar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010801' AND Descripcion = 'Materias Primas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010802' AND Descripcion = 'Materiales Directos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010803' AND Descripcion = 'Productos en Proceso'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010804' AND Descripcion = 'Productos Terminados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010805' AND Descripcion = 'Mercaderías Nacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015500' WHERE Codigo = '1010806' AND Descripcion = 'Importaciones en Tránsito'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010898' AND Descripcion = 'Otras Existencias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010899' AND Descripcion = 'Provisión Obsoles. Mercaderías'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016200' WHERE Codigo = '1010901' AND Descripcion = 'Pagos Provisionales Mensuales (Ppm)'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015900' WHERE Codigo = '1010902' AND Descripcion = 'I.V.A. Crédito Fiscal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016200' WHERE Codigo = '1010903' AND Descripcion = 'Crédito Sence'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016100' WHERE Codigo = '1010904' AND Descripcion = 'Crédito por Donaciones a Universidades'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016200' WHERE Codigo = '1010905' AND Descripcion = 'Crédito por Compras de Activo Fijo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016200' WHERE Codigo = '1010906' AND Descripcion = 'Crédito por Dev. Impto. Renta'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016000' WHERE Codigo = '1010999' AND Descripcion = 'Otros Impuestos por Recuperar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1019900' WHERE Codigo = '1011001' AND Descripcion = 'Seguros Pagados por Anticipado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1019900' WHERE Codigo = '1011002' AND Descripcion = 'Gasto Patente Municipal a Diferir Semestre'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1019900' WHERE Codigo = '1011003' AND Descripcion = 'Gasto de Publicidad Diferida'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1019900' WHERE Codigo = '1011099' AND Descripcion = 'Otros Gastos Pagados por Anticipado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1019900' WHERE Codigo = '1011101' AND Descripcion = 'Impuestos Diferidos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1019900' WHERE Codigo = '1011199' AND Descripcion = 'Otros Impuestos Diferidos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1017000' WHERE Codigo = '1011201' AND Descripcion = 'Activos Fijos en Leasing'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1017000' WHERE Codigo = '1011202' AND Descripcion = 'Intereses Diferidos Leasing C.P.'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1019900' WHERE Codigo = '1011301' AND Descripcion = 'Otros Activos Circulantes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021100' WHERE Codigo = '1020101' AND Descripcion = 'Terreno 1'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021100' WHERE Codigo = '1020102' AND Descripcion = 'Terreno 2'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021000' WHERE Codigo = '1020201' AND Descripcion = 'Edificios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021200' WHERE Codigo = '1020202' AND Descripcion = 'Construcciones en Propiedades de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021000' WHERE Codigo = '1020203' AND Descripcion = 'Instalaciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021200' WHERE Codigo = '1020299' AND Descripcion = 'Otras Construcc. Y Obras de Infraestructura'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021300' WHERE Codigo = '1020301' AND Descripcion = 'Maquinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021300' WHERE Codigo = '1020302' AND Descripcion = 'Equipos Industriales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021500' WHERE Codigo = '1020303' AND Descripcion = 'Equipos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021400' WHERE Codigo = '1020304' AND Descripcion = 'Muebles de Oficina'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1022500' WHERE Codigo = '1020305' AND Descripcion = 'Sofware'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021700' WHERE Codigo = '1020306' AND Descripcion = 'Vehículos Aceptados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021600' WHERE Codigo = '1020307' AND Descripcion = 'Vehículos Rechazados (Automóviles)'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021300' WHERE Codigo = '1020399' AND Descripcion = 'Otras Maquinarias y Equipos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021000' WHERE Codigo = '1020401' AND Descripcion = 'Otros Activos Fijos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029900' WHERE Codigo = '1020501' AND Descripcion = 'Mayor Valor Retazación Técnica Act. Fijo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020601' AND Descripcion = 'Depreciación de Edificios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020602' AND Descripcion = 'Depreciación de Construcciones en Propiedades de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020603' AND Descripcion = 'Depreciación de Instalaciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020604' AND Descripcion = 'Depreciación de Otras Construcc. Y Obras de Infraestructura'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020605' AND Descripcion = 'Depreciación de Maquinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020606' AND Descripcion = 'Depreciación de Equipos Industriales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020607' AND Descripcion = 'Depreciación de Equipos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020608' AND Descripcion = 'Depreciación de Muebles de Oficina'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029200' WHERE Codigo = '1020610' AND Descripcion = 'Depreciación de Vehículos Aceptados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029200' WHERE Codigo = '1020611' AND Descripcion = 'Depreciación de Vehículos Rechazados (Automóviles)'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020612' AND Descripcion = 'Depreciación de Otras Maquinarias y Equipos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020613' AND Descripcion = 'Depreciación de Otros Activos Fijos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020614' AND Descripcion = 'Depreciación de Mayor Valor Relación Técnica Act. Fijo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020701' AND Descripcion = 'Dep. Acum. De Edificios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020702' AND Descripcion = 'Dep. Acum. De Construcciones en Propiedades de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020703' AND Descripcion = 'Dep. Acum. De Instalaciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020704' AND Descripcion = 'Dep. Acum. De Otras Construcc. Y Obras de Infraestructura'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020705' AND Descripcion = 'Dep. Acum. De Maquinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020706' AND Descripcion = 'Dep. Acum. De Equipos Industriales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020707' AND Descripcion = 'Dep. Acum. De Equipos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020708' AND Descripcion = 'Dep. Acum. De Muebles de Oficina'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029200' WHERE Codigo = '1020710' AND Descripcion = 'Dep. Acum. De Vehículos Aceptados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029200' WHERE Codigo = '1020711' AND Descripcion = 'Dep. Acum. De Vehículos Rechazados (Automóviles)'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020712' AND Descripcion = 'Dep. Acum. De Otras Maquinarias y Equipos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020713' AND Descripcion = 'Dep. Acum. De Otros Activos Fijos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020714' AND Descripcion = 'Dep. Acum. De Mayor Valor Relación Técnica Act. Fijo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1030100' WHERE Codigo = '1030101' AND Descripcion = 'Inversiones en Empresas Relacionadas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1030300' WHERE Codigo = '1030301' AND Descripcion = 'Menor Valor de Inversiones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1030400' WHERE Codigo = '1030401' AND Descripcion = 'Mayor Valor de Inversiones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1032000' WHERE Codigo = '1030501' AND Descripcion = 'Deudores a Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1039900' WHERE Codigo = '1030601' AND Descripcion = 'Cuentas por Cobrar a Emp. Relacionada L.P.'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1039900' WHERE Codigo = '1030602' AND Descripcion = 'Doctos por Cobrar a Emp. Relacionada L.P.'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1039900' WHERE Codigo = '1030603' AND Descripcion = 'Otros Doctos y Ctas x Cobrar a EE.RR. L. P.'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1035000' WHERE Codigo = '1030701' AND Descripcion = 'Impuestos Diferidos a Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1035000' WHERE Codigo = '1030702' AND Descripcion = 'Otros Impuestos Diferidos a Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1036000' WHERE Codigo = '1030801' AND Descripcion = 'Derechos de Llaves Estimados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1036000' WHERE Codigo = '1030802' AND Descripcion = 'Derechos de Marca Estimados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1036000' WHERE Codigo = '1030803' AND Descripcion = 'Otros Activos Intangibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1033100' WHERE Codigo = '1030901' AND Descripcion = 'Gastos de Organización y Puesta en Marcha'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1039900' WHERE Codigo = '1030902' AND Descripcion = 'Líneas Telefónicas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1036000' WHERE Codigo = '1030903' AND Descripcion = 'Derechos de Marca Pagados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1034000' WHERE Codigo = '1031202' AND Descripcion = 'Intereses Diferidos Leasing L.P.'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2010100' WHERE Codigo = '2010101' AND Descripcion = 'Obligaciones Con Bancos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2010100' WHERE Codigo = '2010201' AND Descripcion = 'Obligaciones Con Bancos e Instit. Financ. L. P. - Porción C. Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2029900' WHERE Codigo = '2010301' AND Descripcion = 'Obligaciones a Largo Plazo Con Vencimiento Dentro de Un Año'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011200' WHERE Codigo = '2010401' AND Descripcion = 'Acreedores por Leasing Corto Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011200' WHERE Codigo = '2010501' AND Descripcion = 'Obligaciones por Factoring (Neto)'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011100' WHERE Codigo = '2010601' AND Descripcion = 'Proveedores Nacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011100' WHERE Codigo = '2010602' AND Descripcion = 'Proveedores Extranjeros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011100' WHERE Codigo = '2010603' AND Descripcion = 'Cuentas por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011100' WHERE Codigo = '2010604' AND Descripcion = 'Honorarios por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011100' WHERE Codigo = '2010605' AND Descripcion = 'Cheques a Fecha por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011100' WHERE Codigo = '2010606' AND Descripcion = 'Letras por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011100' WHERE Codigo = '2010607' AND Descripcion = 'Facturas por Contabilizar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011100' WHERE Codigo = '2010699' AND Descripcion = 'Otras Cuentas por Pagar del Giro'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011000' WHERE Codigo = '2010801' AND Descripcion = 'Prestamos por Pagar a Socio 1 C. P.'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2017000' WHERE Codigo = '2010901' AND Descripcion = 'Ingresos Percibidos por Adelantado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2017000' WHERE Codigo = '2010902' AND Descripcion = 'Anticipos de Clientes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016200' WHERE Codigo = '2010903' AND Descripcion = 'Arriendos Recibidos en  Garantía'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015100' WHERE Codigo = '2011001' AND Descripcion = 'Provisiones Vacaciones del Personal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015400' WHERE Codigo = '2011002' AND Descripcion = 'Provisión Honorarios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015400' WHERE Codigo = '2011003' AND Descripcion = 'Provisión Gastos por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015100' WHERE Codigo = '2011004' AND Descripcion = 'Provisión Remuneraciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015400' WHERE Codigo = '2011099' AND Descripcion = 'Otras Provisiones de Gastos por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011000' WHERE Codigo = '2011101' AND Descripcion = 'Remuneraciones por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2011102' AND Descripcion = 'Afp por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2011103' AND Descripcion = 'Isapres por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2011104' AND Descripcion = 'Achs por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2011105' AND Descripcion = 'Inp por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2011106' AND Descripcion = 'C.C.A.F. Por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2011107' AND Descripcion = 'Prestamos Ccaf por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2011108' AND Descripcion = 'Descuentos Conv. Empleados por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011000' WHERE Codigo = '2011109' AND Descripcion = 'Finiquitos por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011000' WHERE Codigo = '2011110' AND Descripcion = 'Remuneraciones No Cobradas por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2011199' AND Descripcion = 'Otras Retenciones por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015900' WHERE Codigo = '2011201' AND Descripcion = 'I.V.A. Débito Fiscal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016100' WHERE Codigo = '2011202' AND Descripcion = 'Impuesto de 2° Categoría por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016100' WHERE Codigo = '2011203' AND Descripcion = 'Impuesto Único Al Trabajo por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016100' WHERE Codigo = '2011204' AND Descripcion = 'P.P.M Por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016100' WHERE Codigo = '2011205' AND Descripcion = 'Contribuciones de Bs. Rs. Por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2011206' AND Descripcion = 'Seguro de Cesantía por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016100' WHERE Codigo = '2011299' AND Descripcion = 'Otros Impuestos por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016000' WHERE Codigo = '2011301' AND Descripcion = 'Impuestos a la Renta por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016100' WHERE Codigo = '2011401' AND Descripcion = 'Impuestos Diferidos por Pagar  a Corto Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2019900' WHERE Codigo = '2011402' AND Descripcion = 'Derechos de Aduana Diferidos por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030800' WHERE Codigo = '2011501' AND Descripcion = 'Dividendos por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011200' WHERE Codigo = '2011601' AND Descripcion = 'Otros Acreedores  Corto Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020100' WHERE Codigo = '2020101' AND Descripcion = 'Obligaciones Con Bancos e Inst. Financ. A Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020400' WHERE Codigo = '2020201' AND Descripcion = 'Acreedores por Leasing Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020400' WHERE Codigo = '2020301' AND Descripcion = 'Proveedores de Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020300' WHERE Codigo = '2020401' AND Descripcion = 'Cuentas por Pagar a Empresas Relacionadas a Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020400' WHERE Codigo = '2020501' AND Descripcion = 'Prestamos por Pagar a Socios Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015100' WHERE Codigo = '2020601' AND Descripcion = 'Provisiones Indem. Años Servicio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020600' WHERE Codigo = '2020699' AND Descripcion = 'Otras Provisiones a L. Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020700' WHERE Codigo = '2020701' AND Descripcion = 'Impuestos Diferidos por Pagar  a Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2029900' WHERE Codigo = '2020702' AND Descripcion = 'Derechos de Aduana por Pagar a Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020400' WHERE Codigo = '2020801' AND Descripcion = 'Otros Acreedores  a Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030100' WHERE Codigo = '2030101' AND Descripcion = 'Capital Pagado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030200' WHERE Codigo = '2030201' AND Descripcion = 'Reserva Rev. De Capital'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030400' WHERE Codigo = '2030301' AND Descripcion = 'Reservas Varias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030500' WHERE Codigo = '2030401' AND Descripcion = 'Reservas Futuros Dividendos - Repartos Utilidades Sociales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030400' WHERE Codigo = '2030501' AND Descripcion = 'Reservas Ret. Técnica Activo Fijo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030600' WHERE Codigo = '2031101' AND Descripcion = 'Utilidad Neta Retenida'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030600' WHERE Codigo = '2031201' AND Descripcion = 'Utilidad o Pérdida Neta del Periodo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030800' WHERE Codigo = '2031301' AND Descripcion = '         Dividendos Anticipados/Cuenta Obligada Socios/Retiros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010100' WHERE Codigo = '3010101' AND Descripcion = 'Ingresos por Ventas de Productos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010100' WHERE Codigo = '3010102' AND Descripcion = 'Ingresos por Prestación de Servicios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010201' AND Descripcion = 'Costo de Venta de Productos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010202' AND Descripcion = 'Costo Directos de Prestac. De Servicios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010301' AND Descripcion = 'Sueldo Base'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010302' AND Descripcion = 'Comisiones por Ventas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010303' AND Descripcion = 'Bonos de Producción'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010304' AND Descripcion = 'Horas Extraordinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010305' AND Descripcion = 'Bono de Responsabilidad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010306' AND Descripcion = 'Gratificación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010307' AND Descripcion = 'Aguinaldos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010308' AND Descripcion = 'Otros Bonos Imponibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010309' AND Descripcion = 'Asignación de Movilización'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010310' AND Descripcion = 'Asignación de Colación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010311' AND Descripcion = 'Asig de Pérdida de Caja'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010312' AND Descripcion = 'Asig. Desgaste de Herramientas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010313' AND Descripcion = 'Viáticos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010314' AND Descripcion = 'Aporte Patronal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010315' AND Descripcion = 'Otro Conceptos No Imponibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010401' AND Descripcion = 'Gastos de Uniformes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010402' AND Descripcion = 'Gastos de Capacitación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010403' AND Descripcion = 'Becas de Estudio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010404' AND Descripcion = 'Implementos de Seguridad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010405' AND Descripcion = 'Gasto de Sala Cuna'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010406' AND Descripcion = 'Gastos de Importaciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010407' AND Descripcion = 'Honorarios Servicios de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010408' AND Descripcion = 'Gastos de Aseo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010409' AND Descripcion = 'Gastos de Mantención de Edificios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010410' AND Descripcion = 'Gastos de Remodelación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010411' AND Descripcion = 'Servicio de Mantención de Maquinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010412' AND Descripcion = 'Servicios Técnicos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010413' AND Descripcion = 'Gastos de Fotocopias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010414' AND Descripcion = 'Gastos de Imprenta - Formularios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010415' AND Descripcion = 'Gastos Implementos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010416' AND Descripcion = 'Gastos de Internet y Transmisión de Datos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010417' AND Descripcion = 'Fletes y Embalajes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010418' AND Descripcion = 'Seguros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010419' AND Descripcion = 'Arriendos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010420' AND Descripcion = 'Gastos de Electricidad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010421' AND Descripcion = 'Gastos de Calefacción - Gas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010422' AND Descripcion = 'Gastos de Agua Potable'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010423' AND Descripcion = 'Gastos de Teléfono'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010424' AND Descripcion = 'Gastos Básicos del Mes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010425' AND Descripcion = 'Gastos de Movilización'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010426' AND Descripcion = 'Gastos de Combustibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010427' AND Descripcion = 'Gastos de Viajes (Pasajes- Hotel - Auto )'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010428' AND Descripcion = 'Gastos de Coktail'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010429' AND Descripcion = 'Gastos de Cafetería y Similares'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010430' AND Descripcion = 'Gasto Patente Municipal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010431' AND Descripcion = 'Gasto de IVA  C.Fiscal No Utilizado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010432' AND Descripcion = 'Depreciacion'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010433' AND Descripcion = 'Otros Gastos Directos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010442' AND Descripcion = 'I.V.A. No Recuperable'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010501' AND Descripcion = 'Sueldo Base'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010502' AND Descripcion = 'Comisiones por Ventas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010503' AND Descripcion = 'Bonos de Producción'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010504' AND Descripcion = 'Horas Extraordinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010505' AND Descripcion = 'Bono de Responsabilidad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010506' AND Descripcion = 'Gratificación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010507' AND Descripcion = 'Aguinaldos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010508' AND Descripcion = 'Otros Bonos Imponibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010509' AND Descripcion = 'Asignación de Movilización'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010510' AND Descripcion = 'Asignación de Colación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010511' AND Descripcion = 'Asig de Pérdida de Caja'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010512' AND Descripcion = 'Asig. Desgaste de Herramientas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010513' AND Descripcion = 'Viáticos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010514' AND Descripcion = 'Aporte Patronal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010515' AND Descripcion = 'Otro Conceptos No Imponibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010601' AND Descripcion = 'Gastos de Uniformes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010602' AND Descripcion = 'Gastos de Capacitación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010603' AND Descripcion = 'Becas de Estudio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010604' AND Descripcion = 'Implementos de Seguridad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010605' AND Descripcion = 'Gasto de Sala Cuna'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010606' AND Descripcion = 'Gastos de Importaciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010607' AND Descripcion = 'Dieta Directorio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010608' AND Descripcion = 'Honorarios Profesionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010609' AND Descripcion = 'Honorarios Servicios de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010610' AND Descripcion = 'Gastos Legales Abogados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010611' AND Descripcion = 'Gastos Notariales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010612' AND Descripcion = 'Gastos de Suscripción'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010613' AND Descripcion = 'Gastos de Aseo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010614' AND Descripcion = 'Gastos de Mantención de Edificios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010615' AND Descripcion = 'Gastos de Remodelación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010616' AND Descripcion = 'Servicio de Mantención de Maquinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010617' AND Descripcion = 'Servicios Técnicos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010618' AND Descripcion = 'Artículos de Oficina'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010619' AND Descripcion = 'Gastos de Fotocopias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010620' AND Descripcion = 'Gastos de Imprenta - Formularios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010621' AND Descripcion = 'Gastos de Publicidad - Papelería'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010622' AND Descripcion = 'Gastos de Publicidad - Otros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010623' AND Descripcion = 'Gastos de Regalos a Clientes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010624' AND Descripcion = 'Gastos Implementos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010625' AND Descripcion = 'Gastos de Internet y Transmisión de Datos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010626' AND Descripcion = 'Fletes y Embalajes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010627' AND Descripcion = 'Seguros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010628' AND Descripcion = 'Arriendos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010629' AND Descripcion = 'Gastos de Electricidad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010630' AND Descripcion = 'Gastos de Calefacción - Gas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010631' AND Descripcion = 'Gastos de Agua Potable'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010632' AND Descripcion = 'Gastos de Teléfono'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010633' AND Descripcion = 'Gastos Básicos del Mes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010634' AND Descripcion = 'Gastos de Movilización'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010635' AND Descripcion = 'Gastos de Combustibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010636' AND Descripcion = 'Gastos de Viajes (Pasajes- Hotel - Auto )'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010637' AND Descripcion = 'Gastos de Coktail'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010638' AND Descripcion = 'Gastos de Cafetería y Similares'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010639' AND Descripcion = 'Gasto Patente Municipal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010640' AND Descripcion = 'Gasto de IVA  C.Fiscal No Utilizado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010641' AND Descripcion = 'Gastos de Leasing'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010642' AND Descripcion = 'Impuestos de Timbres - Letras'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010643' AND Descripcion = 'Gasto de Mercaderías Obsoletas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010644' AND Descripcion = 'Estimación de Deudores Incobrables'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010645' AND Descripcion = 'Otros Gastos de Administración'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030100' WHERE Codigo = '3010701' AND Descripcion = 'Depreciación del Ejercicio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3020100' WHERE Codigo = '3020101' AND Descripcion = 'Intereses Cobrados a Clientes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3020100' WHERE Codigo = '3020102' AND Descripcion = 'Intereses Percibidos por Prestamos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3020100' WHERE Codigo = '3020103' AND Descripcion = 'Descuento por Pronto Pago Obtenido'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3020100' WHERE Codigo = '3020104' AND Descripcion = 'Intereses por Fondos Mutuos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3020100' WHERE Codigo = '3020105' AND Descripcion = 'Intereses por Depósitos a Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3020300' WHERE Codigo = '3020106' AND Descripcion = 'Utilidad en Venta de Acciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3020200' WHERE Codigo = '3020201' AND Descripcion = 'Utilidad por Inversiones en Empresas Relacionadas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3050800' WHERE Codigo = '3020301' AND Descripcion = 'Venta de Activos Fijos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3020300' WHERE Codigo = '3020302' AND Descripcion = 'Arriendos Obtenidos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3020300' WHERE Codigo = '3020303' AND Descripcion = 'Castigos Recuperados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3020600' WHERE Codigo = '3020401' AND Descripcion = 'Pérdidas por Inversiones en Empresas Relacionadas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020501' AND Descripcion = 'Gastos Bancarios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020502' AND Descripcion = 'Intereses Bancarios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020503' AND Descripcion = 'Comisiones Bancarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020504' AND Descripcion = 'Gastos de Comercio Exterior'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020505' AND Descripcion = 'Gastos de Protesto Bancario'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020506' AND Descripcion = 'Comisiones Transbank'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020601' AND Descripcion = 'Intereses Factoring'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020602' AND Descripcion = 'Intereses por Leasing'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020603' AND Descripcion = 'Intereses Pagados a Proveedores'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020604' AND Descripcion = 'Intereses por Retenciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3050800' WHERE Codigo = '3020701' AND Descripcion = 'Costo de Venta por Enajenación de  Activos Fijos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020702' AND Descripcion = 'Gastos Rechazados Automóviles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020703' AND Descripcion = 'Gastos Rechazados  Contrib. Bs. Rs.'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051000' WHERE Codigo = '3020704' AND Descripcion = 'Otros Gastos Rechazados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051700' WHERE Codigo = '3020705' AND Descripcion = 'I.V.A. No Recuperable'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030300' WHERE Codigo = '3020801' AND Descripcion = 'Amortizaciones Varias del Ejercicio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051200' WHERE Codigo = '3021001' AND Descripcion = 'Diferencia de Cambio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3060100' WHERE Codigo = '3040101' AND Descripcion = 'Impuesto de Primera Categoría del Ejercicio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051700' WHERE Codigo = '3040102' AND Descripcion = 'Impuesto Adicional'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3060100' WHERE Codigo = '3040103' AND Descripcion = 'Otros Impuestos a la Renta'"
   Call ExecSQL(DbMain, QBase & Q1)

End Function

Public Function UpdateCtaPlanSII_IFRS(ByVal TblName As String)
   Dim Q1 As String
   Dim QBase As String
   
   QBase = "UPDATE " & TblName & " "

   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010101' AND Descripcion = 'Caja Principal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010102' AND Descripcion = 'Caja Chica - Fondo Fijo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010103' AND Descripcion = 'Depósitos en Tránsito $'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010104' AND Descripcion = 'Depósitos en Tránsito Us$'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010105' AND Descripcion = 'Depósitos en Tránsito Otras Monedas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010106' AND Descripcion = 'Banco $'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010107' AND Descripcion = 'Banco US$'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010100' WHERE Codigo = '1010108' AND Descripcion = 'Banco en Otras Monedas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010109' AND Descripcion = 'Depósitos a Plazo en $'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010110' AND Descripcion = 'Depósitos a Plazo en U.F.'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010111' AND Descripcion = 'Depósitos a Plazo en Us$'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010112' AND Descripcion = 'Depósitos a Plazo en Otras Monedas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010113' AND Descripcion = 'Fondos Mutuos en $'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010114' AND Descripcion = 'Fondos Mutuos en Us$'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010115' AND Descripcion = 'Fondos Mutuos en Otras Monedas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010201' AND Descripcion = 'Depósitos a Plazo en $'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010202' AND Descripcion = 'Depósitos a Plazo en U.F.'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010203' AND Descripcion = 'Depósitos a Plazo en Us$'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010204' AND Descripcion = 'Depósitos a Plazo en Otras Monedas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010205' AND Descripcion = 'Fondos Mutuos en $'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010206' AND Descripcion = 'Fondos Mutuos en Us$'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010300' WHERE Codigo = '1010207' AND Descripcion = 'Fondos Mutuos en Otras Monedas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010500' WHERE Codigo = '1010208' AND Descripcion = 'Acciones de S.A. en Bolsa de Valores'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010500' WHERE Codigo = '1010209' AND Descripcion = 'Acciones de S.A. No Bolsa de Valores'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010500' WHERE Codigo = '1010210' AND Descripcion = 'Mayor Valor Bursátil Acciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010500' WHERE Codigo = '1010211' AND Descripcion = 'Swaps de Moneda y de Tasa de Interés'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1010500' WHERE Codigo = '1010212' AND Descripcion = 'Contratos Forward de Tipo de Cambio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1033000' WHERE Codigo = '1010301' AND Descripcion = 'Seguros Pagados por Anticipado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1033200' WHERE Codigo = '1010302' AND Descripcion = 'Gasto Patente Municipal a Diferir Semestre'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1033200' WHERE Codigo = '1010303' AND Descripcion = 'Gasto de Publicidad Diferida'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1033000' WHERE Codigo = '1010304' AND Descripcion = 'Otros Gastos Pagados por Anticipado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012000' WHERE Codigo = '1010401' AND Descripcion = 'Clientes Ventas a Crédito'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012000' WHERE Codigo = '1010402' AND Descripcion = 'Clientes Ventas Al Contado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012000' WHERE Codigo = '1010403' AND Descripcion = 'Clientes Extranjeros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012000' WHERE Codigo = '1010404' AND Descripcion = 'Estimación Deudores Incobrables'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010405' AND Descripcion = 'Cheques a Fecha por Cobrar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010406' AND Descripcion = 'Letras por Cobrar en Cartera'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010407' AND Descripcion = 'Letras por Cobrar en Cob. Bancaria'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010408' AND Descripcion = 'Documentos Protestados en Cartera'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010409' AND Descripcion = 'Documentos en Cobranza Judicial'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010410' AND Descripcion = 'Tarjetas de Crédito por Cobrar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010411' AND Descripcion = 'Cheques Enviados a Factoring'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010412' AND Descripcion = 'Letras Enviados a Factoring'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010413' AND Descripcion = 'Estimac. Doctos. Incobrables Factoring'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1010414' AND Descripcion = 'Otros Documentos por Cobrar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1030500' WHERE Codigo = '1010415' AND Descripcion = 'Cuenta Corriente Socio A'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1030500' WHERE Codigo = '1010416' AND Descripcion = 'Cuenta Corriente Socio B'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010417' AND Descripcion = 'Cuenta Corriente Merc. Empleados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010418' AND Descripcion = 'Anticipos de Sueldo por Descontar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1032500' WHERE Codigo = '1010419' AND Descripcion = 'Anticipo a Proveedores'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010420' AND Descripcion = 'Anticipo de Honorarios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010421' AND Descripcion = 'Fondos por Rendir'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1032400' WHERE Codigo = '1010422' AND Descripcion = 'Anticipos de Gratificación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010423' AND Descripcion = 'Asig. Familiar por Recuperar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010424' AND Descripcion = 'Préstamos Socios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010425' AND Descripcion = 'Arriendos en Garantía por Recuperar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010426' AND Descripcion = 'Otros Deudores por Cobrar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1013000' WHERE Codigo = '1010427' AND Descripcion = 'Cuentas por Cobrar por Arrendamiento Financiero'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010601' AND Descripcion = 'Materias Primas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010602' AND Descripcion = 'Materiales Directos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010603' AND Descripcion = 'Productos en Proceso'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010604' AND Descripcion = 'Productos Terminados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010605' AND Descripcion = 'Mercaderías Nacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015500' WHERE Codigo = '1010606' AND Descripcion = 'Importaciones en Tránsito'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010607' AND Descripcion = 'Otras Existencias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015000' WHERE Codigo = '1010608' AND Descripcion = 'Provisión Obsoles. Mercaderías'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015100' WHERE Codigo = '1010701' AND Descripcion = 'Activos Biológicos a Valor Razonable'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015100' WHERE Codigo = '1010702' AND Descripcion = 'Otros Activos Biológicos Valor Razonable'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016200' WHERE Codigo = '1010801' AND Descripcion = 'Pagos Provisionales Mensuales (Ppm)'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1015900' WHERE Codigo = '1010802' AND Descripcion = 'I.V.A. Crédito Fiscal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016200' WHERE Codigo = '1010803' AND Descripcion = 'Crédito Sence'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016100' WHERE Codigo = '1010804' AND Descripcion = 'Crédito por Donaciones a Universidades'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016200' WHERE Codigo = '1010805' AND Descripcion = 'Crédito por Compras de Activo Fijo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016200' WHERE Codigo = '1010806' AND Descripcion = 'Crédito por Dev. Impto. Renta'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016000' WHERE Codigo = '1010807' AND Descripcion = 'Otros Impuestos por Recuperar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1016200' WHERE Codigo = '1010808' AND Descripcion = 'Otros Créditos por Recuperar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1033000' WHERE Codigo = '1020201' AND Descripcion = 'Seguros Pagados por Anticipado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1033000' WHERE Codigo = '1020202' AND Descripcion = 'Gasto de Publicidad Diferida'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1033000' WHERE Codigo = '1020203' AND Descripcion = 'Otros Gastos Pagados por Anticipado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012000' WHERE Codigo = '1020301' AND Descripcion = 'Deudores a Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1012500' WHERE Codigo = '1020302' AND Descripcion = 'Otras Cuentas por Cobrar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1022500' WHERE Codigo = '1020701' AND Descripcion = 'Software'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021100' WHERE Codigo = '1020901' AND Descripcion = 'Terreno 1'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021100' WHERE Codigo = '1020902' AND Descripcion = 'Terreno 2'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021000' WHERE Codigo = '1020903' AND Descripcion = 'Edificios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021000' WHERE Codigo = '1020904' AND Descripcion = 'Instalaciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021200' WHERE Codigo = '1020905' AND Descripcion = 'Otras Construcc. Y Obras de Infraestructura'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021200' WHERE Codigo = '1020906' AND Descripcion = 'Construcciones en Propiedades de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029900' WHERE Codigo = '1020907' AND Descripcion = 'Mejoras de Bienes Arrendados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021300' WHERE Codigo = '1020908' AND Descripcion = 'Maquinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021300' WHERE Codigo = '1020909' AND Descripcion = 'Equipos Industriales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021500' WHERE Codigo = '1020910' AND Descripcion = 'Equipos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021400' WHERE Codigo = '1020911' AND Descripcion = 'Muebles de Oficina'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1021700' WHERE Codigo = '1020912' AND Descripcion = 'Vehículos Aceptados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029900' WHERE Codigo = '1020913' AND Descripcion = 'Vehículos Rechazados (Automóviles)'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029900' WHERE Codigo = '1020914' AND Descripcion = 'Otras Maquinarias y Equipos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029900' WHERE Codigo = '1020915' AND Descripcion = 'Otros Activos Fijos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020916' AND Descripcion = 'Depreciación Acumulada Construcciones en Propiedades de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020917' AND Descripcion = 'Depreciación Acumulada de Instalaciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020918' AND Descripcion = 'Depreciación Acumulada de Otras Construcc. Y Obras de Infraestructura'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020919' AND Descripcion = 'Depreciación Acumulada de Construcciones en Propiedades de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020920' AND Descripcion = 'Depreciación Acumulada de Mejoras de Bienes Arrendados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020921' AND Descripcion = 'Depreciación Acumulada de Maquinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020922' AND Descripcion = 'Depreciación Acumulada de Equipos Industriales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020923' AND Descripcion = 'Depreciación Acumulada de Equipos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020924' AND Descripcion = 'Depreciación Acumulada de Muebles de Oficina'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029200' WHERE Codigo = '1020925' AND Descripcion = 'Depreciación Acumulada de Vehículos Aceptados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029200' WHERE Codigo = '1020926' AND Descripcion = 'Depreciación Acumulada de Vehículos Rechazados (Automóviles)'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020927' AND Descripcion = 'Depreciación Acumulada de Otras Maquinarias y Equipos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029000' WHERE Codigo = '1020928' AND Descripcion = 'Depreciación Acumulada de Otros Activos Fijos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020929' AND Descripcion = 'Deterioro Acumulado de Edificios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020930' AND Descripcion = 'Deterioro Acumulado de Construcciones en Propiedades de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020931' AND Descripcion = 'Deterioro Acumulado de Instalaciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020932' AND Descripcion = 'Deterioro Acumulado de Otras Construcciones y Obras de Infraestructura'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020933' AND Descripcion = 'Deterioro Acumulado de Construcciones en Propiedades de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020934' AND Descripcion = 'Deterioro Acumulado de Mejoras de Bienes Arrendados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020935' AND Descripcion = 'Deterioro Acumulado de Maquinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020936' AND Descripcion = 'Deterioro Acumulado de Equipos Industriales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020937' AND Descripcion = 'Deterioro Acumulado de Equipos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020938' AND Descripcion = 'Deterioro Acumulado de Muebles de Oficina'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020939' AND Descripcion = 'Deterioro Acumulado de Vehículos Aceptados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020940' AND Descripcion = 'Deterioro Acumulado de Vehículos Rechazados (Automóviles)'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020941' AND Descripcion = 'Deterioro Acumulado de Otras Maquinarias y Equipos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020942' AND Descripcion = 'Deterioro Acumulado de Otros Activos Fijos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3030200' WHERE Codigo = '1020943' AND Descripcion = 'Otros Deterioro Acumulado de Propiedades, Plantas y Equipo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029900' WHERE Codigo = '1021001' AND Descripcion = 'Activos Biológicos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1029900' WHERE Codigo = '1021002' AND Descripcion = 'Otros Activos Biologicos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1035000' WHERE Codigo = '1021301' AND Descripcion = 'Impuestos Diferidos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='1035000' WHERE Codigo = '1021302' AND Descripcion = 'Otros Impuestos Diferidos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2010100' WHERE Codigo = '2010102' AND Descripcion = 'Obligaciones Con Bancos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2010100' WHERE Codigo = '2010103' AND Descripcion = 'Obligaciones Con Bancos e Instit. Financ. L. P. - Porción C. Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2010100' WHERE Codigo = '2010104' AND Descripcion = 'Obligaciones a Largo Plazo Con Vencimiento Dentro de Un Año'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011100' WHERE Codigo = '2010201' AND Descripcion = 'Proveedores Nacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011100' WHERE Codigo = '2010202' AND Descripcion = 'Proveedores Extranjeros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011400' WHERE Codigo = '2010203' AND Descripcion = 'Cuentas por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011400' WHERE Codigo = '2010204' AND Descripcion = 'Honorarios por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011400' WHERE Codigo = '2010205' AND Descripcion = 'Cheques a Fecha por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011400' WHERE Codigo = '2010206' AND Descripcion = 'Letras por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011400' WHERE Codigo = '2010207' AND Descripcion = 'Facturas por Contabilizar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011400' WHERE Codigo = '2010208' AND Descripcion = 'Otras Cuentas por Pagar del Giro'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2017000' WHERE Codigo = '2010209' AND Descripcion = 'Anticipos de Clientes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015100' WHERE Codigo = '2010210' AND Descripcion = 'Provisiones Vacaciones del Personal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011000' WHERE Codigo = '2010211' AND Descripcion = 'Remuneraciones por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2010212' AND Descripcion = 'Afp por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2010213' AND Descripcion = 'Isapres por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2010214' AND Descripcion = 'Achs por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2010215' AND Descripcion = 'Ips por Pagar (Ex Inp)'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2010216' AND Descripcion = 'C.C.A.F. Por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2010217' AND Descripcion = 'Prestamos Ccaf por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2010218' AND Descripcion = 'Descuentos Conv. Empleados por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011000' WHERE Codigo = '2010219' AND Descripcion = 'Finiquitos por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2011000' WHERE Codigo = '2010220' AND Descripcion = 'Remuneraciones No Cobradas por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2010221' AND Descripcion = 'Otras Retenciones por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2010222' AND Descripcion = 'Impuesto Único Al Trabajo por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016100' WHERE Codigo = '2010223' AND Descripcion = 'Contribuciones de Bs. Rs. Por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015500' WHERE Codigo = '2010224' AND Descripcion = 'Seguro de Cesantía por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020600' WHERE Codigo = '2010401' AND Descripcion = 'Provisión por Garantía'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020600' WHERE Codigo = '2010402' AND Descripcion = 'Provisión de Reclamaciones Legales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020600' WHERE Codigo = '2010403' AND Descripcion = 'Participación en Utilidades y Bonos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020600' WHERE Codigo = '2010404' AND Descripcion = 'Otras Provisiones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015900' WHERE Codigo = '2010501' AND Descripcion = 'I.V.A. Débito Fiscal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016100' WHERE Codigo = '2010502' AND Descripcion = 'Impuesto Retenido por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016100' WHERE Codigo = '2010503' AND Descripcion = 'P.P.M. Por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016000' WHERE Codigo = '2010504' AND Descripcion = 'Impuestos a la Renta por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016100' WHERE Codigo = '2010505' AND Descripcion = 'Otros Impuestos por Pagar'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2016100' WHERE Codigo = '2010506' AND Descripcion = 'Impuesto Unico Al Trabajo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015000' WHERE Codigo = '2010601' AND Descripcion = 'Provisiones Indem. Años Servicio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015100' WHERE Codigo = '2010602' AND Descripcion = 'Provisión Vacaciones del Personal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015100' WHERE Codigo = '2010603' AND Descripcion = 'Provisión Honorarios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2015100' WHERE Codigo = '2010604' AND Descripcion = 'Provisión Remuneraciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2017000' WHERE Codigo = '2010701' AND Descripcion = 'Ingresos Percibidos por Adelantado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020100' WHERE Codigo = '2020101' AND Descripcion = 'Intereses Diferidos Leasing L.P.'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020100' WHERE Codigo = '2020102' AND Descripcion = 'Obligaciones Con Bancos e Inst. Financ. A Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020100' WHERE Codigo = '2020104' AND Descripcion = 'Obligaciones Garantizadas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020300' WHERE Codigo = '2020201' AND Descripcion = 'Proveedores de Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020300' WHERE Codigo = '2020202' AND Descripcion = 'Derechos de Aduana por Pagar a Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020300' WHERE Codigo = '2020203' AND Descripcion = 'Otros Acreedores a Largo Plazo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020300' WHERE Codigo = '2020204' AND Descripcion = 'Pasivos de Arrendamientos, No Corrientes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020600' WHERE Codigo = '2020401' AND Descripcion = 'Provisión por Garantía'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020600' WHERE Codigo = '2020402' AND Descripcion = 'Provisión de Reclamaciones Legales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020600' WHERE Codigo = '2020404' AND Descripcion = 'Otras Provisiones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2020600' WHERE Codigo = '2020701' AND Descripcion = 'Provisiones Indem. Años Servicio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030100' WHERE Codigo = '2030101' AND Descripcion = 'Capital Pagado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030200' WHERE Codigo = '2030502' AND Descripcion = 'Reserva Rev. De Capital'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030400' WHERE Codigo = '2030601' AND Descripcion = 'Reservas Varias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030400' WHERE Codigo = '2030602' AND Descripcion = 'Reservas Futuros Dividendos - Repartos Utilidades Sociales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030400' WHERE Codigo = '2030603' AND Descripcion = 'Reservas por Diferencias de Cambio por Conversión'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030400' WHERE Codigo = '2030604' AND Descripcion = 'Reservas de Coberturas de Flujo de Caja'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030400' WHERE Codigo = '2030605' AND Descripcion = 'Reservas de Ganancias y Pérdidas por Planes de Beneficios Definidos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='2030400' WHERE Codigo = '2030606' AND Descripcion = 'Otras Reservas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010100' WHERE Codigo = '3010101' AND Descripcion = 'Ingresos por Ventas de Productos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010100' WHERE Codigo = '3010102' AND Descripcion = 'Ingresos por Prestación de Servicios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010201' AND Descripcion = 'Costo de Venta de Productos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010202' AND Descripcion = 'Costo Directos de Prestación de Servicios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010203' AND Descripcion = 'Sueldo Base'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010204' AND Descripcion = 'Comisiones por Ventas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010205' AND Descripcion = 'Bonos de Producción'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010206' AND Descripcion = 'Horas Extraordinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010207' AND Descripcion = 'Bono de Responsabilidad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010208' AND Descripcion = 'Gratificación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010209' AND Descripcion = 'Aguinaldos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010210' AND Descripcion = 'Otros Bonos Imponibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010211' AND Descripcion = 'Asignación de Movilización'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010212' AND Descripcion = 'Asignación de Colación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010213' AND Descripcion = 'Asignacion de Pérdida de Caja'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010214' AND Descripcion = 'Asig. Desgaste de Herramientas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010215' AND Descripcion = 'Viáticos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010216' AND Descripcion = 'Aporte Patronal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010217' AND Descripcion = 'Otro Conceptos No Imponibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010218' AND Descripcion = 'Gastos de Uniformes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010219' AND Descripcion = 'Gastos de Capacitación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010220' AND Descripcion = 'Becas de Estudio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010221' AND Descripcion = 'Implementos de Seguridad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010222' AND Descripcion = 'Gasto de Sala Cuna'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010223' AND Descripcion = 'Gastos de Importaciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010224' AND Descripcion = 'Honorarios Servicios de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010225' AND Descripcion = 'Gastos de Aseo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010226' AND Descripcion = 'Gastos de Mantención de Edificios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010227' AND Descripcion = 'Gastos de Remodelación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010228' AND Descripcion = 'Servicio de Mantención de Maquinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010229' AND Descripcion = 'Servicios Técnicos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010230' AND Descripcion = 'Gastos de Fotocopias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010231' AND Descripcion = 'Gastos de Imprenta - Formularios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010232' AND Descripcion = 'Gastos Implementos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010233' AND Descripcion = 'Gastos de Internet y Transmisión de Datos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010234' AND Descripcion = 'Fletes y Embalajes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010235' AND Descripcion = 'Seguros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010236' AND Descripcion = 'Arriendos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010237' AND Descripcion = 'Gastos de Electricidad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010238' AND Descripcion = 'Gastos de Calefacción - Gas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010239' AND Descripcion = 'Gastos de Agua Potable'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010240' AND Descripcion = 'Gastos de Teléfono'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010241' AND Descripcion = 'Gastos Básicos del Mes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010242' AND Descripcion = 'Gastos de Movilización'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010243' AND Descripcion = 'Gastos de Combustibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010244' AND Descripcion = 'Gastos de Viajes (Pasajes- Hotel - Auto )'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010245' AND Descripcion = 'Gastos de Cocktail'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010246' AND Descripcion = 'Gastos de Cafetería y Similares'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010247' AND Descripcion = 'Gasto Patente Municipal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010248' AND Descripcion = 'Gasto de IVA Crédito Fiscal No Utilizado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010249' AND Descripcion = 'Depreciación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010250' AND Descripcion = 'Deterioro de Valor'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010251' AND Descripcion = 'Otros Gastos Directos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010200' WHERE Codigo = '3010252' AND Descripcion = 'I.V.A. No Recuperable'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010501' AND Descripcion = 'Sueldo Base'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010502' AND Descripcion = 'Comisiones por Ventas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010503' AND Descripcion = 'Bonos de Producción'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010504' AND Descripcion = 'Horas Extraordinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010505' AND Descripcion = 'Bono de Responsabilidad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010506' AND Descripcion = 'Gratificación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010507' AND Descripcion = 'Aguinaldos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010508' AND Descripcion = 'Otros Bonos Imponibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010509' AND Descripcion = 'Asignación de Movilización'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010510' AND Descripcion = 'Asignación de Colación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010511' AND Descripcion = 'Asig de Pérdida de Caja'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010512' AND Descripcion = 'Asig. Desgaste de Herramientas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010513' AND Descripcion = 'Viáticos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010514' AND Descripcion = 'Aporte Patronal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010515' AND Descripcion = 'Otro Conceptos No Imponibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010516' AND Descripcion = 'Gastos de Uniformes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010517' AND Descripcion = 'Gastos de Capacitación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010518' AND Descripcion = 'Becas de Estudio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010519' AND Descripcion = 'Implementos de Seguridad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010520' AND Descripcion = 'Gasto de Sala Cuna'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010521' AND Descripcion = 'Gastos de Importaciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010522' AND Descripcion = 'Dieta Directorio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010523' AND Descripcion = 'Honorarios Profesionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010524' AND Descripcion = 'Honorarios Servicios de Terceros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010525' AND Descripcion = 'Gastos Legales Abogados'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010526' AND Descripcion = 'Gastos Notariales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010527' AND Descripcion = 'Gastos de Suscripción'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010528' AND Descripcion = 'Gastos de Aseo'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010529' AND Descripcion = 'Gastos de Mantención de Edificios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010530' AND Descripcion = 'Gastos de Remodelación'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010531' AND Descripcion = 'Servicio de Mantención de Maquinarias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010532' AND Descripcion = 'Servicios Técnicos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010533' AND Descripcion = 'Artículos de Oficina'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010534' AND Descripcion = 'Gastos de Fotocopias'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010535' AND Descripcion = 'Gastos de Imprenta - Formularios'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010536' AND Descripcion = 'Gastos de Publicidad - Papelería'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010537' AND Descripcion = 'Gastos de Publicidad - Otros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010538' AND Descripcion = 'Gastos de Regalos a Clientes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010539' AND Descripcion = 'Gastos Implementos Computacionales'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010540' AND Descripcion = 'Gastos de Internet y Transmisión de Datos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010541' AND Descripcion = 'Fletes y Embalajes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010542' AND Descripcion = 'Seguros'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010543' AND Descripcion = 'Arriendos'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010544' AND Descripcion = 'Gastos de Electricidad'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010545' AND Descripcion = 'Gastos de Calefacción - Gas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010546' AND Descripcion = 'Gastos de Agua Potable'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010547' AND Descripcion = 'Gastos de Teléfono'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010548' AND Descripcion = 'Gastos Básicos del Mes'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010549' AND Descripcion = 'Gastos de Movilización'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010550' AND Descripcion = 'Gastos de Combustibles'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010551' AND Descripcion = 'Gastos de Viajes (Pasajes- Hotel - Auto )'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010552' AND Descripcion = 'Gastos de Cocktail'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010553' AND Descripcion = 'Gastos de Cafetería y Similares'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010554' AND Descripcion = 'Gasto Patente Municipal'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010555' AND Descripcion = 'Gasto de IVA C.Fiscal No Utilizado'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010556' AND Descripcion = 'Gastos de Leasing'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010557' AND Descripcion = 'Impuestos de Timbres - Letras'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010558' AND Descripcion = 'Gasto de Mercaderías Obsoletas'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010559' AND Descripcion = 'Estimación de Deudores Incobrables'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010560' AND Descripcion = 'Otros Gastos de Administración'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010561' AND Descripcion = 'Depreciación del Ejercicio'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3010300' WHERE Codigo = '3010562' AND Descripcion = 'Deterioro de Valor'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051500' WHERE Codigo = '3011107' AND Descripcion = 'Intereses Factoring'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051500' WHERE Codigo = '3011108' AND Descripcion = 'Intereses por Leasing'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051500' WHERE Codigo = '3011109' AND Descripcion = 'Intereses Pagados a Proveedores'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051500' WHERE Codigo = '3011110' AND Descripcion = 'Intereses por Retenciones'"
   Call ExecSQL(DbMain, QBase & Q1)
   Q1 = "SET CodCtaPlanSII ='3051200' WHERE Codigo = '3011301' AND Descripcion = 'Diferencias de Cambio'"
   Call ExecSQL(DbMain, QBase & Q1)

End Function

Private Sub UpdateCodActiv2018()
   Dim Q1 As String
   
   'eliminamos los registros de la versión 1 y 2 ya que ahora no se usan
   Q1 = "DELETE * FROM CodActiv WHERE Version < 3"
   Call ExecSQL(DbMain, Q1)
   
   'insertamos los nuevos registros con versión 3
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011101', 'Cultivo de trigo', 3, '011111')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011102', 'Cultivo de maíz', 3, '011112')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011103', 'Cultivo de avena', 3, '011113')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011200', 'Cultivo de arroz', 3, '011114')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011104', 'Cultivo de cebada', 3, '011115')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011105', 'Cultivo de otros cereales (excepto trigo, maíz, avena y cebada)', 3, '011119')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011902', 'Cultivos forrajeros en praderas mejoradas o sembradas; cultivos suplementarios forrajeros', 3, '011121')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011902', 'Cultivos forrajeros en praderas mejoradas o sembradas; cultivos suplementarios forrajeros', 3, '011122')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011106', 'Cultivo de porotos', 3, '011131')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011107', 'Cultivo de lupino', 3, '011132')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011108', 'Cultivo de otras legumbres (excepto porotos y lupino)', 3, '011139')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011301', 'Cultivo de papas', 3, '011141')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011302', 'Cultivo de camotes', 3, '011142')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011303', 'Cultivo de otros tubérculos (excepto papas y camotes)', 3, '011149')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011109', 'Cultivo de semillas de raps', 3, '011151')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011110', 'Cultivo de semillas de maravilla (girasol)', 3, '011152')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012600', 'Cultivo de frutos oleaginosos (incluye el cultivo de aceitunas)', 3, '011159')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011111', 'Cultivo de semillas de cereales, legumbres y oleaginosas (excepto semillas de raps y maravilla)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011111', 'Cultivo de semillas de cereales, legumbres y oleaginosas (excepto semillas de raps y maravilla)', 3, '011160')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011304', 'Cultivo de remolacha azucarera', 3, '011191')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011500', 'Cultivo de tabaco', 3, '011192')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016300', 'Actividades poscosecha', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011600', 'Cultivo de plantas de fibra', 3, '011193')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012900', 'Cultivo de otras plantas perennes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012802', 'Cultivo de plantas aromáticas, medicinales y farmacéuticas', 3, '011194')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012900', 'Cultivo de otras plantas perennes', 3, '011199')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011400', 'Cultivo de caña de azúcar', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012802', 'Cultivo de plantas aromáticas, medicinales y farmacéuticas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011306', 'Cultivo de hortalizas y melones', 3, '011211')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012600', 'Cultivo de frutos oleaginosos (incluye el cultivo de aceitunas)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012802', 'Cultivo de plantas aromáticas, medicinales y farmacéuticas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011306', 'Cultivo de hortalizas y melones', 3, '011212')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011306', 'Cultivo de hortalizas y melones', 3, '011213')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011901', 'Cultivo de flores', 3, '011220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('013000', 'Cultivo de plantas vivas incluida la producción en viveros (excepto viveros forestales)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011305', 'Cultivo de semillas de hortalizas', 3, '011230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011903', 'Cultivos de semillas de flores; cultivo de semillas de plantas forrajeras', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012501', 'Cultivo de semillas de frutas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('013000', 'Cultivo de plantas vivas incluida la producción en viveros (excepto viveros forestales)', 3, '011240')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011306', 'Cultivo de hortalizas y melones', 3, '011250')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012900', 'Cultivo de otras plantas perennes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('023000', 'Recolección de productos forestales distintos de la madera', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012111', 'Cultivo de uva destinada a la producción de pisco y aguardiente', 3, '011311')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012112', 'Cultivo de uva destinada a la producción de vino', 3, '011312')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110200', 'Elaboración de vinos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012120', 'Cultivo de uva para mesa', 3, '011313')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012400', 'Cultivo de frutas de pepita y de hueso', 3, '011321')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012200', 'Cultivo de frutas tropicales y subtropicales (incluye el cultivo de paltas)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012300', 'Cultivo de cítricos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012502', 'Cultivo de otros frutos y nueces de árboles y arbustos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012600', 'Cultivo de frutos oleaginosos (incluye el cultivo de aceitunas)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012502', 'Cultivo de otros frutos y nueces de árboles y arbustos', 3, '011322')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('023000', 'Recolección de productos forestales distintos de la madera', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012700', 'Cultivo de plantas con las que se preparan bebidas (incluye el cultivo de café, té y mate)', 3, '011330')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012801', 'Cultivo de especias', 3, '011340')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014101', 'Cría de ganado bovino para la producción lechera', 3, '012111')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014102', 'Cría de ganado bovino para la producción de carne o como ganado reproductor', 3, '012112')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014410', 'Cría de ovejas (ovinos)', 3, '012120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014200', 'Cría de caballos y otros equinos', 3, '012130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014500', 'Cría de cerdos', 3, '012210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014601', 'Cría de aves de corral para la producción de carne', 3, '012221')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014602', 'Cría de aves de corral para la producción de huevos', 3, '012222')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014909', 'Cría de otros animales n.c.p.', 3, '012223')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014909', 'Cría de otros animales n.c.p.', 3, '012230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014901', 'Apicultura', 3, '012240')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014909', 'Cría de otros animales n.c.p.', 3, '012250')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032200', 'Acuicultura de agua dulce', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014909', 'Cría de otros animales n.c.p.', 3, '012290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014300', 'Cría de llamas, alpacas, vicuñas, guanacos y otros camélidos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014420', 'Cría de cabras (caprinos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032130', 'Reproducción y cría de moluscos, crustáceos y gusanos marinos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032200', 'Acuicultura de agua dulce', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('015000', 'Cultivo de productos agrícolas en combinación con la cría de animales (explotación mixta)', 3, '013000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016300', 'Actividades poscosecha', 3, '014011')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016300', 'Actividades poscosecha', 3, '014012')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016400', 'Tratamiento de semillas para propagación', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016100', 'Actividades de apoyo a la agricultura', 3, '014013')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016100', 'Actividades de apoyo a la agricultura', 3, '014014')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016100', 'Actividades de apoyo a la agricultura', 3, '014015')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016100', 'Actividades de apoyo a la agricultura', 3, '014019')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('813000', 'Actividades de paisajismo, servicios de jardinería y servicios conexos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960901', 'Servicios de adiestramiento, guardería, peluquería, paseo de mascotas (excepto act. veterinarias)', 3, '014021')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016200', 'Actividades de apoyo a la ganadería', 3, '014022')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('017000', 'Caza ordinaria y mediante trampas y actividades de servicios conexas', 3, '015010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('017000', 'Caza ordinaria y mediante trampas y actividades de servicios conexas', 3, '015090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949909', 'Actividades de otras asociaciones n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('022000', 'Extracción de madera', 3, '020010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('021002', 'Silvicultura y otras actividades forestales (excepto explotación de viveros forestales)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('023000', 'Recolección de productos forestales distintos de la madera', 3, '020020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('021001', 'Explotación de viveros forestales', 3, '020030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012900', 'Cultivo de otras plantas perennes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('024001', 'Servicios de forestación a cambio de una retribución o por contrata', 3, '020041')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('024002', 'Servicios de corta de madera a cambio de una retribución o por contrata', 3, '020042')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('024003', 'Servicios de extinción y prevención de incendios forestales', 3, '020043')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('024009', 'Otros servicios de apoyo a la silvicultura n.c.p.', 3, '020049')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032200', 'Acuicultura de agua dulce', 3, '051010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032110', 'Cultivo y crianza de peces marinos', 3, '051020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032120', 'Cultivo, reproducción y manejo de algas marinas', 3, '051030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032130', 'Reproducción y cría de moluscos, crustáceos y gusanos marinos', 3, '051040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032140', 'Servicios relacionados con la acuicultura marina', 3, '051090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032200', 'Acuicultura de agua dulce', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031110', 'Pesca marítima industrial, excepto de barcos factoría', 3, '052010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102050', 'Actividades de elaboración y conservación de pescado, realizadas en barcos factoría', 3, '052020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031120', 'Pesca marítima artesanal', 3, '052030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031200', 'Pesca de agua dulce', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031130', 'Recolección y extracción de productos marinos', 3, '052040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031140', 'Servicios relacionados con la pesca marítima', 3, '052050')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031200', 'Pesca de agua dulce', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('051000', 'Extracción de carbón de piedra', 3, '100000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('052000', 'Extracción de lignito', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089200', 'Extracción de turba', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('192000', 'Fabricación de productos de la refinación del petróleo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('061000', 'Extracción de petróleo crudo', 3, '111000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('062000', 'Extracción de gas natural', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('091001', 'Actividades de apoyo para la extracción de petróleo y gas natural prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('091001', 'Actividades de apoyo para la extracción de petróleo y gas natural prestados por empresas', 3, '112000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('072100', 'Extracción de minerales de uranio y torio', 3, '120000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('071000', 'Extracción de minerales de hierro', 3, '131000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('072910', 'Extracción de oro y plata', 3, '132010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('072991', 'Extracción de zinc y plomo', 3, '132020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('072992', 'Extracción de manganeso', 3, '132030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('072999', 'Extracción de otros minerales metalíferos no ferrosos n.c.p. (excepto zinc, plomo y manganeso)', 3, '132090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('040000', 'Extracción y procesamiento de cobre', 3, '133000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('081000', 'Extracción de piedra, arena y arcilla', 3, '141000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089190', 'Extracción de minerales para la fabricación de abonos y productos químicos n.c.p.', 3, '142100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089300', 'Extracción de sal', 3, '142200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089110', 'Extracción y procesamiento de litio', 3, '142300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089190', 'Extracción de minerales para la fabricación de abonos y productos químicos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089900', 'Explotación de otras minas y canteras n.c.p.', 3, '142900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('081000', 'Extracción de piedra, arena y arcilla', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('101011', 'Explotación de mataderos de bovinos, ovinos, equinos, caprinos, porcinos y camélidos', 3, '151110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('521001', 'Explotación de frigoríficos para almacenamiento y depósito', 3, '151120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('101019', 'Explotación de mataderos de aves y de otros tipos de animales n.c.p.', 3, '151130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('101020', 'Elaboración y conservación de carne y productos cárnicos', 3, '151140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102010', 'Producción de harina de pescado', 3, '151210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102030', 'Elaboración y conservación de otros pescados, en plantas en tierra (excepto barcos factoría)', 3, '151221')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102020', 'Elaboración y conservación de salmónidos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102040', 'Elaboración y conservación de crustáceos, moluscos y otros productos acuáticos, en plantas en tierra', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107500', 'Elaboración de comidas y platos preparados envasados, rotulados y con información nutricional', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102020', 'Elaboración y conservación de salmónidos', 3, '151222')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102030', 'Elaboración y conservación de otros pescados, en plantas en tierra (excepto barcos factoría)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102040', 'Elaboración y conservación de crustáceos, moluscos y otros productos acuáticos, en plantas en tierra', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107500', 'Elaboración de comidas y platos preparados envasados, rotulados y con información nutricional', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102020', 'Elaboración y conservación de salmónidos', 3, '151223')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102030', 'Elaboración y conservación de otros pescados, en plantas en tierra (excepto barcos factoría)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102040', 'Elaboración y conservación de crustáceos, moluscos y otros productos acuáticos, en plantas en tierra', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102060', 'Elaboración y procesamiento de algas', 3, '151230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('103000', 'Elaboración y conservación de frutas, legumbres y hortalizas', 3, '151300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107500', 'Elaboración de comidas y platos preparados envasados, rotulados y con información nutricional', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('104000', 'Elaboración de aceites y grasas de origen vegetal y animal (excepto elaboración de mantequilla)', 3, '151410')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('104000', 'Elaboración de aceites y grasas de origen vegetal y animal (excepto elaboración de mantequilla)', 3, '151420')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('104000', 'Elaboración de aceites y grasas de origen vegetal y animal (excepto elaboración de mantequilla)', 3, '151430')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('105000', 'Elaboración de productos lácteos', 3, '152010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('105000', 'Elaboración de productos lácteos', 3, '152020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('105000', 'Elaboración de productos lácteos', 3, '152030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('106101', 'Molienda de trigo: producción de harina, sémola y gránulos', 3, '153110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('106102', 'Molienda de arroz; producción de harina de arroz', 3, '153120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('106109', 'Elaboración de otros productos de molinería n.c.p.', 3, '153190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('106200', 'Elaboración de almidones y productos derivados del almidón', 3, '153210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('106200', 'Elaboración de almidones y productos derivados del almidón', 3, '153220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('108000', 'Elaboración de piensos preparados para animales', 3, '153300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107100', 'Elaboración de productos de panadería y pastelería', 3, '154110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107100', 'Elaboración de productos de panadería y pastelería', 3, '154120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107200', 'Elaboración de azúcar', 3, '154200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107300', 'Elaboración de cacao, chocolate y de productos de confitería', 3, '154310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107300', 'Elaboración de cacao, chocolate y de productos de confitería', 3, '154320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107400', 'Elaboración de macarrones, fideos, alcuzcuz y productos farináceos similares', 3, '154400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107500', 'Elaboración de comidas y platos preparados envasados, rotulados y con información nutricional', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107901', 'Elaboración de té, café, mate e infusiones de hierbas', 3, '154910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107902', 'Elaboración de levaduras naturales o artificiales', 3, '154920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107903', 'Elaboración de vinagres, mostazas, mayonesas y condimentos en general', 3, '154930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107909', 'Elaboración de otros productos alimenticios n.c.p.', 3, '154990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('103000', 'Elaboración y conservación de frutas, legumbres y hortalizas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107100', 'Elaboración de productos de panadería y pastelería', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107500', 'Elaboración de comidas y platos preparados envasados, rotulados y con información nutricional', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110110', 'Elaboración de pisco (industrias pisqueras)', 3, '155110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110120', 'Destilación, rectificación y mezclas de bebidas alcohólicas; excepto pisco', 3, '155120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201109', 'Fabricación de otras sustancias químicas básicas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110200', 'Elaboración de vinos', 3, '155200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110300', 'Elaboración de bebidas malteadas y de malta', 3, '155300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110401', 'Elaboración de bebidas no alcohólicas', 3, '155410')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110402', 'Producción de aguas minerales y otras aguas embotelladas', 3, '155420')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('353002', 'Elaboración de hielo (excepto fabricación de hielo seco)', 3, '155430')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('120001', 'Elaboración de cigarros y cigarrillos', 3, '160010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('120009', 'Elaboración de otros productos de tabaco n.c.p.', 3, '160090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131200', 'Tejedura de productos textiles', 3, '171100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131100', 'Preparación e hilatura de fibras textiles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131300', 'Acabado de productos textiles', 3, '171200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952900', 'Reparación de otros efectos personales y enseres domésticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139200', 'Fabricación de artículos confeccionados de materiales textiles, excepto prendas de vestir', 3, '172100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139300', 'Fabricación de tapices y alfombras', 3, '172200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139400', 'Fabricación de cuerdas, cordeles, bramantes y redes', 3, '172300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139900', 'Fabricación de otros productos textiles n.c.p.', 3, '172910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131300', 'Acabado de productos textiles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139900', 'Fabricación de otros productos textiles n.c.p.', 3, '172990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170900', 'Fabricación de otros artículos de papel y cartón', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('143000', 'Fabricación de artículos de punto y ganchillo', 3, '173000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139100', 'Fabricación de tejidos de punto y ganchillo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('141001', 'Fabricación de prendas de vestir de materiales textiles y similares', 3, '181010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131300', 'Acabado de productos textiles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('141002', 'Fabricación de prendas de vestir de cuero natural o artificial', 3, '181020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('141003', 'Fabricación de accesorios de vestir', 3, '181030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('141004', 'Fabricación de ropa de trabajo', 3, '181040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131300', 'Acabado de productos textiles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('151100', 'Curtido y adobo de cueros; adobo y teñido de pieles', 3, '182000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('142000', 'Fabricación de artículos de piel', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('151100', 'Curtido y adobo de cueros; adobo y teñido de pieles', 3, '191100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('151200', 'Fabricación de maletas, bolsos y artículos similares, artículos de talabartería y guarnicionería', 3, '191200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('152000', 'Fabricación de calzado', 3, '192000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('151200', 'Fabricación de maletas, bolsos y artículos similares, artículos de talabartería y guarnicionería', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162900', 'Fabricación de otros productos de madera, de artículos de corcho, paja y materiales trenzables', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221900', 'Fabricación de otros productos de caucho', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('323000', 'Fabricación de artículos de deporte', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('161000', 'Aserrado y acepilladura de madera', 3, '201000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162100', 'Fabricación de hojas de madera para enchapado y tableros a base de madera', 3, '202100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162200', 'Fabricación de partes y piezas de carpintería para edificios y construcciones', 3, '202200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162300', 'Fabricación de recipientes de madera', 3, '202300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162900', 'Fabricación de otros productos de madera, de artículos de corcho, paja y materiales trenzables', 3, '202900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170110', 'Fabricación de celulosa y otras pastas de madera', 3, '210110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170190', 'Fabricación de papel y cartón para su posterior uso industrial n.c.p.', 3, '210121')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170190', 'Fabricación de papel y cartón para su posterior uso industrial n.c.p.', 3, '210129')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170200', 'Fabricación de papel y cartón ondulado y de envases de papel y cartón', 3, '210200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170900', 'Fabricación de otros artículos de papel y cartón', 3, '210900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('181109', 'Otras actividades de impresión n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581100', 'Edición de libros', 3, '221101')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581200', 'Edición de directorios y listas de correo', 3, '221109')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581300', 'Edición de diarios, revistas y otras publicaciones periódicas', 3, '221200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '221300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581900', 'Otras actividades de edición', 3, '221900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581300', 'Edición de diarios, revistas y otras publicaciones periódicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('181101', 'Impresión de libros', 3, '222101')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('181109', 'Otras actividades de impresión n.c.p.', 3, '222109')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170900', 'Fabricación de otros artículos de papel y cartón', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('181200', 'Actividades de servicios relacionadas con la impresión', 3, '222200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('182000', 'Reproducción de grabaciones', 3, '223000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('191000', 'Fabricación de productos de hornos de coque', 3, '231000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('192000', 'Fabricación de productos de la refinación del petróleo', 3, '232000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201109', 'Fabricación de otras sustancias químicas básicas n.c.p.', 3, '233000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('210000', 'Fabricación de productos farmacéuticos, sustancias químicas medicinales y productos botánicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('242009', 'Fabricación de productos primarios de metales preciosos y de otros metales no ferrosos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('381200', 'Recogida de desechos peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382200', 'Tratamiento y eliminación de desechos peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201101', 'Fabricación de carbón vegetal (excepto activado); fabricación de briquetas de carbón vegetal', 3, '241110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201109', 'Fabricación de otras sustancias químicas básicas n.c.p.', 3, '241190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('191000', 'Fabricación de productos de hornos de coque', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201200', 'Fabricación de abonos y compuestos de nitrógeno', 3, '241200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382100', 'Tratamiento y eliminación de desechos no peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201300', 'Fabricación de plásticos y caucho sintético en formas primarias', 3, '241300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202100', 'Fabricación de plaguicidas y otros productos químicos de uso agropecuario', 3, '242100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202200', 'Fabricación de pinturas, barnices y productos de revestimiento, tintas de imprenta y masillas', 3, '242200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('210000', 'Fabricación de productos farmacéuticos, sustancias químicas medicinales y productos botánicos', 3, '242300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202300', 'Fabricación de jabones y detergentes, preparados para limpiar, perfumes y preparados de tocador', 3, '242400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202901', 'Fabricación de explosivos y productos pirotécnicos', 3, '242910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202909', 'Fabricación de otros productos químicos n.c.p.', 3, '242990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107909', 'Elaboración de otros productos alimenticios n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201109', 'Fabricación de otras sustancias químicas básicas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('268000', 'Fabricación de soportes magnéticos y ópticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281700', 'Fabricación de maquinaria y equipo de oficina (excepto computadores y equipo periférico)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('203000', 'Fabricación de fibras artificiales', 3, '243000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221100', 'Fabricación de cubiertas y cámaras de caucho; recauchutado y renovación de cubiertas de caucho', 3, '251110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221100', 'Fabricación de cubiertas y cámaras de caucho; recauchutado y renovación de cubiertas de caucho', 3, '251120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221900', 'Fabricación de otros productos de caucho', 3, '251900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281200', 'Fabricación de equipo de propulsión de fluidos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '252010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '252020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '252090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273300', 'Fabricación de dispositivos de cableado', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('231001', 'Fabricación de vidrio plano', 3, '261010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('231002', 'Fabricación de vidrio hueco', 3, '261020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('231003', 'Fabricación de fibras de vidrio', 3, '261030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('231009', 'Fabricación de productos de vidrio n.c.p.', 3, '261090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239300', 'Fabricación de otros productos de porcelana y de cerámica', 3, '269101')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239300', 'Fabricación de otros productos de porcelana y de cerámica', 3, '269109')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239200', 'Fabricación de materiales de construcción de arcilla', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239100', 'Fabricación de productos refractarios', 3, '269200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239200', 'Fabricación de materiales de construcción de arcilla', 3, '269300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239400', 'Fabricación de cemento, cal y yeso', 3, '269400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239500', 'Fabricación de artículos de hormigón, cemento y yeso', 3, '269510')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239500', 'Fabricación de artículos de hormigón, cemento y yeso', 3, '269520')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239500', 'Fabricación de artículos de hormigón, cemento y yeso', 3, '269530')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239500', 'Fabricación de artículos de hormigón, cemento y yeso', 3, '269590')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239600', 'Corte, talla y acabado de la piedra', 3, '269600')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239900', 'Fabricación de otros productos minerales no metálicos n.c.p.', 3, '269910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239900', 'Fabricación de otros productos minerales no metálicos n.c.p.', 3, '269990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131200', 'Tejedura de productos textiles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('241000', 'Industrias básicas de hierro y acero', 3, '271000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('243100', 'Fundición de hierro y acero', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('242001', 'Fabricación de productos primarios de cobre', 3, '272010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('242002', 'Fabricación de productos primarios de aluminio', 3, '272020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('242009', 'Fabricación de productos primarios de metales preciosos y de otros metales no ferrosos n.c.p.', 3, '272090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('243100', 'Fundición de hierro y acero', 3, '273100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('243200', 'Fundición de metales no ferrosos', 3, '273200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('251100', 'Fabricación de productos metálicos para uso estructural', 3, '281100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('251201', 'Fabricación de recipientes de metal para gases comprimidos o licuados', 3, '281211')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('251209', 'Fabricación de tanques, depósitos y recipientes de metal n.c.p.', 3, '281219')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '281280')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('251300', 'Fabricación de generadores de vapor, excepto calderas de agua caliente para calefacción central', 3, '281310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '281380')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259100', 'Forja, prensado, estampado y laminado de metales; pulvimetalurgia', 3, '289100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259200', 'Tratamiento y revestimiento de metales; maquinado', 3, '289200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016200', 'Actividades de apoyo a la ganadería', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('181109', 'Otras actividades de impresión n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952900', 'Reparación de otros efectos personales y enseres domésticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259300', 'Fabricación de artículos de cuchillería, herramientas de mano y artículos de ferretería', 3, '289310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259300', 'Fabricación de artículos de cuchillería, herramientas de mano y artículos de ferretería', 3, '289320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259900', 'Fabricación de otros productos elaborados de metal n.c.p.', 3, '289910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259900', 'Fabricación de otros productos elaborados de metal n.c.p.', 3, '289990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281700', 'Fabricación de maquinaria y equipo de oficina (excepto computadores y equipo periférico)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '291110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291180')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281300', 'Fabricación de otras bombas, compresores, grifos y válvulas', 3, '291210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281200', 'Fabricación de equipo de propulsión de fluidos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291280')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281400', 'Fabricación de cojinetes, engranajes, trenes de engranajes y piezas de transmisión', 3, '291310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281200', 'Fabricación de equipo de propulsión de fluidos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291380')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281500', 'Fabricación de hornos, calderas y quemadores', 3, '291410')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291480')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281600', 'Fabricación de equipo de elevación y manipulación', 3, '291510')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291580')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281900', 'Fabricación de otros tipos de maquinaria de uso general', 3, '291910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291980')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282100', 'Fabricación de maquinaria agropecuaria y forestal', 3, '292110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331201', 'Reparación de maquinaria agropecuaria y forestal', 3, '292180')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282200', 'Fabricación de maquinaria para la conformación de metales y de máquinas herramienta', 3, '292210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281800', 'Fabricación de herramientas de mano motorizadas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281900', 'Fabricación de otros tipos de maquinaria de uso general', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '292280')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282300', 'Fabricación de maquinaria metalúrgica', 3, '292310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331202', 'Reparación de maquinaria metalúrgica, para la minería, extracción de petróleo y para la construcción', 3, '292380')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282400', 'Fabricación de maquinaria para la explotación de minas y canteras y para obras de construcción', 3, '292411')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282400', 'Fabricación de maquinaria para la explotación de minas y canteras y para obras de construcción', 3, '292412')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331202', 'Reparación de maquinaria metalúrgica, para la minería, extracción de petróleo y para la construcción', 3, '292480')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282500', 'Fabricación de maquinaria para la elaboración de alimentos, bebidas y tabaco', 3, '292510')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331203', 'Reparación de maquinaria para la elaboración de alimentos, bebidas y tabaco', 3, '292580')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282600', 'Fabricación de maquinaria para la elaboración de productos textiles, prendas de vestir y cueros', 3, '292610')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331204', 'Reparación de maquinaria para producir textiles, prendas de vestir, artículos de cuero y calzado', 3, '292680')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('252000', 'Fabricación de armas y municiones', 3, '292710')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('303000', 'Fabricación de aeronaves, naves espaciales y maquinaria conexa', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('304000', 'Fabricación de vehículos militares de combate', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '292780')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '292910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259300', 'Fabricación de artículos de cuchillería, herramientas de mano y artículos de ferretería', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282600', 'Fabricación de maquinaria para la elaboración de productos textiles, prendas de vestir y cueros', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '292980')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('275000', 'Fabricación de aparatos de uso doméstico', 3, '293000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281500', 'Fabricación de hornos, calderas y quemadores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281900', 'Fabricación de otros tipos de maquinaria de uso general', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('262000', 'Fabricación de computadores y equipo periférico', 3, '300010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281700', 'Fabricación de maquinaria y equipo de oficina (excepto computadores y equipo periférico)', 3, '300020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('271000', 'Fabricación de motores, generadores y transformadores eléctricos, aparatos de distribución y control', 3, '311010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '311080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('271000', 'Fabricación de motores, generadores y transformadores eléctricos, aparatos de distribución y control', 3, '312010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273300', 'Fabricación de dispositivos de cableado', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '312080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273200', 'Fabricación de otros hilos y cables eléctricos', 3, '313000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273100', 'Fabricación de cables de fibra óptica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('272000', 'Fabricación de pilas, baterías y acumuladores', 3, '314000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('274000', 'Fabricación de equipo eléctrico de iluminación', 3, '315010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '315080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '319010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259900', 'Fabricación de otros productos elaborados de metal n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('263000', 'Fabricación de equipo de comunicaciones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('265100', 'Fabricación de equipo de medición, prueba, navegación y control', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273300', 'Fabricación de dispositivos de cableado', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('274000', 'Fabricación de equipo eléctrico de iluminación', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282200', 'Fabricación de maquinaria para la conformación de metales y de máquinas herramienta', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('293000', 'Fabricación de partes, piezas y accesorios para vehículos automotores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('302000', 'Fabricación de locomotoras y material rodante', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '319080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '321010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '321080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('263000', 'Fabricación de equipo de comunicaciones', 3, '322010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('265100', 'Fabricación de equipo de medición, prueba, navegación y control', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('951200', 'Reparación de equipo de comunicaciones (incluye la reparación teléfonos celulares)', 3, '322080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('264000', 'Fabricación de aparatos electrónicos de consumo', 3, '323000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('263000', 'Fabricación de equipo de comunicaciones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('267000', 'Fabricación de instrumentos ópticos y equipo fotográfico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281700', 'Fabricación de maquinaria y equipo de oficina (excepto computadores y equipo periférico)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952100', 'Reparación de aparatos electrónicos de consumo (incluye aparatos de televisión y radio)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '331110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('266000', 'Fabricación de equipo de irradiación y equipo electrónico de uso médico y terapéutico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325001', 'Actividades de laboratorios dentales', 3, '331120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '331180')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('265100', 'Fabricación de equipo de medición, prueba, navegación y control', 3, '331210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('267000', 'Fabricación de instrumentos ópticos y equipo fotográfico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281900', 'Fabricación de otros tipos de maquinaria de uso general', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331301', 'Reparación de equipo de medición, prueba, navegación y control', 3, '331280')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('265100', 'Fabricación de equipo de medición, prueba, navegación y control', 3, '331310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '331380')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '332010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('267000', 'Fabricación de instrumentos ópticos y equipo fotográfico', 3, '332020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273100', 'Fabricación de cables de fibra óptica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '332080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('265200', 'Fabricación de relojes', 3, '333000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('321100', 'Fabricación de joyas y artículos conexos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('321200', 'Fabricación de bisutería y artículos conexos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('291000', 'Fabricación de vehículos automotores', 3, '341000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('292000', 'Fabricación de carrocerías para vehículos automotores; fabricación de remolques y semirremolques', 3, '342000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('293000', 'Fabricación de partes, piezas y accesorios para vehículos automotores', 3, '343000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139200', 'Fabricación de artículos confeccionados de materiales textiles, excepto prendas de vestir', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('301100', 'Construcción de buques, embarcaciones menores y estructuras flotantes', 3, '351110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331501', 'Reparación de buques, embarcaciones menores y estructuras flotantes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('301100', 'Construcción de buques, embarcaciones menores y estructuras flotantes', 3, '351120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331501', 'Reparación de buques, embarcaciones menores y estructuras flotantes', 3, '351180')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('301200', 'Construcción de embarcaciones de recreo y de deporte', 3, '351210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331501', 'Reparación de buques, embarcaciones menores y estructuras flotantes', 3, '351280')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('302000', 'Fabricación de locomotoras y material rodante', 3, '352000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331509', 'Reparación de otros equipos de transporte n.c.p., excepto vehículos automotores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('303000', 'Fabricación de aeronaves, naves espaciales y maquinaria conexa', 3, '353010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331502', 'Reparación de aeronaves y naves espaciales', 3, '353080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('309100', 'Fabricación de motocicletas', 3, '359100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('309200', 'Fabricación de bicicletas y de sillas de ruedas', 3, '359200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('309900', 'Fabricación de otros tipos de equipo de transporte n.c.p.', 3, '359900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281600', 'Fabricación de equipo de elevación y manipulación', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('310009', 'Fabricación de colchones; fabricación de otros muebles n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331509', 'Reparación de otros equipos de transporte n.c.p., excepto vehículos automotores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('310001', 'Fabricación de muebles principalmente de madera', 3, '361010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('310009', 'Fabricación de colchones; fabricación de otros muebles n.c.p.', 3, '361020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221900', 'Fabricación de otros productos de caucho', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281700', 'Fabricación de maquinaria y equipo de oficina (excepto computadores y equipo periférico)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('293000', 'Fabricación de partes, piezas y accesorios para vehículos automotores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('301100', 'Construcción de buques, embarcaciones menores y estructuras flotantes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('302000', 'Fabricación de locomotoras y material rodante', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('303000', 'Fabricación de aeronaves, naves espaciales y maquinaria conexa', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952400', 'Reparación de muebles y accesorios domésticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('321100', 'Fabricación de joyas y artículos conexos', 3, '369100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('322000', 'Fabricación de instrumentos musicales', 3, '369200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('323000', 'Fabricación de artículos de deporte', 3, '369300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('324000', 'Fabricación de juegos y juguetes', 3, '369400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('264000', 'Fabricación de aparatos electrónicos de consumo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '369910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '369920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202909', 'Fabricación de otros productos químicos n.c.p.', 3, '369930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '369990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139900', 'Fabricación de otros productos textiles n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('151200', 'Fabricación de maletas, bolsos y artículos similares, artículos de talabartería y guarnicionería', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162900', 'Fabricación de otros productos de madera, de artículos de corcho, paja y materiales trenzables', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170900', 'Fabricación de otros artículos de papel y cartón', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221900', 'Fabricación de otros productos de caucho', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259900', 'Fabricación de otros productos elaborados de metal n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('309200', 'Fabricación de bicicletas y de sillas de ruedas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('321200', 'Fabricación de bisutería y artículos conexos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('383001', 'Recuperación y reciclamiento de desperdicios y desechos metálicos', 3, '371000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('383002', 'Recuperación y reciclamiento de papel', 3, '372010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('383003', 'Recuperación y reciclamiento de vidrio', 3, '372020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('383009', 'Recuperación y reciclamiento de otros desperdicios y desechos n.c.p.', 3, '372090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351011', 'Generación de energía eléctrica en centrales hidroeléctricas', 3, '401011')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351012', 'Generación de energía eléctrica en centrales termoeléctricas', 3, '401012')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351012', 'Generación de energía eléctrica en centrales termoeléctricas', 3, '401013')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351019', 'Generación de energía eléctrica en otras centrales n.c.p.', 3, '401019')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351020', 'Transmisión de energía eléctrica', 3, '401020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351030', 'Distribución de energía eléctrica', 3, '401030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('352020', 'Fabricación de gas; distribución de combustibles gaseosos por tubería, excepto regasificación de GNL', 3, '402000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('352010', 'Regasificación de Gas Natural Licuado (GNL)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('353001', 'Suministro de vapor y de aire acondicionado', 3, '403000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('360000', 'Captación, tratamiento y distribución de agua', 3, '410000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('431200', 'Preparación del terreno', 3, '451010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('390000', 'Actividades de descontaminación y otros servicios de gestión de desechos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('431100', 'Demolición', 3, '451020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('410010', 'Construcción de edificios para uso residencial', 3, '452010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('410020', 'Construcción de edificios para uso no residencial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('439000', 'Otras actividades especializadas de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('421000', 'Construcción de carreteras y líneas de ferrocarril', 3, '452020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('422000', 'Construcción de proyectos de servicio público', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('429000', 'Construcción de otras obras de ingeniería civil', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('439000', 'Otras actividades especializadas de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('432100', 'Instalaciones eléctricas', 3, '453000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('432200', 'Instalaciones de gasfitería, calefacción y aire acondicionado', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('432900', 'Otras instalaciones para obras de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('433000', 'Terminación y acabado de edificios', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('439000', 'Otras actividades especializadas de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('433000', 'Terminación y acabado de edificios', 3, '454000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('432900', 'Otras instalaciones para obras de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('439000', 'Otras actividades especializadas de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('439000', 'Otras actividades especializadas de construcción', 3, '455000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('451001', 'Venta al por mayor de vehículos automotores', 3, '501010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('451002', 'Venta al por menor de vehículos automotores nuevos o usados (incluye compraventa)', 3, '501020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('452001', 'Servicio de lavado de vehículos automotores', 3, '502010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522190', 'Actividades de servicios vinculadas al transporte terrestre n.c.p.', 3, '502020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('452002', 'Mantenimiento y reparación de vehículos automotores', 3, '502080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('453000', 'Venta de partes, piezas y accesorios para vehículos automotores', 3, '503000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('454001', 'Venta de motocicletas', 3, '504010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('454002', 'Venta de partes, piezas y accesorios de motocicletas', 3, '504020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('454003', 'Mantenimiento y reparación de motocicletas', 3, '504080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('473000', 'Venta al por menor de combustibles para vehículos automotores en comercios especializados', 3, '505000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('461001', 'Corretaje al por mayor de productos agrícolas', 3, '511010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('461002', 'Corretaje al por mayor de ganado', 3, '511020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('461009', 'Otros tipos de corretajes o remates al por mayor n.c.p.', 3, '511030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('619090', 'Otras actividades de telecomunicaciones n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('462020', 'Venta al por mayor de animales vivos', 3, '512110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('462090', 'Venta al por mayor de otras materias primas agropecuarias n.c.p.', 3, '512120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('462010', 'Venta al por mayor de materias primas agrícolas', 3, '512130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('462090', 'Venta al por mayor de otras materias primas agropecuarias n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463011', 'Venta al por mayor de frutas y verduras', 3, '512210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463012', 'Venta al por mayor de carne y productos cárnicos', 3, '512220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463013', 'Venta al por mayor de productos del mar (pescados, mariscos y algas)', 3, '512230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463020', 'Venta al por mayor de bebidas alcohólicas y no alcohólicas', 3, '512240')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463014', 'Venta al por mayor de productos de confitería', 3, '512250')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463030', 'Venta al por mayor de tabaco', 3, '512260')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463019', 'Venta al por mayor de huevos, lácteos, abarrotes y de otros alimentos n.c.p.', 3, '512290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464100', 'Venta al por mayor de productos textiles, prendas de vestir y calzado', 3, '513100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464901', 'Venta al por mayor de muebles, excepto muebles de oficina', 3, '513910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464902', 'Venta al por mayor de artículos eléctricos y electrónicos para el hogar', 3, '513920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464903', 'Venta al por mayor de artículos de perfumería, de tocador y cosméticos', 3, '513930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464904', 'Venta al por mayor de artículos de papelería y escritorio', 3, '513940')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464905', 'Venta al por mayor de libros', 3, '513951')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464906', 'Venta al por mayor de diarios y revistas', 3, '513952')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464907', 'Venta al por mayor de productos farmacéuticos y medicinales', 3, '513960')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464908', 'Venta al por mayor de instrumentos científicos y quirúrgicos', 3, '513970')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464909', 'Venta al por mayor de otros enseres domésticos n.c.p.', 3, '513990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466302', 'Venta al por mayor de materiales de construcción, artículos de ferretería, gasfitería y calefacción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466100', 'Venta al por mayor de combustibles sólidos, líquidos y gaseosos y productos conexos', 3, '514110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466100', 'Venta al por mayor de combustibles sólidos, líquidos y gaseosos y productos conexos', 3, '514120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466100', 'Venta al por mayor de combustibles sólidos, líquidos y gaseosos y productos conexos', 3, '514130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466100', 'Venta al por mayor de combustibles sólidos, líquidos y gaseosos y productos conexos', 3, '514140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466200', 'Venta al por mayor de metales y minerales metalíferos', 3, '514200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466301', 'Venta al por mayor de madera en bruto y productos primarios de la elaboración de madera', 3, '514310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466302', 'Venta al por mayor de materiales de construcción, artículos de ferretería, gasfitería y calefacción', 3, '514320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466901', 'Venta al por mayor de productos químicos', 3, '514910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466902', 'Venta al por mayor de desechos metálicos (chatarra)', 3, '514920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464908', 'Venta al por mayor de instrumentos científicos y quirúrgicos', 3, '514930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466909', 'Venta al por mayor de desperdicios, desechos y otros productos n.c.p.', 3, '514990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465300', 'Venta al por mayor de maquinaria, equipo y materiales agropecuarios', 3, '515001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465901', 'Venta al por mayor de maquinaria metalúrgica, para la minería, extracción de petróleo y construcción', 3, '515002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465901', 'Venta al por mayor de maquinaria metalúrgica, para la minería, extracción de petróleo y construcción', 3, '515003')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465901', 'Venta al por mayor de maquinaria metalúrgica, para la minería, extracción de petróleo y construcción', 3, '515004')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465902', 'Venta al por mayor de maquinaria para la elaboración de alimentos, bebidas y tabaco', 3, '515005')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465903', 'Venta al por mayor de maquinaria para la industria textil, del cuero y del calzado', 3, '515006')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465100', 'Venta al por mayor de computadores, equipo periférico y programas informáticos', 3, '515007')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465200', 'Venta al por mayor de equipo, partes y piezas electrónicos y de telecomunicaciones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465904', 'Venta al por mayor de maquinaria y equipo de oficina; venta al por mayor de muebles de oficina', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465905', 'Venta al por mayor de equipo de transporte(excepto vehículos automotores, motocicletas y bicicletas)', 3, '515008')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465909', 'Venta al por mayor de otros tipos de maquinaria y equipo n.c.p.', 3, '515009')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('469000', 'Venta al por mayor no especializada', 3, '519000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('471100', 'Venta al por menor en comercios de alimentos, bebidas o tabaco (supermercados e hipermercados)', 3, '521111')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472101', 'Venta al por menor de alimentos en comercios especializados (almacenes pequeños y minimarket)', 3, '521112')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('471100', 'Venta al por menor en comercios de alimentos, bebidas o tabaco (supermercados e hipermercados)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472101', 'Venta al por menor de alimentos en comercios especializados (almacenes pequeños y minimarket)', 3, '521120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('471990', 'Otras actividades de venta al por menor en comercios no especializados n.c.p.', 3, '521200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('471910', 'Venta al por menor en comercios de vestuario y productos para el hogar (grandes tiendas)', 3, '521300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477399', 'Venta al por menor de otros productos en comercios especializados n.c.p.', 3, '521900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472200', 'Venta al por menor de bebidas alcohólicas y no alcohólicas en comercios especializados (botillerías)', 3, '522010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472102', 'Venta al por menor en comercios especializados de carne y productos cárnicos', 3, '522020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472103', 'Venta al por menor en comercios especializados de frutas y verduras (verdulerías)', 3, '522030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472104', 'Venta al por menor en comercios especializados de pescado, mariscos y productos conexos', 3, '522040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472105', 'Venta al por menor en comercios especializados de productos de panadería y pastelería', 3, '522050')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477391', 'Venta al por menor de alimento y accesorios para mascotas en comercios especializados', 3, '522060')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472109', 'Venta al por menor en comercios especializados de huevos, confites y productos alimenticios n.c.p.', 3, '522070')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472109', 'Venta al por menor en comercios especializados de huevos, confites y productos alimenticios n.c.p.', 3, '522090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472300', 'Venta al por menor de tabaco y productos de tabaco en comercios especializados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477201', 'Venta al por menor de productos farmacéuticos y medicinales en comercios especializados', 3, '523111')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477201', 'Venta al por menor de productos farmacéuticos y medicinales en comercios especializados', 3, '523112')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477201', 'Venta al por menor de productos farmacéuticos y medicinales en comercios especializados', 3, '523120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477202', 'Venta al por menor de artículos ortopédicos en comercios especializados', 3, '523130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477203', 'Venta al por menor de artículos de perfumería, de tocador y cosméticos en comercios especializados', 3, '523140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477101', 'Venta al por menor de calzado en comercios especializados', 3, '523210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477102', 'Venta al por menor de prendas y accesorios de vestir en comercios especializados', 3, '523220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475100', 'Venta al por menor de telas, lanas, hilos y similares en comercios especializados', 3, '523230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477103', 'Venta al por menor de carteras, maletas y otros accesorios de viaje en comercios especializados', 3, '523240')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477102', 'Venta al por menor de prendas y accesorios de vestir en comercios especializados', 3, '523250')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475100', 'Venta al por menor de telas, lanas, hilos y similares en comercios especializados', 3, '523290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475300', 'Venta al por menor de tapices, alfombras y cubrimientos para paredes y pisos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475909', 'Venta al por menor de aparatos eléctricos, textiles para el hogar y otros enseres domésticos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477103', 'Venta al por menor de carteras, maletas y otros accesorios de viaje en comercios especializados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475909', 'Venta al por menor de aparatos eléctricos, textiles para el hogar y otros enseres domésticos n.c.p.', 3, '523310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('474200', 'Venta al por menor de equipo de sonido y de video en comercios especializados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475909', 'Venta al por menor de aparatos eléctricos, textiles para el hogar y otros enseres domésticos n.c.p.', 3, '523320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475901', 'Venta al por menor de muebles y colchones en comercios especializados', 3, '523330')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475902', 'Venta al por menor de instrumentos musicales en comercios especializados', 3, '523340')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476200', 'Venta al por menor de grabaciones de música y de video en comercios especializados', 3, '523350')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475909', 'Venta al por menor de aparatos eléctricos, textiles para el hogar y otros enseres domésticos n.c.p.', 3, '523360')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475909', 'Venta al por menor de aparatos eléctricos, textiles para el hogar y otros enseres domésticos n.c.p.', 3, '523390')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475201', 'Venta al por menor de artículos de ferretería y materiales de construcción', 3, '523410')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475202', 'Venta al por menor de pinturas, barnices y lacas en comercios especializados', 3, '523420')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475203', 'Venta al por menor de productos de vidrio en comercios especializados', 3, '523430')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477399', 'Venta al por menor de otros productos en comercios especializados n.c.p.', 3, '523911')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477393', 'Venta al por menor de artículos ópticos en comercios especializados', 3, '523912')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476400', 'Venta al por menor de juegos y juguetes en comercios especializados', 3, '523921')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('474100', 'Venta al por menor de computadores, equipo periférico, programas informáticos y equipo de telecom.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476101', 'Venta al por menor de libros en comercios especializados', 3, '523922')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476102', 'Venta al por menor de diarios y revistas en comercios especializados', 3, '523923')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476103', 'Venta al por menor de artículos de papelería y escritorio en comercios especializados', 3, '523924')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('474100', 'Venta al por menor de computadores, equipo periférico, programas informáticos y equipo de telecom.', 3, '523930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476301', 'Venta al por menor de artículos de caza y pesca en comercios especializados', 3, '523941')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477392', 'Venta al por menor de armas y municiones en comercios especializados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476302', 'Venta al por menor de bicicletas y sus repuestos en comercios especializados', 3, '523942')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476309', 'Venta al por menor de otros artículos y equipos de deporte n.c.p.', 3, '523943')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477394', 'Venta al por menor de artículos de joyería, bisutería y relojería en comercios especializados', 3, '523950')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477310', 'Venta al por menor de gas licuado en bombonas (cilindros) en comercios especializados', 3, '523961')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477395', 'Venta al por menor de carbón, leña y otros combustibles de uso doméstico en comercios especializados', 3, '523969')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477396', 'Venta al por menor de recuerdos, artesanías y artículos religiosos en comercios especializados', 3, '523991')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477397', 'Venta al por menor de flores, plantas, arboles, semillas y abonos en comercios especializados', 3, '523992')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477398', 'Venta al por menor de mascotas en comercios especializados', 3, '523993')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477391', 'Venta al por menor de alimento y accesorios para mascotas en comercios especializados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477399', 'Venta al por menor de otros productos en comercios especializados n.c.p.', 3, '523999')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475300', 'Venta al por menor de tapices, alfombras y cubrimientos para paredes y pisos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477401', 'Venta al por menor de antigüedades en comercios', 3, '524010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477402', 'Venta al por menor de ropa usada en comercios', 3, '524020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477409', 'Venta al por menor de otros artículos de segunda mano en comercios n.c.p.', 3, '524090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649209', 'Otras actividades de concesión de crédito n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479100', 'Venta al por menor por correo, por Internet y vía telefónica', 3, '525110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479100', 'Venta al por menor por correo, por Internet y vía telefónica', 3, '525120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479100', 'Venta al por menor por correo, por Internet y vía telefónica', 3, '525130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('478100', 'Venta al por menor de alimentos, bebidas y tabaco en puestos de venta y mercados (incluye ferias)', 3, '525200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('478200', 'Venta al por menor de productos textiles, prendas de vestir y calzado en puestos de venta y mercados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('478900', 'Venta al por menor de otros productos en puestos de venta y mercados (incluye ferias)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479901', 'Venta al por menor realizada por independientes en la locomoción colectiva (Ley 20.388)', 3, '525911')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479909', 'Otras actividades de venta por menor no realizadas en comercios, puestos de venta o mercados n.c.p.', 3, '525919')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479902', 'Venta al por menor mediante maquinas expendedoras', 3, '525920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479903', 'Venta al por menor por comisionistas (no dependientes de comercios)', 3, '525930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479909', 'Otras actividades de venta por menor no realizadas en comercios, puestos de venta o mercados n.c.p.', 3, '525990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479100', 'Venta al por menor por correo, por Internet y vía telefónica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952300', 'Reparación de calzado y de artículos de cuero', 3, '526010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952200', 'Reparación de aparatos de uso doméstico, equipo doméstico y de jardinería', 3, '526020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952100', 'Reparación de aparatos electrónicos de consumo (incluye aparatos de televisión y radio)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952900', 'Reparación de otros efectos personales y enseres domésticos', 3, '526030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952900', 'Reparación de otros efectos personales y enseres domésticos', 3, '526090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('802000', 'Actividades de servicios de sistemas de seguridad (incluye servicios de cerrajería)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('951200', 'Reparación de equipo de comunicaciones (incluye la reparación teléfonos celulares)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952400', 'Reparación de muebles y accesorios domésticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('551001', 'Actividades de hoteles', 3, '551010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('551002', 'Actividades de moteles', 3, '551020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('559001', 'Actividades de residenciales para estudiantes y trabajadores', 3, '551030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('551003', 'Actividades de residenciales para turistas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('559009', 'Otras actividades de alojamiento n.c.p.', 3, '551090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('551009', 'Otras actividades de alojamiento para turistas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('552000', 'Actividades de camping y de parques para casas rodantes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('561000', 'Actividades de restaurantes y de servicio móvil de comidas', 3, '552010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('561000', 'Actividades de restaurantes y de servicio móvil de comidas', 3, '552020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('563009', 'Otras actividades de servicio de bebidas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('562900', 'Suministro industrial de comidas por encargo; concesión de servicios de alimentación', 3, '552030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('562900', 'Suministro industrial de comidas por encargo; concesión de servicios de alimentación', 3, '552040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('562100', 'Suministro de comidas por encargo (Servicios de banquetería)', 3, '552050')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('562900', 'Suministro industrial de comidas por encargo; concesión de servicios de alimentación', 3, '552090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('561000', 'Actividades de restaurantes y de servicio móvil de comidas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('563009', 'Otras actividades de servicio de bebidas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('491100', 'Transporte interurbano de pasajeros por ferrocarril', 3, '601001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522190', 'Actividades de servicios vinculadas al transporte terrestre n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('491200', 'Transporte de carga por ferrocarril', 3, '601002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522190', 'Actividades de servicios vinculadas al transporte terrestre n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492110', 'Transporte urbano y suburbano de pasajeros vía metro y metrotren', 3, '602110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492120', 'Transporte urbano y suburbano de pasajeros vía locomoción colectiva', 3, '602120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492250', 'Transporte de pasajeros en buses interurbanos', 3, '602130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492130', 'Transporte de pasajeros vía taxi colectivo', 3, '602140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492210', 'Servicios de transporte de escolares', 3, '602150')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492220', 'Servicios de transporte de trabajadores', 3, '602160')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492290', 'Otras actividades de transporte de pasajeros por vía terrestre n.c.p.', 3, '602190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492190', 'Otras actividades de transporte urbano y suburbano de pasajeros por vía terrestre n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492230', 'Servicios de transporte de pasajeros en taxis libres y radiotaxis', 3, '602210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492240', 'Servicios de transporte a turistas', 3, '602220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492290', 'Otras actividades de transporte de pasajeros por vía terrestre n.c.p.', 3, '602230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492290', 'Otras actividades de transporte de pasajeros por vía terrestre n.c.p.', 3, '602290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492300', 'Transporte de carga por carretera', 3, '602300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('493090', 'Otras actividades de transporte por tuberías n.c.p.', 3, '603000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('493010', 'Transporte por oleoductos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('493020', 'Transporte por gasoductos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('501100', 'Transporte de pasajeros marítimo y de cabotaje', 3, '611001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('501200', 'Transporte de carga marítimo y de cabotaje', 3, '611002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('502100', 'Transporte de pasajeros por vías de navegación interiores', 3, '612001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('502200', 'Transporte de carga por vías de navegación interiores', 3, '612002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('511000', 'Transporte de pasajeros por vía aérea', 3, '621010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('512000', 'Transporte de carga por vía aérea', 3, '621020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('511000', 'Transporte de pasajeros por vía aérea', 3, '622001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('512000', 'Transporte de carga por vía aérea', 3, '622002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522400', 'Manipulación de la carga', 3, '630100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('521009', 'Otros servicios de almacenamiento y depósito n.c.p.', 3, '630200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522110', 'Explotación de terminales terrestres de pasajeros', 3, '630310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522120', 'Explotación de estacionamientos de vehículos automotores y parquímetros', 3, '630320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522200', 'Actividades de servicios vinculadas al transporte acuático', 3, '630330')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522300', 'Actividades de servicios vinculadas al transporte aéreo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522130', 'Servicios prestados por concesionarios de carreteras', 3, '630340')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522190', 'Actividades de servicios vinculadas al transporte terrestre n.c.p.', 3, '630390')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331509', 'Reparación de otros equipos de transporte n.c.p., excepto vehículos automotores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522200', 'Actividades de servicios vinculadas al transporte acuático', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522300', 'Actividades de servicios vinculadas al transporte aéreo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('791100', 'Actividades de agencias de viajes', 3, '630400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('791200', 'Actividades de operadores turísticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('799000', 'Otros servicios de reservas y actividades conexas (incluye venta de entradas para teatro, y otros)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522910', 'Agencias de aduanas', 3, '630910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522990', 'Otras actividades de apoyo al transporte n.c.p.', 3, '630920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522920', 'Agencias de naves', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('531000', 'Actividades postales', 3, '641100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('821900', 'Fotocopiado, preparación de documentos y otras actividades especializadas de apoyo de oficina', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('532000', 'Actividades de mensajería', 3, '641200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('611010', 'Telefonía fija', 3, '642010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('612010', 'Telefonía móvil celular', 3, '642020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('613010', 'Telefonía móvil satelital', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('611020', 'Telefonía larga distancia', 3, '642030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('613020', 'Televisión de pago satelital', 3, '642040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('611030', 'Televisión de pago por cable', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('612030', 'Televisión de pago inalámbrica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('611090', 'Otros servicios de telecomunicaciones alámbricas n.c.p.', 3, '642050')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('612090', 'Otros servicios de telecomunicaciones inalámbricas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('613090', 'Otros servicios de telecomunicaciones por satélite n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('619010', 'Centros de llamados y centros de acceso a Internet', 3, '642061')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('619010', 'Centros de llamados y centros de acceso a Internet', 3, '642062')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('619090', 'Otras actividades de telecomunicaciones n.c.p.', 3, '642090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('611090', 'Otros servicios de telecomunicaciones alámbricas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('612020', 'Radiocomunicaciones móviles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('612090', 'Otros servicios de telecomunicaciones inalámbricas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('613090', 'Otros servicios de telecomunicaciones por satélite n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('641100', 'Banca central', 3, '651100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('641910', 'Actividades bancarias', 3, '651910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649201', 'Financieras', 3, '651920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('641990', 'Otros tipos de intermediación monetaria n.c.p.', 3, '651990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649100', 'Leasing financiero', 3, '659110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649100', 'Leasing financiero', 3, '659120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649209', 'Otras actividades de concesión de crédito n.c.p.', 3, '659210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649202', 'Actividades de crédito prendario', 3, '659220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649900', 'Otras actividades de servicios financieros, excepto las de seguros y fondos de pensiones n.c.p.', 3, '659231')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661201', 'Actividades de securitizadoras', 3, '659232')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649209', 'Otras actividades de concesión de crédito n.c.p.', 3, '659290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663091', 'Administradoras de fondos de inversión', 3, '659911')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663092', 'Administradoras de fondos mutuos', 3, '659912')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663093', 'Administradoras de fices (fondos de inversión de capital extranjero)', 3, '659913')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663094', 'Administradoras de fondos para la vivienda', 3, '659914')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663099', 'Administradoras de fondos para otros fines n.c.p.', 3, '659915')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('643000', 'Fondos y sociedades de inversión y entidades financieras similares', 3, '659920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('642000', 'Actividades de sociedades de cartera', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('774000', 'Arrendamiento de propiedad intelectual y similares, excepto obras protegidas por derechos de autor', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949909', 'Actividades de otras asociaciones n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('651100', 'Seguros de vida', 3, '660101')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('652000', 'Reaseguros', 3, '660102')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663010', 'Administradoras de Fondos de Pensiones (AFP)', 3, '660200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('653000', 'Fondos de pensiones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('651210', 'Seguros generales, excepto actividades de Isapres', 3, '660301')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('651100', 'Seguros de vida', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('652000', 'Reaseguros', 3, '660302')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('651220', 'Actividades de Isapres', 3, '660400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661100', 'Administración de mercados financieros', 3, '671100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661202', 'Corredores de bolsa', 3, '671210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661203', 'Agentes de valores', 3, '671220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661209', 'Otros servicios de corretaje de valores y commodities n.c.p.', 3, '671290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661901', 'Actividades de cámaras de compensación', 3, '671910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661902', 'Administración de tarjetas de crédito', 3, '671921')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661903', 'Empresas de asesoría y consultoría en inversión financiera; sociedades de apoyo al giro', 3, '671929')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661904', 'Actividades de clasificadoras de riesgo', 3, '671930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661204', 'Actividades de casas de cambio y operadores de divisa', 3, '671940')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661909', 'Otras actividades auxiliares de las actividades de servicios financieros n.c.p.', 3, '671990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('641990', 'Otros tipos de intermediación monetaria n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('662200', 'Actividades de agentes y corredores de seguros', 3, '672010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('662100', 'Evaluación de riesgos y daños (incluye actividades de liquidadores de seguros)', 3, '672020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('662900', 'Otras actividades auxiliares de las actividades de seguros y fondos de pensiones', 3, '672090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('681011', 'Alquiler de bienes inmuebles amoblados o con equipos y maquinarias', 3, '701001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('681012', 'Compra, venta y alquiler (excepto amoblados) de inmuebles', 3, '701009')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('429000', 'Construcción de otras obras de ingeniería civil', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('681020', 'Servicios imputados de alquiler de viviendas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('682000', 'Actividades inmobiliarias realizadas a cambio de una retribución o por contrata', 3, '702000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('811000', 'Actividades combinadas de apoyo a instalaciones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('771000', 'Alquiler de vehículos automotores sin chofer', 3, '711101')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773001', 'Alquiler de equipos de transporte sin operario, excepto vehículos automotores', 3, '711102')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('771000', 'Alquiler de vehículos automotores sin chofer', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773001', 'Alquiler de equipos de transporte sin operario, excepto vehículos automotores', 3, '711200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773001', 'Alquiler de equipos de transporte sin operario, excepto vehículos automotores', 3, '711300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773002', 'Alquiler de maquinaria y equipo agropecuario, forestal, de construcción e ing. civil, sin operarios', 3, '712100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773002', 'Alquiler de maquinaria y equipo agropecuario, forestal, de construcción e ing. civil, sin operarios', 3, '712200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773003', 'Alquiler de maquinaria y equipo de oficina, sin operarios (sin servicio administrativo)', 3, '712300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773009', 'Alquiler de otros tipos de maquinarias y equipos sin operario n.c.p.', 3, '712900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772100', 'Alquiler y arrendamiento de equipo recreativo y deportivo', 3, '713010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772900', 'Alquiler de otros efectos personales y enseres domésticos (incluye mobiliario para eventos)', 3, '713020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772200', 'Alquiler de cintas de video y discos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772900', 'Alquiler de otros efectos personales y enseres domésticos (incluye mobiliario para eventos)', 3, '713030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773009', 'Alquiler de otros tipos de maquinarias y equipos sin operario n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772900', 'Alquiler de otros efectos personales y enseres domésticos (incluye mobiliario para eventos)', 3, '713090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772100', 'Alquiler y arrendamiento de equipo recreativo y deportivo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773009', 'Alquiler de otros tipos de maquinarias y equipos sin operario n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('620100', 'Actividades de programación informática', 3, '722000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('631100', 'Procesamiento de datos, hospedaje y actividades conexas', 3, '724000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581100', 'Edición de libros', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581200', 'Edición de directorios y listas de correo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581300', 'Edición de diarios, revistas y otras publicaciones periódicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581900', 'Otras actividades de edición', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('582000', 'Edición de programas informáticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('601000', 'Transmisiones de radio', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('602000', 'Programación y transmisiones de televisión', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('620200', 'Actividades de consultoría de informática y de gestión de instalaciones informáticas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('631200', 'Portales web', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('951100', 'Reparación de computadores y equipo periférico', 3, '725000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('620200', 'Actividades de consultoría de informática y de gestión de instalaciones informáticas', 3, '726000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('620900', 'Otras actividades de tecnología de la información y de servicios informáticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('721000', 'Investigaciones y desarrollo experimental en el campo de las ciencias naturales y la ingeniería', 3, '731000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('722000', 'Investigaciones y desarrollo experimental en el campo de las ciencias sociales y las humanidades', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('722000', 'Investigaciones y desarrollo experimental en el campo de las ciencias sociales y las humanidades', 3, '732000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('691001', 'Servicios de asesoramiento y representación jurídica', 3, '741110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('691002', 'Servicio notarial', 3, '741120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('691003', 'Conservador de bienes raíces', 3, '741130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('691004', 'Receptores judiciales', 3, '741140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('691009', 'Servicios de arbitraje; síndicos de quiebra y peritos judiciales; otras actividades jurídicas n.c.p.', 3, '741190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('692000', 'Actividades de contabilidad, teneduría de libros y auditoría; consultoría fiscal', 3, '741200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('732000', 'Estudios de mercado y encuestas de opinión pública', 3, '741300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('702000', 'Actividades de consultoría de gestión', 3, '741400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('701000', 'Actividades de oficinas principales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('855000', 'Actividades de apoyo a la enseñanza', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711001', 'Servicios de arquitectura (diseño de edificios, dibujo de planos de construcción, entre otros)', 3, '742110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '742121')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('091001', 'Actividades de apoyo para la extracción de petróleo y gas natural prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711002', 'Empresas de servicios de ingeniería y actividades conexas de consultoría técnica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099002', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por profesionales', 3, '742122')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('091002', 'Actividades de apoyo para la extracción de petróleo y gas natural prestados por profesionales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711003', 'Servicios profesionales de ingeniería y actividades conexas de consultoría técnica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711002', 'Empresas de servicios de ingeniería y actividades conexas de consultoría técnica', 3, '742131')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711003', 'Servicios profesionales de ingeniería y actividades conexas de consultoría técnica', 3, '742132')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711002', 'Empresas de servicios de ingeniería y actividades conexas de consultoría técnica', 3, '742141')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('741009', 'Otras actividades especializadas de diseño n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711003', 'Servicios profesionales de ingeniería y actividades conexas de consultoría técnica', 3, '742142')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('741009', 'Otras actividades especializadas de diseño n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('749009', 'Otras actividades profesionales, científicas y técnicas n.c.p.', 3, '742190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711003', 'Servicios profesionales de ingeniería y actividades conexas de consultoría técnica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('712001', 'Actividades de plantas de revisión técnica para vehículos automotores', 3, '742210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('712009', 'Otros servicios de ensayos y análisis técnicos (excepto actividades de plantas de revisión técnica)', 3, '742290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('731001', 'Servicios de publicidad prestados por empresas', 3, '743001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('731002', 'Servicios de publicidad prestados por profesionales', 3, '743002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('783000', 'Otras actividades de dotación de recursos humanos', 3, '749110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('782000', 'Actividades de agencias de empleo temporal (incluye empresas de servicios transitorios)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('781000', 'Actividades de agencias de empleo', 3, '749190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('803000', 'Actividades de investigación (incluye actividades de investigadores y detectives privados)', 3, '749210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('801001', 'Servicios de seguridad privada prestados por empresas', 3, '749221')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('802000', 'Actividades de servicios de sistemas de seguridad (incluye servicios de cerrajería)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('801002', 'Servicio de transporte de valores en vehículos blindados', 3, '749222')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('801003', 'Servicios de seguridad privada prestados por independientes', 3, '749229')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('812100', 'Limpieza general de edificios', 3, '749310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('812909', 'Otras actividades de limpieza de edificios e instalaciones industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('812901', 'Desratización, desinfección y exterminio de plagas no agrícolas', 3, '749320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('742001', 'Servicios de revelado, impresión y ampliación de fotografías', 3, '749401')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('742002', 'Servicios y actividades de fotografía', 3, '749402')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('742003', 'Servicios personales de fotografía', 3, '749409')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('829200', 'Actividades de envasado y empaquetado', 3, '749500')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('829110', 'Actividades de agencias de cobro', 3, '749911')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('829120', 'Actividades de agencias de calificación crediticia', 3, '749912')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('749001', 'Asesoría y gestión en la compra o venta de pequeñas y medianas empresas', 3, '749913')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('741001', 'Actividades de diseño de vestuario', 3, '749921')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('741002', 'Actividades de diseño y decoración de interiores', 3, '749922')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('741009', 'Otras actividades especializadas de diseño n.c.p.', 3, '749929')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('821900', 'Fotocopiado, preparación de documentos y otras actividades especializadas de apoyo de oficina', 3, '749931')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('821100', 'Actividades combinadas de servicios administrativos de oficina', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('749003', 'Servicios personales de traducción e interpretación', 3, '749932')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('749002', 'Servicios de traducción e interpretación prestados por empresas', 3, '749933')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('821900', 'Fotocopiado, preparación de documentos y otras actividades especializadas de apoyo de oficina', 3, '749934')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('749004', 'Actividades de agencias y agentes de representación de actores, deportistas y otras figuras públicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('829900', 'Otras actividades de servicios de apoyo a las empresas n.c.p.', 3, '749950')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('823000', 'Organización de convenciones y exposiciones comerciales', 3, '749961')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('823000', 'Organización de convenciones y exposiciones comerciales', 3, '749962')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('822000', 'Actividades de call-center', 3, '749970')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('829900', 'Otras actividades de servicios de apoyo a las empresas n.c.p.', 3, '749990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('639900', 'Otras actividades de servicios de información n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('731001', 'Servicios de publicidad prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('742002', 'Servicios y actividades de fotografía', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('855000', 'Actividades de apoyo a la enseñanza', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('841100', 'Actividades de la administración pública en general', 3, '751110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('681011', 'Alquiler de bienes inmuebles amoblados o con equipos y maquinarias', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('681012', 'Compra, venta y alquiler (excepto amoblados) de inmuebles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('682000', 'Actividades inmobiliarias realizadas a cambio de una retribución o por contrata', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('799000', 'Otros servicios de reservas y actividades conexas (incluye venta de entradas para teatro, y otros)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('841200', 'Regulación de las actividades de organismos que prestan servicios sanitarios, educativos, culturales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('841300', 'Regulación y facilitación de la actividad económica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('910100', 'Actividades de bibliotecas y archivos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('841100', 'Actividades de la administración pública en general', 3, '751120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('842300', 'Actividades de mantenimiento del orden público y de seguridad', 3, '751200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('841100', 'Actividades de la administración pública en general', 3, '751300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('842100', 'Relaciones exteriores', 3, '752100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('889000', 'Otras actividades de asistencia social sin alojamiento', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('842200', 'Actividades de defensa', 3, '752200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('842300', 'Actividades de mantenimiento del orden público y de seguridad', 3, '752300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('712009', 'Otros servicios de ensayos y análisis técnicos (excepto actividades de plantas de revisión técnica)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('843090', 'Otros planes de seguridad social de afiliación obligatoria n.c.p.', 3, '753010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('843010', 'Fondo Nacional de Salud (FONASA)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('843020', 'Instituto de Previsión Social (IPS)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649203', 'Cajas de compensación', 3, '753020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('843090', 'Otros planes de seguridad social de afiliación obligatoria n.c.p.', 3, '753090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850021', 'Enseñanza preescolar privada', 3, '801010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850011', 'Enseñanza preescolar pública', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850022', 'Enseñanza primaria, secundaria científico humanista y técnico profesional privada', 3, '801020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850012', 'Enseñanza primaria, secundaria científico humanista y técnico profesional pública', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850022', 'Enseñanza primaria, secundaria científico humanista y técnico profesional privada', 3, '802100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850012', 'Enseñanza primaria, secundaria científico humanista y técnico profesional pública', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850022', 'Enseñanza primaria, secundaria científico humanista y técnico profesional privada', 3, '802200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850012', 'Enseñanza primaria, secundaria científico humanista y técnico profesional pública', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('853120', 'Enseñanza superior en universidades privadas', 3, '803010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('853110', 'Enseñanza superior en universidades públicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('853201', 'Enseñanza superior en institutos profesionales', 3, '803020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('853202', 'Enseñanza superior en centros de formación técnica', 3, '803030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850022', 'Enseñanza primaria, secundaria científico humanista y técnico profesional privada', 3, '809010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850012', 'Enseñanza primaria, secundaria científico humanista y técnico profesional pública', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854901', 'Enseñanza preuniversitaria', 3, '809020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854909', 'Otros tipos de enseñanza n.c.p.', 3, '809030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854909', 'Otros tipos de enseñanza n.c.p.', 3, '809041')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854902', 'Servicios personales de educación', 3, '809049')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('861020', 'Actividades de hospitales y clínicas privadas', 3, '851110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('861010', 'Actividades de hospitales y clínicas públicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('861020', 'Actividades de hospitales y clínicas privadas', 3, '851120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('861010', 'Actividades de hospitales y clínicas públicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('862031', 'Servicios de médicos prestados de forma independiente', 3, '851211')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('862021', 'Centros médicos privados (establecimientos de atención ambulatoria)', 3, '851212')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('862010', 'Actividades de centros de salud municipalizados (servicios de salud pública)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('862032', 'Servicios de odontólogos prestados de forma independiente', 3, '851221')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('862022', 'Centros de atención odontológica privados (establecimientos de atención ambulatoria)', 3, '851222')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('869010', 'Actividades de laboratorios clínicos y bancos de sangre', 3, '851910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('869092', 'Servicios prestados de forma independiente por otros profesionales de la salud', 3, '851920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('869091', 'Otros servicios de atención de la salud humana prestados por empresas', 3, '851990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('871000', 'Actividades de atención de enfermería en instituciones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('872000', 'Actividades de atención en instituciones para personas con discapacidad mental y toxicómanos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('873000', 'Actividades de atención en instituciones para personas de edad y personas con discapacidad física', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('750001', 'Actividades de clínicas veterinarias', 3, '852010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('750002', 'Actividades de veterinarios, técnicos y otro personal auxiliar, prestados de forma independiente', 3, '852021')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('750002', 'Actividades de veterinarios, técnicos y otro personal auxiliar, prestados de forma independiente', 3, '852029')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('879000', 'Otras actividades de atención en instituciones', 3, '853100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('872000', 'Actividades de atención en instituciones para personas con discapacidad mental y toxicómanos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('873000', 'Actividades de atención en instituciones para personas de edad y personas con discapacidad física', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('889000', 'Otras actividades de asistencia social sin alojamiento', 3, '853200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('561000', 'Actividades de restaurantes y de servicio móvil de comidas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('855000', 'Actividades de apoyo a la enseñanza', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('881000', 'Actividades de asistencia social sin alojamiento para personas de edad y personas con discapacidad', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382100', 'Tratamiento y eliminación de desechos no peligrosos', 3, '900010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382200', 'Tratamiento y eliminación de desechos peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('812909', 'Otras actividades de limpieza de edificios e instalaciones industriales n.c.p.', 3, '900020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('813000', 'Actividades de paisajismo, servicios de jardinería y servicios conexos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('381100', 'Recogida de desechos no peligrosos', 3, '900030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('381200', 'Recogida de desechos peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382100', 'Tratamiento y eliminación de desechos no peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382200', 'Tratamiento y eliminación de desechos peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('370000', 'Evacuación y tratamiento de aguas servidas', 3, '900040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('370000', 'Evacuación y tratamiento de aguas servidas', 3, '900050')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('370000', 'Evacuación y tratamiento de aguas servidas', 3, '900090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('390000', 'Actividades de descontaminación y otros servicios de gestión de desechos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('941100', 'Actividades de asociaciones empresariales y de empleadores', 3, '911100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('941200', 'Actividades de asociaciones profesionales', 3, '911210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('941200', 'Actividades de asociaciones profesionales', 3, '911290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('942000', 'Actividades de sindicatos', 3, '912000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949100', 'Actividades de organizaciones religiosas', 3, '919100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949200', 'Actividades de organizaciones políticas', 3, '919200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949901', 'Actividades de centros de madres', 3, '919910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('889000', 'Otras actividades de asistencia social sin alojamiento', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949902', 'Actividades de clubes sociales', 3, '919920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949903', 'Fundaciones y corporaciones; asociaciones que promueven actividades culturales o recreativas', 3, '919930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949909', 'Actividades de otras asociaciones n.c.p.', 3, '919990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591100', 'Actividades de producción de películas cinematográficas, videos y programas de televisión', 3, '921110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591200', 'Actividades de postproducción de películas cinematográficas, videos y programas de televisión', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591300', 'Actividades de distribución de películas cinematográficas, videos y programas de televisión', 3, '921120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591400', 'Actividades de exhibición de películas cinematográficas y cintas de video', 3, '921200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('602000', 'Programación y transmisiones de televisión', 3, '921310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591100', 'Actividades de producción de películas cinematográficas, videos y programas de televisión', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('601000', 'Transmisiones de radio', 3, '921320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900001', 'Servicios de producción de obras de teatro, conciertos, espectáculos de danza, otras prod. escénicas', 3, '921411')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900009', 'Otras actividades creativas, artísticas y de entretenimiento n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900001', 'Servicios de producción de obras de teatro, conciertos, espectáculos de danza, otras prod. escénicas', 3, '921419')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900009', 'Otras actividades creativas, artísticas y de entretenimiento n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900002', 'Actividades artísticas realizadas por bandas de música, compañías de teatro, circenses y similares', 3, '921420')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900009', 'Otras actividades creativas, artísticas y de entretenimiento n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900003', 'Actividades de artistas realizadas de forma independiente: actores, músicos, escritores, entre otros', 3, '921430')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900009', 'Otras actividades creativas, artísticas y de entretenimiento n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('799000', 'Otros servicios de reservas y actividades conexas (incluye venta de entradas para teatro, y otros)', 3, '921490')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854200', 'Enseñanza cultural', 3, '921911')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('563001', 'Actividades de discotecas y cabaret (night club), con predominio del servicio de bebidas', 3, '921912')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932909', 'Otras actividades de esparcimiento y recreativas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932100', 'Actividades de parques de atracciones y parques temáticos', 3, '921920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932909', 'Otras actividades de esparcimiento y recreativas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900001', 'Servicios de producción de obras de teatro, conciertos, espectáculos de danza, otras prod. escénicas', 3, '921930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932909', 'Otras actividades de esparcimiento y recreativas n.c.p.', 3, '921990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('799000', 'Otros servicios de reservas y actividades conexas (incluye venta de entradas para teatro, y otros)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('639100', 'Actividades de agencias de noticias', 3, '922001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900004', 'Servicios prestados por periodistas independientes', 3, '922002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('742002', 'Servicios y actividades de fotografía', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('910100', 'Actividades de bibliotecas y archivos', 3, '923100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591200', 'Actividades de postproducción de películas cinematográficas, videos y programas de televisión', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('910200', 'Actividades de museos, gestión de lugares y edificios históricos', 3, '923200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('910300', 'Actividades de jardines botánicos, zoológicos y reservas naturales', 3, '923300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931109', 'Gestión de otras instalaciones deportivas n.c.p.', 3, '924110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492290', 'Otras actividades de transporte de pasajeros por vía terrestre n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931909', 'Otras actividades deportivas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932909', 'Otras actividades de esparcimiento y recreativas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931209', 'Actividades de otros clubes deportivos n.c.p.', 3, '924120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931201', 'Actividades de clubes de fútbol amateur y profesional', 3, '924131')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931201', 'Actividades de clubes de fútbol amateur y profesional', 3, '924132')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931101', 'Hipódromos', 3, '924140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931909', 'Otras actividades deportivas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931901', 'Promoción y organización de competencias deportivas', 3, '924150')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854100', 'Enseñanza deportiva y recreativa', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854100', 'Enseñanza deportiva y recreativa', 3, '924160')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931909', 'Otras actividades deportivas n.c.p.', 3, '924190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('799000', 'Otros servicios de reservas y actividades conexas (incluye venta de entradas para teatro, y otros)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('920090', 'Otras actividades de juegos de azar y apuestas n.c.p.', 3, '924910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('920010', 'Actividades de casinos de juegos', 3, '924920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932901', 'Gestión de salas de pool; gestión (explotación) de juegos electrónicos', 3, '924930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931102', 'Gestión de salas de billar; gestión de salas de bolos (bowling)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('781000', 'Actividades de agencias de empleo', 3, '924940')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932909', 'Otras actividades de esparcimiento y recreativas n.c.p.', 3, '924990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931909', 'Otras actividades deportivas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960100', 'Lavado y limpieza, incluida la limpieza en seco, de productos textiles y de piel', 3, '930100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960200', 'Peluquería y otros tratamientos de belleza', 3, '930200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960310', 'Servicios funerarios', 3, '930310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960320', 'Servicios de cementerios', 3, '930320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960310', 'Servicios funerarios', 3, '930330')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960310', 'Servicios funerarios', 3, '930390')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960320', 'Servicios de cementerios', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960902', 'Actividades de salones de masajes, baños turcos, saunas, servicio de baños públicos', 3, '930910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960909', 'Otras actividades de servicios personales n.c.p.', 3, '930990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('970000', 'Actividades de los hogares como empleadores de personal doméstico', 3, '950001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949904', 'Consejo de administración de edificios y condominios', 3, '950002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('990000', 'Actividades de organizaciones y órganos extraterritoriales', 3, '990000')"
   Call ExecSQL(DbMain, Q1)
      
End Sub
Public Sub CorrigeCodF22_1(ByVal Tbl As String)
   Dim StrCod As String
   Dim Q1 As String
   Dim Rs As Recordset
   
   If Tbl = "Cuentas" Then 'si no está usando plan predefinido, nos vamos
   
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         If vFld(Rs("Valor")) = "" Then
            Call CloseRs(Rs)
            Exit Sub       'no es plan predefinido
         End If
      Else
         Call CloseRs(Rs)
         Exit Sub       'no es plan predefinido
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   'Cod Form22 = 816
   StrCod = ""
   StrCod = StrCod & " '1010401'"
   StrCod = StrCod & ",'1010402'"
   StrCod = StrCod & ",'1010403'"
   StrCod = StrCod & ",'1010404'"
   
   StrCod = StrCod & ",'1010501'"
   StrCod = StrCod & ",'1010502'"
   StrCod = StrCod & ",'1010503'"
   StrCod = StrCod & ",'1010504'"
   StrCod = StrCod & ",'1010505'"
   StrCod = StrCod & ",'1010506'"
   StrCod = StrCod & ",'1010507'"
   StrCod = StrCod & ",'1010508'"
   StrCod = StrCod & ",'1010509'"
   StrCod = StrCod & ",'1010599'"
   
   StrCod = StrCod & ",'1010604'"
   StrCod = StrCod & ",'1010605'"
   StrCod = StrCod & ",'1010606'"
   StrCod = StrCod & ",'1010607'"
   StrCod = StrCod & ",'1010608'"
   StrCod = StrCod & ",'1010609'"
   StrCod = StrCod & ",'1010610'"
   StrCod = StrCod & ",'1010611'"
   StrCod = StrCod & ",'1010699'"
   
   StrCod = StrCod & ",'1030501'"

   Q1 = "UPDATE " & Tbl & " SET CodF22 = 816 WHERE Codigo IN (" & StrCod & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'Cod Form22 = 778
   StrCod = ""
   StrCod = StrCod & " '1030601'"
   StrCod = StrCod & ",'1030602'"
   StrCod = StrCod & ",'1030603'"
   
   Q1 = "UPDATE " & Tbl & " SET CodF22 = 778 WHERE Codigo IN (" & StrCod & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'Cod Form22 = 817
   StrCod = ""
   StrCod = StrCod & " '2010101'"
   StrCod = StrCod & ",'2010201'"
   StrCod = StrCod & ",'2010301'"
   StrCod = StrCod & ",'2010401'"
   StrCod = StrCod & ",'2010501'"
   
   StrCod = StrCod & ",'2010601'"
   StrCod = StrCod & ",'2010602'"
   StrCod = StrCod & ",'2010603'"
   StrCod = StrCod & ",'2010604'"
   StrCod = StrCod & ",'2010605'"
   StrCod = StrCod & ",'2010606'"
   StrCod = StrCod & ",'2010607'"
   StrCod = StrCod & ",'2010699'"
   
   StrCod = StrCod & ",'2011101'"
   StrCod = StrCod & ",'2011102'"
   StrCod = StrCod & ",'2011103'"
   StrCod = StrCod & ",'2011104'"
   StrCod = StrCod & ",'2011105'"
   StrCod = StrCod & ",'2011106'"
   StrCod = StrCod & ",'2011107'"
   StrCod = StrCod & ",'2011108'"
   StrCod = StrCod & ",'2011109'"
   StrCod = StrCod & ",'2011110'"
   StrCod = StrCod & ",'2011199'"
   
   Q1 = "UPDATE " & Tbl & " SET CodF22 = 817 WHERE Codigo IN (" & StrCod & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'Cod Form22 = 635
   StrCod = ""
   StrCod = StrCod & " '3020501'"
   StrCod = StrCod & ",'3020502'"
   StrCod = StrCod & ",'3020503'"
   StrCod = StrCod & ",'3020504'"
   StrCod = StrCod & ",'3020505'"
   StrCod = StrCod & ",'3020506'"
   
   Q1 = "UPDATE " & Tbl & " SET CodF22 = 635 WHERE Codigo IN (" & StrCod & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
End Sub
Public Sub CorrigeCodF22_2(ByVal Tbl As String)
   Dim StrCod As String
   Dim Q1 As String
   Dim Rs As Recordset
   
   If Tbl = "Cuentas" Then 'si no está usando plan predefinido, nos vamos
   
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         If vFld(Rs("Valor")) = "" Then
            Call CloseRs(Rs)
            Exit Sub       'no es plan predefinido
         End If
      Else
         Call CloseRs(Rs)
         Exit Sub       'no es plan predefinido
      End If
      
      Call CloseRs(Rs)
      
   End If
   
'   'Cod Form22 = 628
'   StrCod = ""
'   StrCod = StrCod & " '3010101'"
'
'   Q1 = "UPDATE " & Tbl & " SET CodF22 = 628 WHERE Codigo IN (" & StrCod & ")"
'   Call ExecSQL(DbMain, Q1)
   
   'Cod Form22 = 843
   StrCod = ""
   StrCod = StrCod & " '2030000'"
   
   Q1 = "UPDATE " & Tbl & " SET CodF22 = 843 WHERE Codigo IN (" & StrCod & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'Cod Form22 = 844
   StrCod = ""
   StrCod = StrCod & " '2030101'"
   StrCod = StrCod & ",'2030201'"
   StrCod = StrCod & ",'2030301'"
   StrCod = StrCod & ",'2030501'"
   StrCod = StrCod & ",'2031101'"
   StrCod = StrCod & ",'2031201'"
   
   Q1 = "UPDATE " & Tbl & " SET CodF22 = 844 WHERE Codigo IN (" & StrCod & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
End Sub

Public Sub CorrigeTipoCapPropio_1(ByVal Tbl As String)
   Dim Q1 As String
   Dim Rs As Recordset
   
   If Tbl = "Cuentas" Then 'si no está usando plan predefinido, nos vamos
   
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         If vFld(Rs("Valor")) = "" Then
            Call CloseRs(Rs)
            Exit Sub       'no es plan predefinido
         End If
      Else
         Call CloseRs(Rs)
         Exit Sub       'no es plan predefinido
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   Q1 = "UPDATE " & Tbl & " SET TipoCapPropio = " & CAPPROPIO_ACTIVO_VALINTO
   Q1 = Q1 & " WHERE Codigo IN('1010601', '1010602')"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
   
End Sub

'2814014 pipe
Public Function CorrigeBaseAdm_V372() As Boolean   'agregada 27 may 2022 2814014
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
    Dim MaxTipoDoc As Long
   Dim Rs As Recordset

    Dim Base As Integer, Base2 As Integer

   Base = 10000
   Base2 = 20000


   On Error Resume Next

   '--------------------- Versión 372 -----------------------------------

   If lDbVerAdm = 372 And lUpdOK = True Then

      'agregamos tipo doc OII = "Otros Ingr. Saldo Inicial Libro de Caja" y OEI = "Otros Egr. Saldo Inicial Libro de Caja"
      Q1 = "SELECT Id FROM TipoDocs WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'VPEE'"
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
          Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, CodF29Count, CodF29IVA, TieneAfecto, TieneExento, ExigeRUT)"
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Vale Pago Electronico con Exento', 'VPEE', 'ACTIVO', -1, 110, 111, -1, -1, 0)"
         Call ExecSQL(DbMain, Q1)

    Q1 = ""
      Q1 = "UPDATE TipoDocs SET CodF29ExCount = 586, CodF29Exento = 142 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'VPEE' "
      Call ExecSQL(DbMain, Q1)

      Q1 = ""
      Q1 = "UPDATE TipoDocs SET CodF29Neto = 6111 "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND Diminutivo = 'VPEE' "
      Call ExecSQL(DbMain, Q1)

        
      Q1 = "UPDATE TipoDocs SET"
      Q1 = Q1 & "  CodF29Neto = " & Base + 111
      Q1 = Q1 & " WHERE TipoDocs.Diminutivo = 'VPEE'"
      Call ExecSQL(DbMain, Q1)
      


         
      'Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneAfecto = " & VAL_OBLIGATORIO & " WHERE Diminutivo IN('BOV', 'BEX', 'DVB', 'FAV', 'MRG', 'VPE') ")
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneAfecto = " & VAL_OBLIGATORIO & " WHERE Diminutivo IN('VPEE') ")

      'Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneExento = " & VAL_OBLIGATORIO & " WHERE Diminutivo IN('BEX', 'BOE', 'EXP', 'FVE', 'VSD') ")
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneExento = " & VAL_OBLIGATORIO & " WHERE Diminutivo IN('VPEE') ")


       'Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneNumDocHasta = " & VAL_OBLIGATORIO & " WHERE Diminutivo IN('BOV', 'BEX', 'BOE', 'MRG') ")
      Call ExecSQL(DbMain, "UPDATE TipoDocs SET TieneNumDocHasta = " & VAL_OBLIGATORIO & " WHERE Diminutivo IN('VPEE') ")
          
      End If

      Call CloseRs(Rs)




      If lUpdOK Then
         lDbVerAdm = 373
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V372 = lUpdOK

End Function


'3092471
Public Function CampoImportEmpresas() As Boolean   'Agregada 15 jun 2023
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 739 -----------------------------------
 '  If lDbVerAdm = 739 And lUpdOK = True Then
   
      Err.Clear
      
      'Agregamos campo idCCosto a MovActivoFijo
      Set Tbl = DbMain.TableDefs("Empresas")
     
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("Import", dbLong)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         'MsgBeep vbExclamation
         'MsgBox "Error " & Err & ", " & Error & vbLf & "Empresas.Import", vbExclamation
         lUpdOK = False
      End If
      
      If Err <> 0 Then
         'MsgBeep vbExclamation
         'MsgBox "Error " & Err & ", " & Error, vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
            
'      If lUpdOK Then
'         lDbVerAdm = 374
'         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
'         Call ExecSQL(DbMain, Q1)
'      End If
   
  ' End If
   
   CampoImportEmpresas = lUpdOK

End Function
'3092471
