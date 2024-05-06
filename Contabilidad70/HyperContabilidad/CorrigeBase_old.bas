Attribute VB_Name = "MCorrigeBase"
Option Explicit

Private lDbVer As Integer
Private lUpdOK As Boolean
Private lEmpAnoEnArchivo As Integer   'corresponde al año del archivo, que no necesariamente es el gEmpresa.Ano cuando se invoca al CorrigeBase del año anterior de una empresa
Private lEmpIdEnArchivo As Long        'corresponde al IdEmpresa del archivo, para usar cuando se invoca al CorrigeBase del año anterior de una empresa en el CrearNuevoAño, enque aún el gEmpresa.Id es 0
Public Const MAX_COL = 64
' Para hacer manteciones a ciertas tablas con manejo de versión
Public Sub CorrigeBase()
   Dim Q1 As String
   Dim Rs As Recordset

   If Not gEmprSeparadas Then
      Exit Sub
   End If
   
   On Error Resume Next
   

   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa WHERE Tipo = " & TPE_DBINFO
   Q1 = Q1 & " ORDER BY Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do Until Rs.EOF
         
      Select Case vFld(Rs("Codigo"))
      
         Case 1: ' RUT
         
         Case 2: ' Ano
            lEmpAnoEnArchivo = Val(vFld(Rs("Valor")))
            
         Case 3: ' idEmpresa
            lEmpIdEnArchivo = Val(vFld(Rs("Valor")))
         
      End Select
   
      Rs.MoveNext
   Loop
  
   Call CloseRs(Rs)
   
   'Agregamos IdEmpresa y Año a ParamEmpresa porque se usa de en algunos CorrigeBase, por si no están ya en la base.
   
   Call AppendIdEmpresaAno("ParamEmpresa", True)    'por si no tiene estos datos adicionales
   
   lDbVer = 0
   lUpdOK = True
   
   If Not CorrigeBase_2005_01() Then
      Exit Sub
   End If
   
   If Not CorrigeBase_V22() Then    '12 Julio 2005
      Exit Sub
   End If
   
   If Not CorrigeBase_V23() Then    '25 Agosto 2005
      Exit Sub
   End If
   
   If Not CorrigeBase_V24() Then    '15 Octubre 2005
      Exit Sub
   End If
   
   If Not CorrigeBase_V25() Then    '14 Diciembre 2005
      Exit Sub
   End If
   
   If Not CorrigeBase_V26() Then    '23 Diciembre 2005
      Exit Sub
   End If
   
   If Not CorrigeBase_V27() Then    '23 Enero 2006
      Exit Sub
   End If
   
   If Not CorrigeBase_V28() Then    '08 Marzo 2006
      Exit Sub
   End If
   
   If Not CorrigeBase_V29() Then    '17 Marzo 2006
      Exit Sub
   End If
   
   If Not CorrigeBase_V30() Then    '30 Marzo 2006
      Exit Sub
   End If
   
   If Not CorrigeBase_V31() Then    '05 Abril 2006
      Exit Sub
   End If
     
   If Not CorrigeBase_V32() Then    '24 Abril 2006
      Exit Sub
   End If
     
   If Not CorrigeBase_V33() Then    '15 Junio 2006  corresponde a V 1.0.22
      Exit Sub
   End If
     
   If Not CorrigeBase_V34() Then    '27 Nov 2006 v.1.0.26
      Exit Sub
   End If
   
   If Not CorrigeBase_V35() Then    '1 Feb 2008 v.1.0.36
      Exit Sub
   End If
   
   If Not CorrigeBase_V36() Then    '11 Mar 2008 v.1.0.37
      Exit Sub
   End If
   
   If Not CorrigeBase_V37() Then    '17 Jun 2008 v.2.0.1
      Exit Sub
   End If
  
   If Not CorrigeBase_V38() Then    '3 Jul 2008 v.2.0.1
      Exit Sub
   End If
   
   If Not CorrigeBase_V39() Then    '15 Jul 2008 v.2.0.4
      Exit Sub
   End If
   
   If Not CorrigeBase_V40() Then    '21 Jul 2008 v.2.0.4
      Exit Sub
   End If
   
   If Not CorrigeBase_V41() Then    '9 oct. 2008 v.2.0.4
      Exit Sub
   End If
     
   If Not CorrigeBase_V42() Then    '25 nov. 2008 v.2.0.4
      Exit Sub
   End If
    
   If Not CorrigeBase_V43() Then    '21 abr. 2009 v.2.0.4
      Exit Sub
   End If
    
   If Not CorrigeBase_V44() Then    '11 Jun 2009 v.2.0.5
      Exit Sub
   End If
                
   If Not CorrigeBase_V45() Then    '23 Oct.2009 v.2.0.7
      Exit Sub
   End If
                
   If Not CorrigeBase_V46() Then    '30 nov.2009 v.2.0.7
      Exit Sub
   End If
                   
   If Not CorrigeBase_V47() Then    '7 ene 2010 v.2.0.7
      Exit Sub
   End If
   
   If Not CorrigeBase_V48() Then    '13 dic 2010 v.2.0.9
      Exit Sub
   End If
                   
   If Not CorrigeBase_V49() Then    '21 Enero 2011 v.2.0.9
      Exit Sub
   End If
   
   If Not CorrigeBase_V50() Then    '7 abr 2011 v2.0.11    eliminacion de códigos Form 22, reporte 53  OJO Poner en versión 3!!!
      Exit Sub
   End If
   
   'aqui se da inicio a versión 3.0
   
   If Not CorrigeBase_V299() Then    '13 julio 2011
      Exit Sub
   End If
                   
   If Not CorrigeBase_V300() Then    '29 julio 2011
      Exit Sub
   End If
                   
   If Not CorrigeBase_V301() Then    '31 agosto 2011
      Exit Sub
   End If
   
   If Not CorrigeBase_V302() Then    '7 sept 2011
      Exit Sub
   End If
   
   If Not CorrigeBase_V303() Then    '26 oct. 2011
      Exit Sub
   End If
   
   If Not CorrigeBase_V304() Then    '16 nov. 2011
      Exit Sub
   End If
   
   If Not CorrigeBase_V305() Then    '16 abr 2012
      Exit Sub
   End If
   
   If Not CorrigeBase_V306() Then    '26 sep 2012   IFRS
      Exit Sub
   End If
   
   If Not CorrigeBase_V307() Then    '26 sep 2012
      Exit Sub
   End If

   If Not CorrigeBase_V308() Then    '12 dic 12
      Exit Sub
   End If

   If Not CorrigeBase_V309() Then    '12 dic 12
      Exit Sub
   End If
   
   If Not CorrigeBase_V310() Then    '10 jul 2013
      Exit Sub
   End If
   
   If Not CorrigeBase_V311() Then    '13 ago 2013
      Exit Sub
   End If
   
   If Not CorrigeBase_V312() Then    'entregada 10 sep 2013
      Exit Sub
   End If
   
   If Not CorrigeBase_V313() Then    'entregada 14 sept 2013
      Exit Sub
   End If
   
   If Not CorrigeBase_V314() Then    'entregada 8 nov 2013
      Exit Sub
   End If
   
   If Not CorrigeBase_V315() Then    'entregada 13 nov 2013
      Exit Sub
   End If
   
   If Not CorrigeBase_V316() Then    'entregada 10 ene 2014
      Exit Sub
   End If
   
   If Not CorrigeBase_V317() Then    'entregada 21 enero 2014
      Exit Sub
   End If
   
   If Not CorrigeBase_V318() Then    'entregada 7 jun 2014
      Exit Sub
   End If
   
   If Not CorrigeBase_V319() Then    'entregada 20 jun 14
      Exit Sub
   End If
   
   If Not CorrigeBase_V320() Then    'entregada 21 ago 2014
      Exit Sub
   End If
         
   If Not CorrigeBase_V321() Then    'entregada 4 sept 2014
      Exit Sub
   End If
         
   If Not CorrigeBase_V322() Then    'entregada 23 oct 2014
      Exit Sub
   End If
   
   If Not CorrigeBase_V323() Then    'entregada dic 2014
      Exit Sub
   End If
         
   If Not CorrigeBase_V324() Then    'entregada 29/01/2015
      Exit Sub
   End If
   
   If Not CorrigeBase_V325() Then    'entregada 8/07/15
      Exit Sub
   End If
   
   If Not CorrigeBase_V326() Then    'entregada 19/10/15
      Exit Sub
   End If
   
   If Not CorrigeBase_V327() Then    'entregada 4 dic 2015
      Exit Sub
   End If
   
   If Not CorrigeBase_V328() Then    'entregada 4 mar 2016
      Exit Sub
   End If
            
   If Not CorrigeBase_V329() Then    'agregada 25 jul 2016 - entregada
      Exit Sub
   End If
            
   If Not CorrigeBase_V330() Then    'agregada 5 oct 2016 - entregada 20 oct. 2016
      Exit Sub
   End If
            
   If Not CorrigeBase_V331() Then    'agregada 26 oct 2016 - entregada
      Exit Sub
   End If
            
   If Not CorrigeBase_V332() Then    'agregada 26 oct 2016 - entregada
      Exit Sub
   End If
            
   If Not CorrigeBase_V333() Then   'agregada 1 dic 2016
      Exit Sub
   End If
            
   If Not CorrigeBase_V334() Then   'agregada 3 mar 2017
      Exit Sub
   End If
            
   If Not CorrigeBase_V335() Then   'agregada 11 abril 2017
      Exit Sub
   End If
   
   If Not CorrigeBase_V336() Then   'agregada 13 abril 2017
      Exit Sub
   End If
            
   If Not CorrigeBase_V337() Then   'agregada 24 abril 2017
      Exit Sub
   End If
            
   If Not CorrigeBase_V338() Then   'agregada 15 mayo 2017
      Exit Sub
   End If
            
   If Not CorrigeBase_V339() Then   'entregada 7 jul 2017
      Exit Sub
   End If
            
   If Not CorrigeBase_V340() Then   'agregada 7 jul 2017
      Exit Sub
   End If
            

   If Not CorrigeBase_V341() Then   'agregada 11 jul 2017
      Exit Sub
   End If
 
   If Not CorrigeBase_V342() Then   'agregada 20 jul 2017, entregada 1 ago 2017
      Exit Sub
   End If
 
   If Not CorrigeBase_V343() Then   'agregada 2 ago 2017
      Exit Sub
   End If
 
   If Not CorrigeBase_V344() Then   'agregada 16 ago 2017
      Exit Sub
   End If
 
   If Not CorrigeBase_V345() Then   'agregada 30 ago 2017
      Exit Sub
   End If

   If Not CorrigeBase_V346() Then   'agregada 5 sept 2017
      Exit Sub
   End If

   If Not CorrigeBase_V347() Then   'agregada 12 sept 2017
      Exit Sub
   End If

   If Not CorrigeBase_V348() Then   'agregada 9 nov 2017
      Exit Sub
   End If

   If Not CorrigeBase_V349() Then   'agregada 14 dic 2017
      Exit Sub
   End If

   If Not CorrigeBase_V350() Then   'agregada 27 dic 2017
      Exit Sub
   End If

   If Not CorrigeBase_V351() Then   'agregada 4 ene 2018
      Exit Sub
   End If

   If Not CorrigeBase_V352() Then   'agregada 10 ene 2018
      Exit Sub
   End If

   If Not CorrigeBase_V353() Then   'agregada 17 ene 2018
      Exit Sub
   End If

   If Not CorrigeBase_V354() Then   'agregada 4 abr 2018
      Exit Sub
   End If
   
   If Not CorrigeBase_V355() Then   'agregada 19 abr 2018
      Exit Sub
   End If

   If Not CorrigeBase_V356() Then   'agregada 14 may 2018
      Exit Sub
   End If

   If Not CorrigeBase_V357() Then   'agregada 22 may 2018
      Exit Sub
   End If

   If Not CorrigeBase_V358() Then   'agregada 4 jun 2018
      Exit Sub
   End If

   If Not CorrigeBase_V359() Then   'agregada 16 ago 2018
      Exit Sub
   End If

   If Not CorrigeBase_V360() Then   'agregada 4 sep 2018
      Exit Sub
   End If

   If Not CorrigeBase_V361() Then   'agregada 17 oct 2018
      Exit Sub
   End If

   If Not CorrigeBase_V362() Then   'agregada 29 oct 2018
      Exit Sub
   End If

   If Not CorrigeBase_V363() Then   'agregada 27 feb 2019
      Exit Sub
   End If

   If Not CorrigeBase_V364() Then   'agregada 3 may 2019
      Exit Sub
   End If
   
   If Not CorrigeBase_V365() Then   'Inicio Versión 7.0.0 - agregada 19 nov 2018
      Exit Sub
   End If

   If Not CorrigeBase_V700() Then   'agregada 7 ene 2019
      Exit Sub
   End If

   If Not CorrigeBase_V701() Then   'agregada 14 ene 2019
      Exit Sub
   End If

   If Not CorrigeBase_V702() Then   'agregada 17 ene 2019
      Exit Sub
   End If

   If Not CorrigeBase_V703() Then   'agregada 27 feb 2019
      Exit Sub
   End If

   If Not CorrigeBase_V704() Then   'agregada 6 ago 2019
      Exit Sub
   End If

   If Not CorrigeBase_V705() Then   'agregada 27 dic 2019
      Exit Sub
   End If

   If Not CorrigeBase_V706() Then   'agregada 3 ene 2020
      Exit Sub
   End If

   If Not CorrigeBase_V707() Then   'agregada 6 ene 2020
      Exit Sub
   End If

   If Not CorrigeBase_V708() Then   'agregada 9 ene 2020
      Exit Sub
   End If

   If Not CorrigeBase_V709() Then   'agregada 21 ene 2020
      Exit Sub
   End If

   If Not CorrigeBase_V710() Then   'agregada 28 ene 2020
      Exit Sub
   End If

   If Not CorrigeBase_V711() Then   'agregada 2 mar 2020
      Exit Sub
   End If

   If Not CorrigeBase_V712() Then   'agregada 6 may 2020
      Exit Sub
   End If

   If Not CorrigeBase_V713() Then   'agregada 10 jun 2020
      Exit Sub
   End If

   If Not CorrigeBase_V714() Then   'agregada 4 ago 2020
      Exit Sub
   End If

'   Call DelCompAperturaDuplicadosSY      'Para Solucionar Caso Soledad Yañez 7 sept 2020
   
   If Not CorrigeBase_V715() Then   'agregada 9 sep 2020
      Exit Sub
   End If

   If Not CorrigeBase_V716() Then   'agregada 16 sep 2020
      Exit Sub
   End If

   If Not CorrigeBase_V717() Then   'agregada 23 sep 2020
      Exit Sub
   End If

   If Not CorrigeBase_V718() Then   'agregada 27 oct 2020
      Exit Sub
   End If

   If Not CorrigeBase_V719() Then   'agregada 2 nov 2020
      Exit Sub
   End If

   If Not CorrigeBase_V720() Then   'agregada 16 nov 2020
      Exit Sub
   End If

   If Not CorrigeBase_V721() Then   'agregada 14 dic 2020
      Exit Sub
   End If

   If Not CorrigeBase_V722() Then   'agregada 18 ene 2021
      Exit Sub
   End If

   If Not CorrigeBase_V723() Then   'agregada 22 jun 2021
      Exit Sub
   End If

   If Not CorrigeBase_V724() Then   'agregada 9 sep 2021
      Exit Sub
   End If

   If Not CorrigeBase_V725() Then   'agregada 29 sep 2021
      Exit Sub
   End If

   If Not CorrigeBase_V726() Then   'agregada 4 oct 2021
      Exit Sub
   End If

   If Not CorrigeBase_V727() Then   'agregada 22 oct 2021
      Exit Sub
   End If
   
   If Not CorrigeBase_V728() Then   'agregada 19 nov 2021 FPG ADO 2678539
      Exit Sub
   End If
   
   If Not CorrigeBase_V729() Then   'agregada 10 ene 2022 FPG ADO 2699584 Item 4.2
      Exit Sub
   End If
   
   If Not CorrigeBase_V730() Then   'agregada 10 ene 2022 FPG ADO
      Exit Sub
   End If
   
   '2863248
   If Not CorrigeBase_V731() Then   'agregada 5 ago 2022 FFV ADO 2863248
      Exit Sub
   End If
   'fin 2863248

    '2860036
   If Not CorrigeBase_V732() Then   'agregada 19 ago 2022 FFV ADO 2860036
      Exit Sub
   End If
   'fin 2860036

    '2861570
     If Not CorrigeBase_V733() Then   'agregada 20 sep 2022 FFV ADO 2861570
      Exit Sub
     End If
    'fin 2861570
    
    If Not CorrigeBase_V734() Then   'agregada 20 sep 2022 FFV ADO 2861570
      Exit Sub
     End If
     
    If Not CorrigeBase_V735() Then   'agregada 18 oct 2022 FPR ADO 2913643
      Exit Sub
     End If
     
    If Not CorrigeBase_V736() Then   'agregada 25 ENE 2023 FPR ADO 2862611
      Exit Sub
     End If
     
     '3043065 SF 14006128
     If Not CorrigeBase_V737() Then   'agregada 25 ENE 2023 FPR ADO 2862611
      Exit Sub
     End If
     '3043065 SF 14006128
     
     '2861733 tema 2
    If Not CorrigeBase_V738() Then   'agregada 08 MAR 2023 FFV ADO 2932636 tema 2
      Exit Sub
     End If


'   Call AppendIdEmpresaAno("Comprobante", True)      'para solucionar tema de cliente
'   Call AppendIdEmpresaAno("MovComprobante", True)
   
   If lDbVer > 739 Then
      MsgBox1 "¡ ATENCION !" & vbCrLf & vbCrLf & "La base de datos corresponde a una versión posterior de este programa." & vbCrLf & "Debe actualizar el programa antes de continuar, de lo contrario podría dañar la información..", vbCritical
      Call CloseDb(DbMain)
      End
   End If

End Sub


'2861733 tema 2

Public Function CorrigeBase_V738() As Boolean   'Agregada 14 dic 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 738 -----------------------------------
   If lDbVer = 738 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo idCCosto a MovActivoFijo
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("idCCosto", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.idCCosto", vbExclamation
         lUpdOK = False
      End If
      
      
       'Agregamos campo IdAreaNeg a MovActivoFijo
       ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdAreaNeg", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.IdAreaNeg", vbExclamation
         lUpdOK = False
      End If
      
      
      Q1 = "DROP INDEX idx_IdCCosto ON MovActivoFijo "
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX idx_IdCCosto ON MovActivoFijo (IdCCosto) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "DROP INDEX idx_IdAreaNeg ON MovActivoFijo "
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX idx_IdAreaNeg ON MovActivoFijo (IdAreaNeg) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      
       If ERR <> 0 Then
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error, vbExclamation
         lUpdOK = False
      End If
            
      If lUpdOK Then
         lDbVer = 739
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V738 = lUpdOK

End Function
'2861733 tema 2

'3043065 SF 14006128
Public Function CorrigeBase_V737() As Boolean   'Agregada 14 dic 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double
   Dim DbMainV737  As Database
    Dim idxDocumento As Index

   On Error Resume Next
   
   '--------------------- Versión 737 -----------------------------------
   If lDbVer = 737 And lUpdOK = True Then
   
       ERR.Clear
    '14690904
    Call CloseDb(DbMain)
    Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
    '14690904
    
   Set DbMainV737 = DbMain
    
    'Call OpenDbEmp2(DbMainV737, gEmpresa.Rut, gEmpresa.Ano)
    
       Q1 = "DROP INDEX idx_TipoLibDoc ON Documento "
      Call ExecSQL(DbMainV737, Q1, False)

      Q1 = "CREATE INDEX idx_TipoLibDoc ON Documento (TipoLib) "
      Rc = ExecSQL(DbMainV737, Q1, True)
      
       Q1 = "DROP INDEX idx_FEmision ON Documento "
      Call ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE INDEX idx_FEmision ON Documento (FEmision) "
      Rc = ExecSQL(DbMain, Q1, False)
      
       Q1 = "DROP INDEX idx_Exento ON Documento "
      Call ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE INDEX idx_Exento ON Documento (Exento) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "DROP INDEX idx_TipoRetencion ON Documento "
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX idx_TipoRetencion ON Documento (TipoRetencion) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "DROP INDEX idx_PorcentRetencion ON Documento "
      Call ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE INDEX idx_PorcentRetencion ON Documento (PorcentRetencion) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "DROP INDEX idx_IdEmpresa ON Documento "
      Call ExecSQL(DbMain, Q1, False)
      
       Q1 = "CREATE INDEX idx_IdEmpresa ON Documento (IdEmpresa) "
      Rc = ExecSQL(DbMain, Q1, False)
      
       Q1 = "DROP INDEX idx_AnoDOC ON Documento "
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX idx_AnoDOC ON Documento (Ano) "
      Rc = ExecSQL(DbMain, Q1, False)
      
       If ERR <> 0 Then
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error, vbExclamation
         lUpdOK = False
      End If
      
      If lUpdOK Then
         lDbVer = 738
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V737 = lUpdOK

End Function
'3043065 SF 14006128

Public Function CorrigeBase_V736() As Boolean 'Agregada 25 ENE 2023 FPG
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 736 -----------------------------------
   If lDbVer = 736 And lUpdOK = True Then
   
      ERR.Clear

      Q1 = "INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) VALUES (" & gEmpresa.id & "," & gEmpresa.Ano & ",'DATOSII',1,'1')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) VALUES (" & gEmpresa.id & "," & gEmpresa.Ano & ",'DATOSII',2,'0')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) VALUES (" & gEmpresa.id & "," & gEmpresa.Ano & ",'DATOSII',3,'0')"
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVer = 737
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V736 = lUpdOK

End Function

Public Function CorrigeBase_V735() As Boolean 'Agregada 19 nov 2021 FPG Solicitado por Victor Morales
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 735 -----------------------------------
   If lDbVer = 735 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo CodArea a tabla Empresa
      Set Tbl = DbMain.TableDefs("Empresa")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodArea", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.CodArea", vbExclamation
         lUpdOK = False
      End If
      
      'Agregamos campo Celular a tabla Empresa
      Set Tbl = DbMain.TableDefs("Empresa")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Celular", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.Celular", vbExclamation
         lUpdOK = False
      End If
      
      'Agregamos campo Celular a tabla Empresa
      Set Tbl = DbMain.TableDefs("Empresa")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Villa", dbText, 80)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.Villa", vbExclamation
         lUpdOK = False
      End If
      
      If lUpdOK Then
         lDbVer = 736
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V735 = lUpdOK

End Function


Private Function CorrigeBase_V734() As Boolean
                                                                                          
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 734 -----------------------------------
   If lDbVer = 734 And lUpdOK = True Then

        ERR.Clear
      'Call CompactDb2(DbMain, True, gEmpresa.ConnStr)  'no hubo error
      If OpenDbEmp() Then
          Set Tbl = DbMain.TableDefs("Documento")
          
          'agregamos campo Tratamiento a tabla Documento
          ERR.Clear
          Tbl.Fields.Append Tbl.CreateField("Tratamiento", dbLong)
    
          If ERR = 0 Then
             Tbl.Fields.Refresh
          ElseIf ERR <> 3191 Then ' ya existe
             MsgBeep vbExclamation
             MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.Tratamiento", vbExclamation
             lUpdOK = False
          End If
      End If

        Q1 = "INSERT INTO Param "
        Q1 = Q1 & " (Tipo, Codigo, Valor)"
        Q1 = Q1 & " VALUES( 'TIPOLIB', 8, 'Otros Documentos Full')"
        Call ExecSQL(DbMain, Q1)

        Q1 = "INSERT INTO Param "
        Q1 = Q1 & " (Tipo, Codigo, Valor)"
        Q1 = Q1 & " VALUES( 'TIPOLIBCOD', 8, 'LIBOTROFULL')"
        Call ExecSQL(DbMain, Q1)

        Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja, DocImpExp, DocBoletas, CodDocSII, CodDocDTESII)"
         Q1 = Q1 & " VALUES(8,1, 'Otros Documentos Full', 'ODF', 'ACTIVO', -1, 0, 1, 0, 0, 0, 0, '', '')"
         Call ExecSQL(DbMain, Q1)

        Q1 = "INSERT INTO Param "
        Q1 = Q1 & " (Tipo, Codigo, Valor)"
        Q1 = Q1 & " VALUES( 'TRATAMIENTO', 1, 'Activo')"
        Call ExecSQL(DbMain, Q1)

        Q1 = "INSERT INTO Param "
        Q1 = Q1 & " (Tipo, Codigo, Valor)"
        Q1 = Q1 & " VALUES( 'TRATAMIENTO', 2, 'Pasivo')"
        Call ExecSQL(DbMain, Q1)

'        If Not IsObject(DbMain.TableDefs("DocumentoFull")) Then
'            Call DuplicTable(DbMain, "Documento", "DocumentoFull")
'        End If
'
'        Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
'
'        If Not IsObject(DbMain.TableDefs("ComprobanteFull")) Then
'            Call DuplicTable(DbMain, "Comprobante", "ComprobanteFull")
'        End If
'
'        Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
'
'        If Not IsObject(DbMain.TableDefs("MovComprobanteFull")) Then
'            Call DuplicTable(DbMain, "MovComprobante", "MovComprobanteFull")
'
'
'        End If
'
        'Call CloseDb(DbMain)
        'Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
'
'
'        If IsObject(DbMain.TableDefs("DocumentoFull")) Then
'             ERR.Clear
'****************************
''            ERR.Clear
''
''             Set Tbl = DbMain.TableDefs("Documento")
''
''             ERR.Clear
''             Tbl.Fields.Append Tbl.CreateField("Tratamiento", dbLong)
''
''             If ERR = 0 Then
''                Tbl.Fields.Refresh
''             ElseIf ERR <> 3191 Then ' ya existe
''                MsgBeep vbExclamation
''                MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.Tratamiento", vbExclamation
''                lUpdOK = False
''             End If
'*****************************+


      'Set DbMain = AuxDb

'        End If
        'Call CloseDb(DbMain)
                                                                                                                
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 735
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V734 = lUpdOK

End Function
'2861570
Private Function CorrigeBase_V733() As Boolean
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 733 -----------------------------------
   If lDbVer = 733 And lUpdOK = True Then
         
      ERR.Clear
      
'      If gFunciones.ExpFUT Then   'se elimina esta verificación porque crea problemas con empresas nuevas hasta que saquemos actualización que ya no usa campo (FCA 29/11/2017)
         Call CreateTblFirmas
         
'      End If
      
    
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 734
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V733 = lUpdOK

End Function
'2861570

'2860036
Private Function CorrigeBase_V732() As Boolean
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 732 -----------------------------------
   If lDbVer = 732 And lUpdOK = True Then
         
      ERR.Clear
      
'      If gFunciones.ExpFUT Then   'se elimina esta verificación porque crea problemas con empresas nuevas hasta que saquemos actualización que ya no usa campo (FCA 29/11/2017)
         Call CreateTblMembrete
         
'      End If
      
    
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 733
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V732 = lUpdOK

End Function
'fn 2860036



'2863248
Private Function CorrigeBase_V731() As Boolean
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
    Q1 = "SELECT iddoc, numdoc,femision,femisionori,fechacreacion FROM documento WHERE femision < 0 and fechacreacion > 0 "
      Set Rs = OpenRs(DbMain, Q1)
      
      Do While Not Rs.EOF
        
         Q1 = "UPDATE documento SET femision= fechacreacion,femisionori=fechacreacion,estado =  " & ED_PENDIENTE
         Q1 = Q1 & " where iddoc = " & vFld(Rs("iddoc"))
         Call ExecSQL(DbMain, Q1)
      
      
         Rs.MoveNext
         
      Loop
      
      CloseRs (Rs)

    If lDbVer = 731 And lUpdOK = True Then
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 732
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
  End If
   
   CorrigeBase_V731 = lUpdOK

End Function

'fin 2863248

Private Function CorrigeBase_V730() As Boolean
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 730 -----------------------------------
   If lDbVer = 730 And lUpdOK = True Then
         
      ERR.Clear
      
'      If gFunciones.ExpFUT Then   'se elimina esta verificación porque crea problemas con empresas nuevas hasta que saquemos actualización que ya no usa campo (FCA 29/11/2017)
         Call CreateTblPercepciones
         Call CreateTblDetPercepciones
'      End If
      Set Tbl = DbMain.TableDefs("Cuentas")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Percepcion", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.Percepcion", vbExclamation
         lUpdOK = False
      End If

    
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 731
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V730 = lUpdOK

End Function

Public Function CorrigeBase_V729() As Boolean 'Agregada 10 ene 2022 FPG Solicitado por Victor Morales ADO 2699584 Item 4.2
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 729 -----------------------------------
   If lDbVer = 729 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo PorcExisteOServ a tabla Empresa
      Set Tbl = DbMain.TableDefs("Empresa")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("PorcExisteOServ", dbInteger)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.PorcExisteOServ", vbExclamation
         lUpdOK = False
      End If
      
      
      If lUpdOK Then
         lDbVer = 730
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V729 = lUpdOK

End Function

Public Function CorrigeBase_V728() As Boolean 'Agregada 19 nov 2021 FPG Solicitado por Victor Morales
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 728 -----------------------------------
   If lDbVer = 728 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo FDesde3Porc a tabla Entidades
      Set Tbl = DbMain.TableDefs("Entidades")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FDesde3Porc", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.FDesde3Porc", vbExclamation
         lUpdOK = False
      End If
      
      'Agregamos campo FHasta3Porc a tabla Entidades
      Set Tbl = DbMain.TableDefs("Entidades")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FHasta3Porc", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.FHasta3Porc", vbExclamation
         lUpdOK = False
      End If
      
      If lUpdOK Then
         lDbVer = 729
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V728 = lUpdOK

End Function


Public Function CorrigeBase_V727() As Boolean   'Agregada 21 oct 2021
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 727 -----------------------------------
   If lDbVer = 727 And lUpdOK = True Then
   
      ERR.Clear
   
      'Actualizamos EsTotalDoc para Libro de Retenciones
      'Para solucionar tema del clientes en Boletas de Pago
      Q1 = "UPDATE MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc SET EsTotalDoc = 0 "
      Q1 = Q1 & " WHERE IdTipoValLib = " & LIBRETEN_IMPUESTO
      Q1 = Q1 & " AND Documento.TipoLib = " & LIB_RETEN
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVer = 728
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V727 = lUpdOK

End Function

Public Function CorrigeBase_V726() As Boolean 'Agregada 4 oct 2021
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 726 -----------------------------------
   If lDbVer = 726 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo Ret3Porc a tabla Entidades
      Set Tbl = DbMain.TableDefs("Entidades")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Ret3Porc", dbBoolean)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.Ret3Porc", vbExclamation
         lUpdOK = False
      End If
      
      'Agregamos campo ValRet3Porc a tabla Documento
      Set Tbl = DbMain.TableDefs("Documento")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("ValRet3Porc", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.ValRet3Porc", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdCuentaRet3Porc", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.IdCuentaRet3Porc", vbExclamation
         lUpdOK = False
      End If

      
      If lUpdOK Then
         lDbVer = 727
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V726 = lUpdOK

End Function


Public Function CorrigeBase_V725() As Boolean   'Agregada 29 sep 2021
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 725 -----------------------------------
   If lDbVer = 725 And lUpdOK = True Then
   
      Call ExecSQL(DbMain, "DROP TABLE AjusteIVAMensual", False)   'Esto es por error de haber creado la tabla en la LPContab.mdb y linkearla - FCA 29 sep 2021
      Q1 = "CREATE TABLE AjusteIVAMensual ( IdEmpresa Long, Ano Integer, "
      Q1 = Q1 & "Mes Byte, Valor Double )"
      Call ExecSQL(DbMain, Q1)
      DbMain.TableDefs.Refresh
      
      If lUpdOK Then
         lDbVer = 726
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V725 = lUpdOK

End Function


Public Function CorrigeBase_V724() As Boolean   'Agregada 9 sep 2021
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 724 -----------------------------------
   If lDbVer = 724 And lUpdOK = True Then
   
      ERR.Clear
   
      'Actualizamos EsTotalDoc por si acaso  FCA 8 sep 2021
      'Para solucionar tema de cliente que importó documentos desde Facturación
      Q1 = "UPDATE MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc SET EsTotalDoc = -1 "
      Q1 = Q1 & " WHERE IdTipoValLib = " & LIBVENTAS_TOTAL
      Q1 = Q1 & " AND Documento.TipoLib IN ( " & LIB_VENTAS & "," & LIB_COMPRAS & ")"
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVer = 725
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V724 = lUpdOK

End Function


Public Function CorrigeBase_V723() As Boolean   'Agregada 22 jun 2021
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 723 -----------------------------------
   If lDbVer = 723 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo MontoAfectaBaseImp a tabla LibroCaja
      Set Tbl = DbMain.TableDefs("LibroCaja")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("MontoAfectaBaseImp", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LibroCaja.MontoAfectaBaseImp", vbExclamation
         lUpdOK = False
      End If
            
      If lUpdOK Then
         lDbVer = 724
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V723 = lUpdOK

End Function


Public Function CorrigeBase_V722() As Boolean   'Agregada 18 ene 2021
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 722 -----------------------------------
   If lDbVer = 722 And lUpdOK = True Then
   
      ERR.Clear
      
       'Agregamos tabla BaseImponible14D
      
      Set Tbl = New TableDef
      Tbl.Name = "BaseImponible14D"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdBaseImponible14D", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14D.IdBaseImponible14D", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdEmpresa", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14D.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("Ano", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14D.Ano", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
      Set Fld = Tbl.CreateField("Tipo", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14D.Tipo", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("Nivel", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14D.Nivel", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("Codigo", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14D.Codigo", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Fecha", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14D.Fecha", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Valor", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14D.Valor", vbExclamation
         lUpdOK = False
      End If

      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON BaseImponible14D (IdBaseImponible14D) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)

         Q1 = "CREATE UNIQUE INDEX IdxItem ON BaseImponible14D (Empresa, Ano, Codigo, Fecha)"
         Rc = ExecSQL(DbMain, Q1, False)

      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla BaseImponible14D", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing

            
      If lUpdOK Then
         lDbVer = 723
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V722 = lUpdOK

End Function

Public Function CorrigeBase_V721() As Boolean   'Agregada 14 dic 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 721 -----------------------------------
   If lDbVer = 721 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo DepLey21256 a MovActivoFijo
      Set Tbl = DbMain.TableDefs("Entidades")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqTribEnt", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.FranqTribEnt", vbExclamation
         lUpdOK = False
      End If
            
      If lUpdOK Then
         lDbVer = 722
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V721 = lUpdOK

End Function

Public Function CorrigeBase_V720() As Boolean   'Agregada 16 nov 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 720 -----------------------------------
   If lDbVer = 720 And lUpdOK = True Then
   
      ERR.Clear
      
      'agrandamos campo Orden en Comprobante Tipo
      Call AlterField(DbMain, "CT_MovComprobante", "Orden", dbInteger)
      
      'Agregamos campo DepLey21256 a MovActivoFijo
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DepLey21256", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepLey21256", vbExclamation
         lUpdOK = False
      End If

      'Agregamos campo DepLey21256 a MovActivoFijo
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DepLey21256Hist", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepLey21256Hist", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear

            
      If lUpdOK Then
         lDbVer = 721
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V720 = lUpdOK

End Function

Public Function CorrigeBase_V719() As Boolean   'Agregada 2 nov 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 719 -----------------------------------
   If lDbVer = 719 And lUpdOK = True Then
   
      ERR.Clear
      
      'copiamos Capital Ptopio Tributario desde tabla ParamEmpresa a EmpresasAno dado que ahor ase requiere revisar años antriores en el informe de CPS
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'CAPPROPIO' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         CapPropio = Val(vFld(Rs("Valor")))
      End If
      Call CloseRs(Rs)
      
      If CapPropio <> 0 Then
         Q1 = "UPDATE EmpresasAno SET CPS_CapPropioTrib = " & CapPropio
         Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      End If
            
      If lUpdOK Then
         lDbVer = 720
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V719 = lUpdOK

End Function



Public Function CorrigeBase_V718() As Boolean   'Agregada 27 oct 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
  

   On Error Resume Next
   
   '--------------------- Versión 718 -----------------------------------
   If lDbVer = 718 And lUpdOK = True Then
   
      ERR.Clear
      
      
      'Agregamos campo Descrip a tabla DetCapPropioSimpl
      
      Set Tbl = DbMain.TableDefs("DetCapPropioSimpl")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Descrip", dbText, 80)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "DetCapPropioSimpl.Descrip", vbExclamation
         lUpdOK = False
      End If
      Set Tbl = Nothing
     
      
      'parche para captura de documentos registro de compras y ventas SII, con documento asociado
      
      Q1 = "UPDATE Documento LEFT JOIN Documento as Documento1 ON Documento.IdDocAsoc = Documento1.IdDoc"
      Q1 = Q1 & " SET Documento.IdDocAsoc = NULL, Documento.TipoDocAsoc = NULL, Documento.NumDocAsoc = NULL, Documento.DTEDocAsoc = NULL, "
      Q1 = Q1 & " Documento.SaldoDoc = NULL "
      Q1 = Q1 & " WHERE Documento1.IdDoc Is Null"
      Q1 = Q1 & " AND NOT  Documento.IdDocAsoc IS NULL"
      Q1 = Q1 & " AND  Documento.IdDocAsoc <> 0"
      
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVer = 719
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V718 = lUpdOK

End Function


Public Function CorrigeBase_V717() As Boolean   'Agregada 23 sept 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
  

   On Error Resume Next
   
   '--------------------- Versión 717 -----------------------------------
   If lDbVer = 717 And lUpdOK = True Then
   
      ERR.Clear
      
      
      'Agregamos tabla DetCapPropioSimpl
      
      Set Tbl = New TableDef
      Tbl.Name = "DetCapPropioSimpl"

      ERR.Clear
      Set Fld = Tbl.CreateField("IdDetCapPropioSimpl", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetCapPropioSimpl.IdDetCapPropioSimpl", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdEmpresa", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetCapPropioSimpl.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("Ano", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetCapPropioSimpl.Ano", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("TipoDetCPS", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetCapPropioSimpl.TipoDetCPS", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IngresoManual", dbBoolean)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetCapPropioSimpl.IngresoManual", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdCuenta", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetCapPropioSimpl.IdCuenta", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("CodCuenta", dbText, 15)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetCapPropioSimpl.CodCuenta", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Fecha", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetCapPropioSimpl.Fecha", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdMovComp", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetCapPropioSimpl.IdMovComp", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Valor", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetCapPropioSimpl.Valor", vbExclamation
         lUpdOK = False
      End If

      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON DetCapPropioSimpl (IdDetCapPropioSimpl) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE INDEX IdxEmpAno ON DetCapPropioSimpl (IdEmpresa, Ano, TipoDetCPS, IngresoManual, Fecha, CodCuenta )"
         Rc = ExecSQL(DbMain, Q1, False)
      

      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla DetCapPropioSimpl", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing
      If lUpdOK Then
         lDbVer = 718
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V717 = lUpdOK

End Function


Public Function CorrigeBase_V716() As Boolean   'Agregada 16 sept 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 716 -----------------------------------
   If lDbVer = 716 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("Socios")
      
      'agregamos campo MontoIngresadoUsuario
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("MontoIngresadoUsuario", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Socios.MontoIngresadoUsuario", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo MontoIngresadoUsuario
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("MontoATraspasar", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Socios.MontoATraspasar", vbExclamation
         lUpdOK = False
      End If
                 
      If lUpdOK Then
         lDbVer = 717
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V716 = lUpdOK

End Function


Public Function CorrigeBase_V715() As Boolean   'Agregada 9 sep 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 715 -----------------------------------
   If lDbVer = 715 And lUpdOK = True Then
   
      ERR.Clear
      
      'Tabla ImpAdic mejoramos índice
      Q1 = "DROP INDEX IdxTipo ON ImpAdic "
      Call ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE INDEX IdxTipo ON ImpAdic (IdEmpresa, Ano, TipoLib, TipoValor)"
      Call ExecSQL(DbMain, Q1, False)
           
      If lUpdOK Then
         lDbVer = 716
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V715 = lUpdOK

End Function


Public Function CorrigeBase_V714() As Boolean   'Agregada 5 ago 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 714 -----------------------------------
   If lDbVer = 714 And lUpdOK = True Then
   
      ERR.Clear
      
      'Parche por algunos casos en que quedó con año y código anterior
      Q1 = "UPDATE ImpAdic INNER JOIN Cuentas ON ImpAdic.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & " SET ImpAdic.Ano = " & lEmpAnoEnArchivo & ", ImpAdic.CodCuenta = Cuentas.Codigo "
      Call ExecSQL(DbMain, Q1)
       
           
      If lUpdOK Then
         lDbVer = 715
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V714 = lUpdOK

End Function
    

Public Function CorrigeBase_V713() As Boolean   'Agregada 10 jun 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 713 -----------------------------------
   If lDbVer = 713 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("Empresa")
      
      'agregamos campo Franq14ASemiIntegrado
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Franq14ASemiIntegrado", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.Franq14ASemiIntegrado", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo FranqProPymeGeneral
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqProPymeGeneral", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqProPymeGeneral", vbExclamation
         lUpdOK = False
      End If
    
      'agregamos campo FranqProPymeTransp
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqProPymeTransp", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqProPymeTransp", vbExclamation
         lUpdOK = False
      End If
    
      'agregamos campo FranqRentasPresuntas
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqRentasPresuntas", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqRentasPresuntas", vbExclamation
         lUpdOK = False
      End If
    
      'agregamos campo FranqRentaEfectiva
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqRentaEfectiva", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqRentaEfectiva", vbExclamation
         lUpdOK = False
      End If
    
      'agregamos campo FranqOtro
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqOtro", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqOtro", vbExclamation
         lUpdOK = False
      End If
    
      'agregamos campo FranqNoSujetoArt14
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqNoSujetoArt14", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqNoSujetoArt14", vbExclamation
         lUpdOK = False
      End If
    
           
      If lUpdOK Then
         lDbVer = 714
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V713 = lUpdOK

End Function

Public Function CorrigeBase_V712() As Boolean   'Agregada 6 may 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 712 -----------------------------------
   If lDbVer = 712 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo TipoDepLey21210 a MovActivoFijo
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoDepLey21210", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.TipoDepLey21210", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
      
      'Agregamos campo DepDecimaParte2 a MovActivoFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DepDecimaParte2", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepDecimaParte2", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
     
      'Agregamos campo DepDecimaParte2Hist a MovActivoFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DepDecimaParte2Hist", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepDecimaParte2Hist", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
     
      'Agregamos campo PatenteRol a MovActivoFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("PatenteRol", dbText, 30)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.PatenteRol", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
     
      'Agregamos campo NombreProy a MovActivoFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NombreProy", dbText, 60)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.NombreProy", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
     
      'Agregamos campo FechaProy a MovActivoFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FechaProy", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.FechaProy", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
     
      Tbl.Fields.Append Tbl.CreateField("TipoDepLey21210Hist", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.TipoDepLey21210Hist", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
           
      If lUpdOK Then
         lDbVer = 713
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V712 = lUpdOK

End Function
            

Public Function CorrigeBase_V711() As Boolean   'agregada 2 mar 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 711 -----------------------------------

   If lDbVer = 711 And lUpdOK = True Then

      'Agregamos tabla CtasAjustesExContRLI
      
      Set Tbl = New TableDef
      Tbl.Name = "CtasAjustesExContRLI"

      ERR.Clear
      Set Fld = Tbl.CreateField("IdCtaAjustesRLI", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExContRLI.IdCtaAjustesRLI", vbExclamation
         lUpdOK = False
      End If
      
     
      ERR.Clear
      Set Fld = Tbl.CreateField("IdEmpresa", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExContRLI.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
     
      ERR.Clear
      Set Fld = Tbl.CreateField("Ano", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExContRLI.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
     
      ERR.Clear
      Set Fld = Tbl.CreateField("TipoAjuste", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExContRLI.TipoAjuste", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdGrupo", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExContRLI.IdGrupo", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdItem", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExContRLI.IdItem", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdCuenta", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExContRLI.IdCuenta", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("CodCuenta", dbText, 15)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExContRLI.CodCuenta", vbExclamation
         lUpdOK = False
      End If

      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON CtasAjustesExContRLI (IdCtaAjustesRLI) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE INDEX IdxItem ON CtasAjustesExContRLI (IdEmpresa, Ano, TipoAjuste, IdGrupo, IdItem)"
         Rc = ExecSQL(DbMain, Q1, False)
      

      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla CtasAjustesExContRLI", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing
   
      If lUpdOK Then
         lDbVer = 712
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V711 = lUpdOK

End Function

Public Function CorrigeBase_V710() As Boolean   'Agregada 28 ene 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 710 -----------------------------------
   If lDbVer = 710 And lUpdOK = True Then
   
      ERR.Clear
      
       'Agregamos campo DocOtroEsCargo a tabla Documento (esto para traer los OtrosDocs del año anterior)
      Set Tbl = DbMain.TableDefs("Documento")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DocOtroEsCargo", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.DocOtroEsCargo", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
     
           
      If lUpdOK Then
         lDbVer = 711
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V710 = lUpdOK

End Function
Public Function CorrigeBase_V709() As Boolean   'Agregada 22 ene 2020   (este corrige base no lo ponemos en SQL Server dado que aún no hay clientes que lo estén usando)
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 709 -----------------------------------
   If lDbVer = 709 And lUpdOK = True Then
   
      ERR.Clear
      
      'esta función verifica que el idEMpresa y el año en cada una de las tablas corresponda. Repetimos aplicación porque a algunos clientes al momento de crear el nuevo año y hacer el corrigebase del año anterior el gEmpresa.Id no estaba asignado y quedaba en cero en las tablas
      Call VerificaTblEmpAno
           
      If lUpdOK Then
         lDbVer = 710
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V709 = lUpdOK

End Function


Public Function CorrigeBase_V708() As Boolean   'Agregada 9 ene 2020   (este corrige base no lo ponemos en SQL Server dado que aún no hay clientes que lo estén usando)
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 708 -----------------------------------
   If lDbVer = 708 And lUpdOK = True Then
   
      ERR.Clear
      
      'esta función verifica que el año en cada una de las tablas corresponda. Repetimos aplicación porque a algunos cliente no se aplicó la anterior (versión 706)
      Call VerificaTblEmpAno
           
      If lUpdOK Then
         lDbVer = 709
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V708 = lUpdOK

End Function


Public Function CorrigeBase_V707() As Boolean   'Agregada 6 ene 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 707 -----------------------------------
   If lDbVer = 707 And lUpdOK = True Then
   
      ERR.Clear
      
      'agregamos campo a la tabla Socios para la cantidad de acciones
      
      Set Tbl = DbMain.TableDefs("Socios")
      
      'agregamos campo CantAcciones
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CantAcciones", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Socios.CantAcciones", vbExclamation
         lUpdOK = False
      End If
                 
      If lUpdOK Then
         lDbVer = 708
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V707 = lUpdOK

End Function


Public Function CorrigeBase_V706() As Boolean   'Agregada 3 ene 2020   (este corrige base no lo ponemos en SQL Server dado que aún no hay clientes que lo estén usando)
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 706 -----------------------------------
   If lDbVer = 706 And lUpdOK = True Then
   
      ERR.Clear
      
      'cambiamos código plan SII para cuenta Deudores incobrables en la tabla Cuentas
      QBase = "UPDATE Cuentas "

      Q1 = " SET CodCtaPlanSII ='3030700' WHERE Codigo = '3010644' AND " & GenLike(DbMain, "Deudores Incobrables", "Descripcion")
      Call ExecSQL(DbMain, QBase & Q1)
    
      'esta función verifica que el año en cada una de las tablas corresponda
      Call VerificaTblEmpAno
           
      If lUpdOK Then
         lDbVer = 707
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V706 = lUpdOK

End Function


Public Function CorrigeBase_V705() As Boolean   'Agregada 27 dic 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 705 -----------------------------------
   If lDbVer = 705 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("Empresa")
      
      'agregamos campo FranqSocProfPrimCat
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqSocProfPrimCat", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqSocProfPrimCat", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo FranqSocProfSegCat
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqSocProfSegCat", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqSocProfSegCat", vbExclamation
         lUpdOK = False
      End If
    
           
      If lUpdOK Then
         lDbVer = 706
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V705 = lUpdOK

End Function

Public Function CorrigeBase_V704() As Boolean   'agregada 8 ago 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   If Not gEmprSeparadas Then
      CorrigeBase_V704 = lUpdOK
      Exit Function
   End If
   
   On Error Resume Next

   '--------------------- Versión 704 -----------------------------------

   If lDbVer = 704 And lUpdOK = True Then
   
      Q1 = "UPDATE Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & " SET IdCuentaAfecto = MovDocumento.IdCuenta "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND MovDocumento.IdTipoValLib = " & LIBCOMPRAS_AFECTO
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & " SET IdCuentaExento = MovDocumento.IdCuenta "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND MovDocumento.IdTipoValLib = " & LIBCOMPRAS_EXENTO
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & " SET IdCuentaTotal = MovDocumento.IdCuenta "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND MovDocumento.IdTipoValLib = " & LIBCOMPRAS_TOTAL
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & " SET IdCuentaAfecto = MovDocumento.IdCuenta "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND MovDocumento.IdTipoValLib = " & LIBVENTAS_AFECTO
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & " SET IdCuentaExento = MovDocumento.IdCuenta "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND MovDocumento.IdTipoValLib = " & LIBVENTAS_EXENTO
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & " SET IdCuentaTotal = MovDocumento.IdCuenta "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_VENTAS & " AND MovDocumento.IdTipoValLib = " & LIBVENTAS_TOTAL
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & " SET IdCuentaAfecto = MovDocumento.IdCuenta "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_RETEN & " AND MovDocumento.IdTipoValLib = " & LIBRETEN_BRUTO
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & " SET IdCuentaExento = MovDocumento.IdCuenta "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_RETEN & " AND MovDocumento.IdTipoValLib = " & LIBRETEN_HONORSINRET
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & " SET IdCuentaTotal = MovDocumento.IdCuenta "
      Q1 = Q1 & " WHERE TipoLib = " & LIB_RETEN & " AND MovDocumento.IdTipoValLib = " & LIBRETEN_NETO
      Call ExecSQL(DbMain, Q1)
      
   
   
      
      If lUpdOK Then
         lDbVer = 705
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V704 = lUpdOK

End Function
Public Function CorrigeBase_V703() As Boolean   'agregada 26 feb 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   If Not gEmprSeparadas Then
      CorrigeBase_V703 = lUpdOK
      Exit Function
   End If
   
   On Error Resume Next

   '--------------------- Versión 703 -----------------------------------

   If lDbVer = 703 And lUpdOK = True Then
   
      Q1 = "DROP INDEX IdCuenta ON CtasAjustesExCont "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "DROP INDEX IdItem ON CtasAjustesExCont "
      Rc = ExecSQL(DbMain, Q1, False)
            
      Q1 = "CREATE INDEX IdxItem ON CtasAjustesExCont (IdEmpresa, Ano, TipoAjuste, IdItem)"
      Rc = ExecSQL(DbMain, Q1, False)
      
      
      If lUpdOK Then
         lDbVer = 704
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V703 = lUpdOK

End Function
Public Function CorrigeBase_V702() As Boolean   'agregada 17 ene 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   If Not gEmprSeparadas Then
      CorrigeBase_V702 = lUpdOK
      Exit Function
   End If
   
   On Error Resume Next

   '--------------------- Versión 702 -----------------------------------

   If lDbVer = 702 And lUpdOK = True Then
   
      'Agregamos campo IdCtaAfectoOld a tabla Documento (esto para traer los datos del año anterior al crear nuevo año en empresas juntas)
      Set Tbl = DbMain.TableDefs("ImpAdic")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Ano", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "ImpAdic.Ano", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
                  
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCuenta", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "ImpAdic.CodCuenta", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
                                    
               
      Q1 = "UPDATE ImpAdic INNER JOIN Cuentas ON ImpAdic.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & " SET ImpAdic.Ano = " & lEmpAnoEnArchivo & ", ImpAdic.CodCuenta = Cuentas.Codigo "
      Call ExecSQL(DbMain, Q1)
      
               
      If lUpdOK Then
         lDbVer = 703
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V702 = lUpdOK

End Function
Public Function CorrigeBase_V701() As Boolean   'agregada 14 ene 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   If Not gEmprSeparadas Then
      CorrigeBase_V701 = lUpdOK
      Exit Function
   End If
   
   On Error Resume Next

   '--------------------- Versión 700 -----------------------------------

   If lDbVer = 701 And lUpdOK = True Then
   
      'Agregamos campo IdCtaAfectoOld a tabla Documento (esto para traer los datos del año anterior al crear nuevo año en empresas juntas)
      Set Tbl = DbMain.TableDefs("Documento")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaAfectoOld", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.CodCtaAfectoOld", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
                  
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaExentoOld", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.CodCtaExentoOld", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
                                    
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaTotalOld", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.CodCtaTotalOld", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
      
      'Agregamos campo IdCompOld a tabla CT_Comprobante (esto para traer los datos de comprobantes tipo de empresa vacía (Reset de Comp Tipo))
      Set Tbl = DbMain.TableDefs("CT_Comprobante")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdCompOld", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "CT_Comprobante.IdCompOld", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      
      'asignamos IdEmpresa a Comprobante Tipo
      Q1 = "UPDATE CT_Comprobante SET IdEmpresa = " & gEmpresa.id
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE CT_MovComprobante SET IdEmpresa = " & gEmpresa.id
      Call ExecSQL(DbMain, Q1)
               
      If lUpdOK Then
         lDbVer = 702
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V701 = lUpdOK

End Function
Public Function CorrigeBase_V700() As Boolean   'agregada 7 ene 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   If Not gEmprSeparadas Then
      CorrigeBase_V700 = lUpdOK
      Exit Function
   End If
   
   On Error Resume Next

   '--------------------- Versión 700 -----------------------------------

   If lDbVer = 700 And lUpdOK = True Then
   
      'Agregamos campo IdCuentaOld a tabla Cuentas (esto para traer los datos del año anterior al crear nuevo año en empresas juntas)
      Set Tbl = DbMain.TableDefs("Cuentas")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdCuentaOld", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.IdCuentaOld", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
            
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdPadreOld", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.IdPadreOld", vbExclamation
         lUpdOK = False
      End If

      Set Tbl = DbMain.TableDefs("CuentasBasicas")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdCuentaOld", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "CuentasBasicas.IdCuentaOld", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
                  
      Set Tbl = DbMain.TableDefs("ParamEmpresa")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("ValorOld", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "ParamEmpresa.ValorOld", vbExclamation
         lUpdOK = False
      End If
      ERR.Clear
      
      Q1 = "DROP INDEX CodFECU ON Cuentas "
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Agregamos campo Vigente a Sucursales
      Set Tbl = DbMain.TableDefs("Sucursales")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Vigente", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Sucursales.Vigente", vbExclamation
         lUpdOK = False
      End If
     
      Q1 = "UPDATE Sucursales SET Vigente = -1 "
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVer = 701
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V700 = lUpdOK

End Function


Public Function CorrigeBase_V365() As Boolean   'agregada 12/7/2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   If Not gEmprSeparadas Then
      CorrigeBase_V365 = lUpdOK
      Exit Function
   End If
   
   On Error Resume Next

   '--------------------- Versión 365 -----------------------------------

   If lDbVer = 365 And lUpdOK = True Then
   
      'Agregamos IdEmpresa y Ano a las tablas que corresponde
      
      Call AppendIdEmpresaAno("ActFijoCompsFicha", True)
      Call AppendIdEmpresaAno("ActFijoFicha", True)
      Call AppendIdEmpresaAno("AFComponentes", False)
      Call AppendIdEmpresaAno("AFGrupos", False)
      Call AppendIdEmpresaAno("AjustesExtLibCaja", True)
      Call AppendIdEmpresaAno("AreaNegocio", False)
      Call AppendIdEmpresaAno("AsistImpPrimCat", True)
      Call AppendIdEmpresaAno("BaseImponible14Ter", True)
      Call AppendIdEmpresaAno("Cartola", True)
      Call AppendIdEmpresaAno("CentroCosto", False)
      Call AppendIdEmpresaAno("Colores", False, False)
      Call AppendIdEmpresaAno("Comprobante", True)
      Call AppendIdEmpresaAno("Contactos", False)
      Call AppendIdEmpresaAno("CT_Comprobante", False)
      Call AppendIdEmpresaAno("CT_MovComprobante", False)
      Call AppendIdEmpresaAno("CtasAjustesExCont", True)
      Call AppendIdEmpresaAno("Cuentas", True)
      Call AppendIdEmpresaAno("CuentasBasicas", True)
      Call AppendIdEmpresaAno("CuentasRazon", False, False)
      Call AppendIdEmpresaAno("DetCartola", True)
      Call AppendIdEmpresaAno("DetSaldosAp", True)
      Call AppendIdEmpresaAno("DocCuotas", True)
      Call AppendIdEmpresaAno("Documento", True)
'      Call AppendIdEmpresaAno("Empresa", True)             'ya tiene IdEmpresa (se llama Id)
      Call AppendIdEmpresaAno("Entidades", False)
      Call AppendIdEmpresaAno("EstadoMes", True, False)
      Call AppendIdEmpresaAno("Glosas", False)
      Call AppendIdEmpresaAno("ImpAdic", True)
'      Call AppendIdEmpresaAno("InfoAnualDJ1847", True)    'ya tiene IdEmpresa, ano e índice
      Call AppendIdEmpresaAno("LibroCaja", True)
      Call AppendIdEmpresaAno("LockAction", True)
      Call AppendIdEmpresaAno("LogComprobantes", True)
      Call AppendIdEmpresaAno("LogImpreso", True)
      Call AppendIdEmpresaAno("MovActivoFijo", True)
      Call AppendIdEmpresaAno("MovComprobante", True)
      Call AppendIdEmpresaAno("MovDocumento", True)
      Call AppendIdEmpresaAno("Notas", False)
'      Call AppendIdEmpresaAno("ParamEmpresa", True)        'ya se hace al inicio del CorrigeBase
      Call AppendIdEmpresaAno("ParamRazon", False, False)
      Call AppendIdEmpresaAno("PropIVA_TotMensual", True, False)
      Call AppendIdEmpresaAno("Socios", True)
      Call AppendIdEmpresaAno("Sucursales", False)
      
      
      'arreglamos índices ya existentes
      
      'Tabla Colores
      Q1 = "DROP INDEX Nivel ON Colores "
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX IdxEmpresa ON Colores (IdEmpresa, Nivel) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla CT_Comprobante
      Q1 = "DROP INDEX IdComp ON CT_Comprobante "
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX IdComp ON CT_Comprobante (IdComp) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla CT_MovComprobante
      Q1 = "DROP INDEX IdMov ON CT_MovComprobante "
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX IdMov ON CT_Comprobante (IdMov) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla CtasAjustesExCont
      Q1 = "DROP INDEX IdCtaAjustes ON CtasAjustesExCont "
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX IdCtaAjustes ON CtasAjustesExCont (IdCtaAjustes) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla Cuentas
      Q1 = "DROP INDEX IdCuenta ON Cuentas "
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX IdCuenta ON Cuentas (IdCuenta) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla CuentasRazon
      Q1 = "DROP INDEX IdxRazon ON CuentasRazon "
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE INDEX IdxRazon ON CuentasRazon (IdEmpresa, IdRazon)"
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla Empresa
      Q1 = "DROP INDEX Id ON Empresa "
      Rc = ExecSQL(DbMain, Q1, False)

      Set Tbl = DbMain.TableDefs("Empresa")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Ano", dbInteger)
   
      If ERR = 0 Then
         Tbl.Fields.Refresh
      End If
      
      Q1 = "UPDATE Empresa SET Ano = " & lEmpAnoEnArchivo
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE UNIQUE INDEX IdxEmpresa ON Empresa (Id, Ano) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla EstadoMes
      Q1 = "DROP INDEX Estado ON EstadoMes "
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "DROP INDEX Mes ON EstadoMes "
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX IdxEmpresa ON EstadoMes (IdEmpresa, Ano, Mes) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla LogComprobantes
      Q1 = "DROP INDEX IdLog ON LogComprobantes "
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX IdLog ON LogComprobantes (IdLog) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla MovComprobante
      Q1 = "CREATE INDEX IdCCosto ON MovComprobante (IdCCosto) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX IdAreaNeg ON MovComprobante (IdAreaNeg) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla ParamRazon
      Q1 = "DROP INDEX IdxParamRazon ON ParamRazon "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE UNIQUE INDEX IdxParamRazon ON ParamRazon (IdEmpresa, IdRazon) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla PropIVA_TotMensual
      Q1 = "DROP INDEX IdxMes ON PropIVA_TotMensual "
      Rc = ExecSQL(DbMain, Q1, False)
            
      Q1 = "CREATE UNIQUE INDEX IdxEmpresa ON PropIVA_TotMensual (IdEmpresa, Ano, Mes) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Tabla Sucursales
      Q1 = "DROP INDEX IdSucursal ON Sucursales "
      Rc = ExecSQL(DbMain, Q1, False)
            
      Q1 = "CREATE UNIQUE INDEX IdSucursal ON Sucursales (IdSucursal) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)
      
      'Agregamos campo Vigente a AreaNegocio
      Set Tbl = DbMain.TableDefs("AreaNegocio")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Vigente", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "AreaNegocio.Vigente", vbExclamation
         lUpdOK = False
      End If
      
      Q1 = "UPDATE AreaNegocio SET Vigente = -1 "
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo Vigente a CentroCosto
      Set Tbl = DbMain.TableDefs("CentroCosto")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Vigente", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "CentroCosto.Vigente", vbExclamation
         lUpdOK = False
      End If
     
      Q1 = "UPDATE CentroCosto SET Vigente = -1 "
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo Vigente a Socios
      Set Tbl = DbMain.TableDefs("Socios")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Vigente", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Socios.Vigente", vbExclamation
         lUpdOK = False
      End If
     
      Q1 = "UPDATE Socios SET Vigente = -1 "
      Call ExecSQL(DbMain, Q1)
      
      DbMain.TableDefs.Refresh
           
      'agregamos IdEmpresa y Año a vista vMovCompIdDoc
      Call DbMain.QueryDefs.Delete("vMovCompIdDoc")
      
      Q1 = "SELECT IdDoc, IdEmpresa, Ano, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber"
      Q1 = Q1 & " FROM MovComprobante GROUP BY IdDoc, IdEmpresa, Ano "
      
      Call DbMain.CreateQueryDef("vMovCompIdDoc", Q1)
     
      'agregamos IdEmpresa y Año a vista vPrimeraDocCuota (vista para seleccionar la primera cuota no pagada de los documentos)
      Call DbMain.QueryDefs.Delete("vPrimeraDocCuota")
      
      Q1 = "SELECT IdDoc, IdEmpresa, Ano, Min(NumCuota) as NumCuota1 FROM DocCuotas WHERE Estado = " & ED_PENDIENTE & " GROUP BY IdDoc, IdEmpresa, Ano ORDER BY IdDoc"
      Call DbMain.CreateQueryDef("vPrimeraDocCuota", Q1)
      
      If lUpdOK Then
         lDbVer = 700
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V365 = lUpdOK

End Function

Public Function CorrigeBase_V364() As Boolean   'agregada 3 may 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Plan As String


   On Error Resume Next

   '--------------------- Versión 364 -----------------------------------

   If lDbVer = 364 And lUpdOK = True Then
      
      'Agregamos campos CodCtaAfectoVta, CodCtaExentoVta, CodCtaTotalVta y Giro a Entidades
      Set Tbl = DbMain.TableDefs("Entidades")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaAfectoVta", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCtaAfectoVta", vbExclamation
         lUpdOK = False
      End If
                 
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaExentoVta", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCtaExentoVta", vbExclamation
         lUpdOK = False
      End If
                 
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaTotalVta", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCtaTotalVta", vbExclamation
         lUpdOK = False
      End If
                 

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("EsDelGiro", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.EsDelGiro", vbExclamation
         lUpdOK = False
      End If
      
                 
      'Agregamos campo CodCCosto y CodAreaNeg para Ventas a Entidades
      Set Tbl = DbMain.TableDefs("Entidades")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCCostoAfectoVta", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCCostoAfectoVta", vbExclamation
         lUpdOK = False
      End If
                 
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodAreaNegAfectoVta", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodAreaNegAfectoVta", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCCostoExentoVta", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCCostoExentoVta", vbExclamation
         lUpdOK = False
      End If
                 
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodAreaNegExentoVta", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodAreaNegExentoVta", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCCostoTotalVta", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCCostoTotalVta", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodAreaNegTotalVta", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodAreaNegTotalVta", vbExclamation
         lUpdOK = False
      End If
      
      Set Tbl = Nothing
      
      'vamos a poner EsDelGiro en Si por omisión, sólo para entidades que lo tienen en NULL
   
      Q1 = "UPDATE Entidades SET EsDelGiro = -1 "
      Call ExecSQL(DbMain, Q1)
   

      
      If lUpdOK Then
         lDbVer = 365
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V364 = lUpdOK

End Function

Public Function CorrigeBase_V363() As Boolean   'agregada 27 feb 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 363 -----------------------------------

   If lDbVer = 363 And lUpdOK = True Then
   
      Q1 = "DROP INDEX IdxItem ON CtasAjustesExCont "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE UNIQUE INDEX Idx ON CtasAjustesExCont (IdCtaAjustes)"
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX IdxItem ON CtasAjustesExCont (TipoAjuste, IdItem)"
      Rc = ExecSQL(DbMain, Q1, False)
      
      If lUpdOK Then
         lDbVer = 364
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V363 = lUpdOK

End Function

Public Function CorrigeBase_V362() As Boolean   'agregada 29 oct 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 362 -----------------------------------

   If lDbVer = 362 And lUpdOK = True Then
   
      'Agregamos campo URLDte a tabla Documento (esto para la importación de documentos desde facturación)
      Set Tbl = DbMain.TableDefs("Documento")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("UrlDTE", dbText, 250)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.UrlDTE", vbExclamation
         lUpdOK = False
      End If
      
      If lUpdOK Then
         lDbVer = 363
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V362 = lUpdOK

End Function
Public Function CorrigeBase_V361() As Boolean   'agregada 17 oct 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 361 -----------------------------------

   If lDbVer = 361 And lUpdOK = True Then
   
      'Actualizamos los códigos de actividad economica de acuerdo a nueva codificación SII de Nov 2018
      Q1 = "UPDATE Empresa LEFT JOIN CodActiv ON Empresa.CodActEconom = CodActiv.OldCodigo SET Empresa.CodActEconom = CodActiv.Codigo, ActEconom = 0"
      Call ExecSQL(DbMain, Q1)
      
      MsgBox1 "ATENCIÓN: El SII cambió la codificación de Actividades Económicas." & vbCrLf & vbCrLf & "Hemos actualizado el Código de Actividad Económica de la empresa seleccionada, de acuerdo a la nueva codificación entregada por el SII, respetando el esquema de conversión entregado por el mismo SII." & vbCrLf & vbCrLf & "Sin embargo, debe revisar que el código asignado se ajuste a la actividad de la empresa.", vbInformation + vbOKOnly
      
      If lUpdOK Then
         lDbVer = 362
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V361 = lUpdOK

End Function
           
Public Function CorrigeBase_V360() As Boolean   'agregada 4 sep 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 360 -----------------------------------

   If lDbVer = 360 And lUpdOK = True Then
      
      'agrandamos campos de Entidades
      Call AlterField(DbMain, "Entidades", "Nombre", dbText, 100)
      Call AlterField(DbMain, "Entidades", "Giro", dbText, 80)
      Call AlterField(DbMain, "Entidades", "Email", dbText, 100)
        
      'Agregamos campo CodCuentaOld a tabla Cuentas (esto para la importación de documentos de año anterior)
      Set Tbl = DbMain.TableDefs("MovDocumento")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCuentaOld", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovDocumento.CodCuentaOld", vbExclamation
         lUpdOK = False
      End If
   
      'habían quedado mal creados
      Q1 = "CREATE UNIQUE INDEX Idx ON CtasAjustesExCont (IdCuentas14TER) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX IdxItem ON CtasAjustesExCont (TipoAjuste, IdItem)"
      Rc = ExecSQL(DbMain, Q1, False)

      If lUpdOK Then
         lDbVer = 361
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V360 = lUpdOK

End Function
            
Public Function CorrigeBase_V359() As Boolean   'agregada 16 ago 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String

      Call AlterField(DbMain, "Entidades", "Nombre", dbText, 100)
      Call AlterField(DbMain, "Entidades", "Giro", dbText, 80)

   On Error Resume Next

   '--------------------- Versión 359 -----------------------------------

   If lDbVer = 359 And lUpdOK = True Then
      
      'Agrandamos campo Cartola a la tabla Cartola
      Call AlterField(DbMain, "LibroCaja", "IdDoc", dbLong)
     
   
      If lUpdOK Then
         lDbVer = 360
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V359 = lUpdOK

End Function

Public Function CorrigeBase_V358() As Boolean   'agregada 4 jun 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Plan As String


   On Error Resume Next

   '--------------------- Versión 358 -----------------------------------

   If lDbVer = 358 And lUpdOK = True Then
      
      'Agregamos campo CodCCosto y CodAreaNeg a Entidades
      Set Tbl = DbMain.TableDefs("Entidades")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCCostoAfecto", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCCostoAfecto", vbExclamation
         lUpdOK = False
      End If
                 
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodAreaNegAfecto", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodAreaNegAfecto", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCCostoExento", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCCostoExento", vbExclamation
         lUpdOK = False
      End If
                 
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodAreaNegExento", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodAreaNegExento", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCCostoTotal", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCCostoTotal", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodAreaNegTotal", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodAreaNegTotal", vbExclamation
         lUpdOK = False
      End If
      
      Set Tbl = Nothing
    
      If lUpdOK Then
         lDbVer = 359
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V358 = lUpdOK

End Function

Public Function CorrigeBase_V357() As Boolean   'agregada 22 may 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Plan As String


   On Error Resume Next

   '--------------------- Versión 357 -----------------------------------

   If lDbVer = 357 And lUpdOK = True Then
      
      'Agregamos campo PropIVA a Entidades
      Set Tbl = DbMain.TableDefs("Entidades")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("PropIVA", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.PropIVA", vbExclamation
         lUpdOK = False
      End If
                 
      Set Tbl = Nothing
    
      If lUpdOK Then
         lDbVer = 358
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V357 = lUpdOK

End Function

Public Function CorrigeBase_V356() As Boolean   'agregada 14 may 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Plan As String


   On Error Resume Next

   '--------------------- Versión 356 -----------------------------------

   If lDbVer = 356 And lUpdOK = True Then
      
      'Agregamos campos CodCtaAfecto, CodCtaExento y CodCtaTotal a Entidades
      Set Tbl = DbMain.TableDefs("Entidades")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaAfecto", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCtaAfecto", vbExclamation
         lUpdOK = False
      End If
                 
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaExento", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCtaExento", vbExclamation
         lUpdOK = False
      End If
                 
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaTotal", dbText, 15)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.CodCtaTotal", vbExclamation
         lUpdOK = False
      End If
                 

      Set Tbl = Nothing
    
    
      If lUpdOK Then
         lDbVer = 357
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V356 = lUpdOK

End Function

Public Function CorrigeBase_V355() As Boolean   'agregada 18 abr 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Plan As String


   On Error Resume Next

   '--------------------- Versión 355 -----------------------------------

   If lDbVer = 355 And lUpdOK = True Then
      
      Q1 = "CREATE UNIQUE INDEX Idx ON Documento (IdDoc) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)
         
      If lUpdOK Then
         lDbVer = 356
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V355 = lUpdOK

End Function

Public Function CorrigeBase_V354() As Boolean   'agregada 5 abr 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Plan As String


   On Error Resume Next

   '--------------------- Versión 354 -----------------------------------

   If lDbVer = 354 And lUpdOK = True Then
      
      'Agregamos campo CodCtaPlanSII a Cuentas
      Set Tbl = DbMain.TableDefs("Cuentas")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodCtaPlanSII", dbText, 10)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.CodCtaPlanSII", vbExclamation
         lUpdOK = False
      End If
                 
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
      
         If vFld(Rs("Valor")) = "BÁSICO" Or vFld(Rs("Valor")) = "INTERMEDIO" Or vFld(Rs("Valor")) = "AVANZADO" Then
            Call UpdateCtaPlanSII("Cuentas")
         
         ElseIf vFld(Rs("Valor")) = "IFRS" Then
            Call UpdateCtaPlanSII_IFRS("Cuentas")
         
         End If
      End If
      
      Call CloseRs(Rs)
      
      'agregamos tabla InfoAnualDJ1847. La tabla tiene un sólo registro
      Set Tbl = New TableDef
      Tbl.Name = "InfoAnualDJ1847"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdEmpresa", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "InfoAnualDJ1847.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
            
      ERR.Clear
      Set Fld = Tbl.CreateField("Ano", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "InfoAnualDJ1847.Ano", vbExclamation
         lUpdOK = False
      End If
            
      ERR.Clear
      Set Fld = Tbl.CreateField("IdEntSupervisora", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "InfoAnualDJ1847.IdEntSupervisora", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("AnoAjusteIFRS", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "InfoAnualDJ1847.AnoAjusteIFRS", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("FolioInicial", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "InfoAnualDJ1847.FolioInicial", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("FolioFinal", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "InfoAnualDJ1847.FolioFinal", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdAjustesRLI", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "InfoAnualDJ1847.IdAjustesRLI", vbExclamation
         lUpdOK = False
      End If
      

      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON InfoAnualDJ1847 (IdEmpresa, Ano) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla AsistImpPrimCat", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing

      
      If lUpdOK Then
         lDbVer = 355
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V354 = lUpdOK

End Function

Public Function CorrigeBase_V353() As Boolean   'agregada 17 ene 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 353 -----------------------------------

   If lDbVer = 353 And lUpdOK = True Then
      
      'Agregamos tabla AsistImpPrimCat
      
      Set Tbl = New TableDef
      Tbl.Name = "AsistImpPrimCat"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdAsistImpPrimCat", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AsistImpPrimCat.IdAsistImpPrimCat", vbExclamation
         lUpdOK = False
      End If
            
      ERR.Clear
      Set Fld = Tbl.CreateField("IdItem", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AsistImpPrimCat.IdItem", vbExclamation
         lUpdOK = False
      End If
            
      ERR.Clear
      Set Fld = Tbl.CreateField("RemEjAntNominal", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AsistImpPrimCat.RemEjAntNominal", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("RemEjAntAct", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AsistImpPrimCat.RemEjAntAct", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("GeneradoAno", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AsistImpPrimCat.GeneradoAno", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("CredUtilizado", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AsistImpPrimCat.CredUtilizado", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("RemEjSgte", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AsistImpPrimCat.RemEjSgte", vbExclamation
         lUpdOK = False
      End If
      

      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON AsistImpPrimCat (IdAsistImpPrimCat) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)

         Q1 = "CREATE UNIQUE INDEX IdxItem ON AsistImpPrimCat (IdItem)"
         Rc = ExecSQL(DbMain, Q1, False)

      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla AsistImpPrimCat", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing
      
      If gEmpresa.Ano = 2016 Then
         Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('IMP1CAT',0,'0.24')"
      ElseIf gEmpresa.Ano = 2017 Then
         Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('IMP1CAT',0,'0.25')"
      Else
         Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('IMP1CAT',0,'0.25')"
      End If
      
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVer = 354
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V353 = lUpdOK

End Function


Public Function CorrigeBase_V352() As Boolean   'Agregada 10 ene 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 352 -----------------------------------
   If lDbVer = 352 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("Empresa")
      
      'agregamos campo FranqRentaAtribuida
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqRentaAtribuida", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqRentaAtribuida", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo FranqSemiIntegrado
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqSemiIntegrado", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqSemiIntegrado", vbExclamation
         lUpdOK = False
      End If
    
      Set Tbl = DbMain.TableDefs("Documento")
      
      'agregamos campo IdANegCCosto
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdANegCCosto", dbText, 20)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.IdANegCCosto", vbExclamation
         lUpdOK = False
      End If
    
           
      If lUpdOK Then
         lDbVer = 353
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V352 = lUpdOK

End Function


Public Function CorrigeBase_V351() As Boolean   'agregada 4 ene 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 351 -----------------------------------

   If lDbVer = 351 And lUpdOK = True Then
      
      'Agregamos tabla AjustesExtraLibCaja
      
      Set Tbl = New TableDef
      Tbl.Name = "BaseImponible14Ter"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdBaseImponible14Ter", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14Ter.IdBaseImponible14Ter", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("TipoBaseImp", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14Ter.TipoBaseImp", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdItemBaseImp", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14Ter.IdItemBaseImp", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Valor", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "BaseImponible14Ter.Valor", vbExclamation
         lUpdOK = False
      End If

      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON BaseImponible14Ter (IdBaseImponible14Ter) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)

         Q1 = "CREATE UNIQUE INDEX IdxItem ON BaseImponible14Ter (TipoBaseImp, IdItemBaseImp)"
         Rc = ExecSQL(DbMain, Q1, False)

      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla BaseImponible14Ter", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing
      
   
      If lUpdOK Then
         lDbVer = 352
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V351 = lUpdOK

End Function

            
Public Function CorrigeBase_V350() As Boolean   'agregada 27 dic 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Plan As String


   On Error Resume Next

   '--------------------- Versión 350 -----------------------------------

   If lDbVer = 350 And lUpdOK = True Then
      
      'Agregamos campo TipoPartida a Cuentas
      Set Tbl = DbMain.TableDefs("Cuentas")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoPartida", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.TipoPartida", vbExclamation
         lUpdOK = False
      End If
                 
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
      
         If vFld(Rs("Valor")) = "BÁSICO" Or vFld(Rs("Valor")) = "INTERMEDIO" Or vFld(Rs("Valor")) = "AVANZADO" Then
            Call UpdateTipoPartidaCtas("Cuentas")
         
         ElseIf vFld(Rs("Valor")) = "IFRS" Then
            Call UpdateTipoPartidaIFRS("Cuentas")
         
         End If
      End If
      
      Call CloseRs(Rs)
      
      If gEmpresa.Ano >= 2017 Then
         Q1 = "UPDATE Cuentas SET CodF22 = '' WHERE CodF22 NOT IN (" & LSTCODF22_2017 & ")"     'se eliminan códigos F22 ya no válidos desde 2017
         Call ExecSQL(DbMain, Q1)
      End If
      
      If lUpdOK Then
         lDbVer = 351
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V350 = lUpdOK

End Function
Public Function CorrigeBase_V349() As Boolean   'agregada 14 dic 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 349 -----------------------------------

   If lDbVer = 349 And lUpdOK = True Then
      
      'Agregamos tabla AjustesExtraLibCaja
      
      Set Tbl = New TableDef
      Tbl.Name = "AjustesExtLibCaja"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdAjustesExtLibCaja", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AjustesExtLibCaja.IdAjustesExtLibCaja", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("TipoAjuste", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AjustesExtLibCaja.TipoAjuste", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdItemAjuste", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AjustesExtLibCaja.IdItemAjuste", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Valor", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AjustesExtLibCaja.Valor", vbExclamation
         lUpdOK = False
      End If

      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON AjustesExtLibCaja (IdAjustesExtLibCaja) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)

         Q1 = "CREATE UNIQUE INDEX IdxItem ON AjustesExtLibCaja (TipoAjuste, IdItemAjuste)"
         Rc = ExecSQL(DbMain, Q1, False)

      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla AjustesExtLibCaja", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing
      
   
      If lUpdOK Then
         lDbVer = 350
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V349 = lUpdOK

End Function
            
Public Function CorrigeBase_V348() As Boolean   'agregada 9 nov 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 348 -----------------------------------

   If lDbVer = 348 And lUpdOK = True Then
      
      'Agregamos campo NumDocAsoc y DTEDocAsoc a Documento
      Set Tbl = DbMain.TableDefs("Documento")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NumDocAsoc", dbText, 20)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.NumDocAsoc", vbExclamation
         lUpdOK = False
      End If
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DTEDocAsoc", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.DTEDocAsoc", vbExclamation
         lUpdOK = False
      End If
     
   
      If lUpdOK Then
         lDbVer = 349
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V348 = lUpdOK

End Function
            
Public Function CorrigeBase_V347() As Boolean   'agregada 12 sept 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 347 -----------------------------------

   If lDbVer = 347 And lUpdOK = True Then
      
      'Agrandamos campo Cartola a la tabla Cartola
      Call AlterField(DbMain, "Cartola", "Cartola", dbInteger)
     
   
      If lUpdOK Then
         lDbVer = 348
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V347 = lUpdOK

End Function

Public Function CorrigeBase_V346() As Boolean   'agregada 5 sept 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 346 -----------------------------------

   If lDbVer = 346 And lUpdOK = True Then
      
      'Agregamos Campo Ingreso a tabla LibCaja
      Set Tbl = DbMain.TableDefs("LibroCaja")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Ingreso", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LibroCaja.Ingreso", vbExclamation
         lUpdOK = False
      End If
   
      'Agregamos Campo Egreso a tabla LibCaja
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Egreso", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LibroCaja.Egreso", vbExclamation
         lUpdOK = False
      End If
   
      If lUpdOK Then
         lDbVer = 347
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V346 = lUpdOK

End Function
            

Public Function CorrigeBase_V345() As Boolean   'agregada 30 ago 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 345 -----------------------------------

   If lDbVer = 345 And lUpdOK = True Then
   
      Q1 = "DROP INDEX IdxDoc ON LibroCaja "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX IdxDoc ON LibroCaja (IdDoc)"
      Rc = ExecSQL(DbMain, Q1, False)
   
      Q1 = "CREATE INDEX IdxFecha ON LibroCaja (FechaIngresoLibro)"
      Rc = ExecSQL(DbMain, Q1, False)
   
      'Agregamos Campo CompraBienRaiz a tabla Documento
      Set Tbl = DbMain.TableDefs("Documento")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CompraBienRaiz", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.CompraBienRaiz", vbExclamation
         lUpdOK = False
      End If
   
      If lUpdOK Then
         lDbVer = 346
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V345 = lUpdOK

End Function
            

Public Function CorrigeBase_V344() As Boolean   'agregada 16 ago 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 344 -----------------------------------

   If lDbVer = 344 And lUpdOK = True Then

      'Agregamos Campo OtrosIngEg14TER a Comprobante para indicar que debemos importar este comprobante como "Otros Ingresos y Egresos" a Libro de Caja
      Set Tbl = DbMain.TableDefs("Comprobante")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("OtrosIngEg14TER", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Comprobante.OtrosIngEg14TER", vbExclamation
         lUpdOK = False
      End If
   
      Set Tbl = DbMain.TableDefs("CT_Comprobante")
     
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("OtrosIngEg14TER", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "OtrosIngEg14TER", vbExclamation
         lUpdOK = False
      End If
   
      If lUpdOK Then
         lDbVer = 345
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V344 = lUpdOK

End Function
            

Public Function CorrigeBase_V343() As Boolean   'agregada 2 ago 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 343 -----------------------------------

   If lDbVer = 343 And lUpdOK = True Then

      'Agregamos tabla CtasAjustesExCont
      
      Set Tbl = New TableDef
      Tbl.Name = "CtasAjustesExCont"

      ERR.Clear
      Set Fld = Tbl.CreateField("IdCtaAjustes", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExCont.IdCuentas14TER", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("TipoAjuste", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExCont.TipoAjuste", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdItem", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExCont.IdItem", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdCuenta", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExCont.IdCuenta", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("CodCuenta", dbText, 15)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CtasAjustesExCont.CodCuenta", vbExclamation
         lUpdOK = False
      End If

      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON CtasAjustesExCont (IdCuentas14TER) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)

         Q1 = "CREATE UNIQUE INDEX IdxItem ON CtasAjustesExCont (TipoAjuste, IdItem)"
         Rc = ExecSQL(DbMain, Q1, False)

      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla CtasAjustesExCont", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing
      
      'agregamos campo IdComp a LibroCaja para identificar comprobante  asociado a otros ingresos y egresos
      Set Tbl = DbMain.TableDefs("LibroCaja")
         
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdComp", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LibroCaja.IdComp", vbExclamation
         lUpdOK = False
      End If

   
      If lUpdOK Then
         lDbVer = 344
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V343 = lUpdOK

End Function
            

Public Function CorrigeBase_V342() As Boolean   'agregada 20 jul 2017, entregada 1 ago 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 342 -----------------------------------

   If lDbVer = 342 And lUpdOK = True Then

      Set Tbl = DbMain.TableDefs("Documento")
      
      'agregamos campo NumCuotas a Documento para registrar que el documento es venta con cuotas, a credito
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NumCuotas", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.NumCuotas", vbExclamation
         lUpdOK = False
      End If

      Set Tbl = DbMain.TableDefs("DocCuotas")
      
      'agregamos campo Estado a DocCuotas para registrar si la cuota está pagada
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Estado", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "DocCuotas.Estado", vbExclamation
         lUpdOK = False
      End If
            
      'agregamos campo IdDocCuota a MovComprobante para registrar  la cuota si hay
      
      Set Tbl = DbMain.TableDefs("MovComprobante")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdDocCuota", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovComprobante.IdDocCuota", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos una vista para seleccionar la primera cuota no pagada de los documentos
      Q1 = "SELECT IdDoc, Min(NumCuota) as NumCuota1 FROM DocCuotas WHERE Estado = " & ED_PENDIENTE & " GROUP BY IdDoc ORDER BY IdDoc"
      Call DbMain.CreateQueryDef("vPrimeraDocCuota", Q1)
      
      'agregamos una vista para seleccionar la última cuota de los documentos (pagada o no pagada)
'      Q1 = "SELECT IdDoc, Max(NumCuota) as NumCuota1 FROM DocCuotas GROUP BY IdDoc ORDER BY IdDoc"
'      Call DbMain.CreateQueryDef("vUltimaDocCuota", Q1)
      
      

      If lUpdOK Then
         lDbVer = 343
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V342 = lUpdOK

End Function
            

Public Function CorrigeBase_V341() As Boolean   'agregada 11 jul 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 341 -----------------------------------

   If lDbVer = 341 And lUpdOK = True Then

      Set Tbl = DbMain.TableDefs("DocCuotas")
      Set Fld = Tbl.Fields(0)
      Fld.Name = "IdDocCuota"
      
      Q1 = "CREATE UNIQUE INDEX Idx ON DocCuotas (IdDocCuota) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "DROP INDEX IdxIdDoc ON DocCuotas "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX IdxIdDoc ON DocCuotas (IdDoc)"
      Rc = ExecSQL(DbMain, Q1, False)

      If lUpdOK Then
         lDbVer = 342
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V341 = lUpdOK

End Function
            

Public Function CorrigeBase_V340() As Boolean   'agregada 6 jul 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 340 -----------------------------------

   If lDbVer = 340 And lUpdOK = True Then

      'Agregamos tabla DocCuotas
      
      Set Tbl = New TableDef
      Tbl.Name = "DocCuotas"

      ERR.Clear
      Set Fld = Tbl.CreateField("IdDocCuotas", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DocCuotas.IdDocCuotas", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdDoc", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DocCuotas.IdDoc", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("NumCuota", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DocCuotas.NumCuota", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("FechaExigPago", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DocCuotas.FechaExigPago", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("MontoCuota", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DocCuotas.MontoCuota", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("FechaIngPercibido", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DocCuotas.FechaIngPercibido", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdCompPago", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DocCuotas.IdCompPago", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdLibCaja", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DocCuotas.IdLibCaja", vbExclamation
         lUpdOK = False
      End If


      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON DocCuotas (IdDocCuotas) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)

         Q1 = "CREATE UNIQUE INDEX IdxIdDoc ON DocCuotas (IdDoc)"
         Rc = ExecSQL(DbMain, Q1, False)

      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla DocCuotas", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing
   
   
   
      If lUpdOK Then
         lDbVer = 341
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V340 = lUpdOK

End Function


Public Function CorrigeBase_V339() As Boolean   'entregada 7 jul 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 339 -----------------------------------

   If lDbVer = 339 And lUpdOK = True Then

      Set Tbl = DbMain.TableDefs("Documento")
      
      'agregamos campo EntRelacionada a Documento para registrar que el documento es venta con una entidad relacionada
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("EntRelacionada", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.EntRelacionada", vbExclamation
         lUpdOK = False
      End If
      
      'Asignamos el campo EntRelacionada de los Documentos con el campo EntRelacionada de la IdEntidad del documento original, para los documentos ya ingresados.
      Q1 = "UPDATE Documento INNER JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad SET Documento.EntRelacionada = Entidades.EntRelacionada"
      Call ExecSQL(DbMain, Q1)
      
      
      Set Tbl = DbMain.TableDefs("LibroCaja")
      
      'agregamos campo IdEntReal al Libro de Caja para registrar que el ID de la entidad a la cual se le hizo realmente la venta
      'dado que en el caso de los ingresos, el RUT es el de la empresa emisora (no el que está en el libro de ventas), salvo en el caso de (FCV) y (LFV) y las notas de crédito/débito asociadas a dichos documentos.
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdEntReal", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LibroCaja.IdEntReal", vbExclamation
         lUpdOK = False
      End If


      Set Tbl = Nothing
      
      'Asignamos al campo IdEntReal del Libro de Caja el campo IdEntidad del documento, para los documentos ya ingresados.
      Q1 = "UPDATE LibroCaja INNER JOIN Documento ON LibroCaja.IdDoc = Documento.IdDoc SET LibroCaja.IdEntReal = Documento.IdEntidad"
      Call ExecSQL(DbMain, Q1)
   
      'Asignamos el campo ConEntRel del Libro de Caja con el campo EntRelacionada de la IdEntidad del documento original, para los documentos ya ingresados.
      Q1 = "UPDATE LibroCaja INNER JOIN Entidades ON LibroCaja.IdEntReal = Entidades.IdEntidad SET LibroCaja.ConEntRel = Entidades.EntRelacionada"
      Call ExecSQL(DbMain, Q1)
   
      If lUpdOK Then
         lDbVer = 340
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V339 = lUpdOK

End Function


Public Function CorrigeBase_V338() As Boolean   'agregada 15 mayo 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 338 -----------------------------------

   If lDbVer = 338 And lUpdOK = True Then

      Set Tbl = DbMain.TableDefs("LibroCaja")
      
      'agregamos campo FechaIngresoLibro a LibroCaja para almacenar la fecha en que se registró el documento en el libro de caja
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FechaIngresoLibro", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LibroCaja.FechaIngresoLibro", vbExclamation
         lUpdOK = False
      End If


      Set Tbl = Nothing
      
      'Asignamos a la Fecha de Ingreso Libro la fecha de la Operación, para los documentos ya ingresados.
      Q1 = "UPDATE LibroCaja SET FechaIngresoLibro = FechaOperacion"
      Call ExecSQL(DbMain, Q1)
   
      If lUpdOK Then
         lDbVer = 339
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V338 = lUpdOK

End Function


Public Function CorrigeBase_V337() As Boolean   'agregada 24 abril 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 337 -----------------------------------

   If lDbVer = 337 And lUpdOK = True Then

      Set Tbl = DbMain.TableDefs("Cuentas")
      
      'agregamos campo CodF22_14Ter a tabla Cuentas
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF22_14Ter", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.CodF22_14Ter", vbExclamation
         lUpdOK = False
      End If


      Set Tbl = Nothing
      
   
      'Agregamos tabla Socios
      
      Set Tbl = New TableDef
      Tbl.Name = "Socios"

      ERR.Clear
      Set Fld = Tbl.CreateField("IdSocio", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Socios.IdSocio", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("RUT", dbText, 12)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Socios.RUT", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Nombre", dbText, 50)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Socios.Nombre", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("PjePart", dbSingle)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Socios.PjePart", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("MontoSuscrito", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Socios.MontoSuscrito", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("MontoPagado", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Socios.MontoPagado", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdCuentaAportes", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Socios.IdCuentaAportes", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdCuentaRetiros", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Socios.IdCuentaRetiros", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdTipoSocio", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Socios.IdTipoSocio", vbExclamation
         lUpdOK = False
      End If


      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON Socios (IdSocio) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)

         Q1 = "CREATE UNIQUE INDEX IdxRut ON Socios (RUT)"
         Rc = ExecSQL(DbMain, Q1, False)

      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla Socios", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing
   
   
   
      If lUpdOK Then
         lDbVer = 338
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V337 = lUpdOK

End Function


Public Function CorrigeBase_V336() As Boolean   'agregada 13 abril 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 336 -----------------------------------

   If lDbVer = 336 And lUpdOK = True Then

      Set Tbl = DbMain.TableDefs("Entidades")
      
      'agregamos campo EntRelacionada a tabla Entidades
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("EntRelacionada", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.EntRelacionada", vbExclamation
         lUpdOK = False
      End If


      Set Tbl = Nothing
      
   
      If lUpdOK Then
         lDbVer = 337
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V336 = lUpdOK

End Function


Public Function CorrigeBase_V335() As Boolean   'agregada 11 abril 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 335 -----------------------------------

   If lDbVer = 335 And lUpdOK = True Then

      Set Tbl = DbMain.TableDefs("LibroCaja")
      
      'agregamos campo IVAIrrec a tabla LibroCaja
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IVAIrrec", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LibroCaja.IVAIrrec", vbExclamation
         lUpdOK = False
      End If


      Set Tbl = Nothing
      
   
      If lUpdOK Then
         lDbVer = 336
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V335 = lUpdOK

End Function


Public Function CorrigeBase_V334() As Boolean   'agregada 3 mar 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 334 -----------------------------------

   If lDbVer = 334 And lUpdOK = True Then

      'agregamos entidad especial fija para Formulario de Importaciones
      Q1 = "INSERT INTO Entidades( RUT, Codigo, Nombre, Clasif" & ENT_PROVEEDOR & ")VALUES('55555555', 'DIN', 'DIN', 1)"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVer = 335
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V334 = lUpdOK

End Function


Public Function CorrigeBase_V333() As Boolean   'agregada 1 dic 2016
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 333 -----------------------------------

   If lDbVer = 333 And lUpdOK = True Then

      'agregamos campo EsRecuperable a tabla ImpAdic, para almacenar Si Es Recuperable en cada impuesto adicional
      Set Tbl = DbMain.TableDefs("ImpAdic")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("EsRecuperable", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "ImpAdic.EsRecuperable", vbExclamation
         lUpdOK = False
      End If

            
      Set Tbl = DbMain.TableDefs("Documento")
      'agregamos campo CodSIIDTEIVAIrrec a tabla Documento, para almacenar código SII del IVA Irrecuperable
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodSIIDTEIVAIrrec", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.CodSIIDTEIVAIrrec", vbExclamation
         lUpdOK = False
      End If

      'agregamos campo TipoDocAsoc a tabla Documento, para almacenar tipo doc y saber si es una factura de compra. Esto para libro electrónico de compras
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoDocAsoc", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.TipoDocAsoc", vbExclamation
         lUpdOK = False
      End If

      'agregamos campo IVAActFijo a tabla Documento, para almacenar el valor del IVA Activo fijo del documento
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IVAActFijo", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.IVAActFijo", vbExclamation
         lUpdOK = False
      End If

      Set Tbl = Nothing
      
      Set Tbl = DbMain.TableDefs("MovDocumento")
      
      'agregamos campo Tasa a tabla MovDocumento, para almacenar Tasa en cada impuesto adicional
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Tasa", dbSingle)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovDocumento.Tasa", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo EsRecuperable a tabla MovDocumento, para almacenar Si Es Recuperable en cada impuesto adicional
            
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("EsRecuperable", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovDocumento.EsRecuperable", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo CodSIIDTE a tabla MovDocumento, para almacenar el código SIIDTE en cada impuesto adicional
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodSIIDTE", dbText, 2)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovDocumento.CodSIIDTE", vbExclamation
         lUpdOK = False
      End If

      Set Tbl = Nothing
      If lUpdOK Then
         lDbVer = 334
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V333 = lUpdOK

End Function


Public Function CorrigeBase_V332() As Boolean   'agregada 16 nov
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 332 -----------------------------------

   If lDbVer = 332 And lUpdOK = True Then

      Set Tbl = DbMain.TableDefs("ImpAdic")
      
      'agregamos campo Tasa a tabla ImpAdic, para almacenar Tasa de la empresa para ese impuesto
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Tasa", dbSingle)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "ImpAdic.Tasa", vbExclamation
         lUpdOK = False
      End If


      Set Tbl = Nothing
      

      Call UpdImpAdicionales2016
   
      If lUpdOK Then
         lDbVer = 333
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V332 = lUpdOK

End Function


Public Function CorrigeBase_V331() As Boolean   'agregada 26 oct 16 - entregada
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 331 -----------------------------------

   If lDbVer = 331 And lUpdOK = True Then

      'Agregamos ImpAdic
      Set Tbl = New TableDef
      Tbl.Name = "ImpAdic"

      ERR.Clear
      Set Fld = Tbl.CreateField("IdImpAdic", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ImpAdic.IdImpAdic", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("TipoLib", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ImpAdic.TipoLib", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("TipoValor", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ImpAdic.TipoValor", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdCuenta", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ImpAdic.IdCuenta", vbExclamation
         lUpdOK = False
      End If

      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON ImpAdic (IdImpAdic) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)

         Q1 = "CREATE UNIQUE INDEX IdxTipo ON ImpAdic (TipoLib, TipoValor)"
         Rc = ExecSQL(DbMain, Q1, False)

      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla ImpAdic", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing


      If lUpdOK Then
         lDbVer = 332
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V331 = lUpdOK

End Function

Public Function CorrigeBase_V330() As Boolean   'abierta 5 oct 2016
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
  

   On Error Resume Next
   
   '--------------------- Versión 330 -----------------------------------
   If lDbVer = 330 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("Entidades")
      
      'agregamos campo IdEmpresa a tabla entidades (para importación desde facturación)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdEmpresa", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
                            
      Set Tbl = DbMain.TableDefs("Documento")
      
      'agregamos campo IdEmpresa a tabla Documento (para importación desde facturación)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdEmpresa", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
                                                      
      'agregamos campo FImpFacturacion a tabla Documento (para indicar fecha de importación desde facturación)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FImpFacturacion", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.FImpFacturacion", vbExclamation
         lUpdOK = False
      End If
      
      If lUpdOK Then
         lDbVer = 331
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V330 = lUpdOK

End Function

Public Function CorrigeBase_V329() As Boolean   'agregada 25 jul 2016 - entregada
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
  

   On Error Resume Next
   
   '--------------------- Versión 329 -----------------------------------
   If lDbVer = 329 And lUpdOK = True Then
   
      ERR.Clear
      
      Call AlterField(DbMain, "Entidades", "CodActEcon", dbText, 8)
                                        
      If lUpdOK Then
         lDbVer = 330
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V329 = lUpdOK

End Function

Public Function CorrigeBase_V328() As Boolean   'entregada 4 ene 2016
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
  

   On Error Resume Next
   
   '--------------------- Versión 328 -----------------------------------
   If lDbVer = 328 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("Documento")
      
      'agregamos campo ObligaLibComprasVentas
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IVAInmueble", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.IVAInmueble", vbExclamation
         lUpdOK = False
      End If
                            
      If lUpdOK Then
         lDbVer = 329
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V328 = lUpdOK

End Function

Public Function CorrigeBase_V327() As Boolean   'entregada 4 dic 2015
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
  

   On Error Resume Next
   
   '--------------------- Versión 327 -----------------------------------
   If lDbVer = 327 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("Empresa")
      
      'agregamos campo ObligaLibComprasVentas
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("ObligaLibComprasVentas", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.ObligaLibComprasVentas", vbExclamation
         lUpdOK = False
      End If
                 
      Q1 = "DROP INDEX IdxDoc ON LibroCaja "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE UNIQUE INDEX IdxDoc ON LibroCaja (IdDoc)"
      Rc = ExecSQL(DbMain, Q1, False)
           
      If lUpdOK Then
         lDbVer = 328
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V327 = lUpdOK

End Function


Public Function CorrigeBase_V326() As Boolean   'entregada 19/10/15
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 326 -----------------------------------

   If lDbVer = 326 And lUpdOK = True Then

      'Agregamos Tabla LibroCaja
      Set Tbl = New TableDef
      Tbl.Name = "LibroCaja"

      ERR.Clear
      Set Fld = Tbl.CreateField("IdLibroCaja", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.IdLibroCaja", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdDoc", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.IdDoc", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("TipoOper", dbInteger)         'Ingreso o Egreso
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.TipoOper", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("TipoDoc", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.IdTipoDoc", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("TipoLib", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.IdTipoLib", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("NumDoc", dbText, 20)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.NumDoc", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("NumDocHasta", dbText, 20)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.NumDocHasta", vbExclamation
         lUpdOK = False
      End If


      ERR.Clear
      Set Fld = Tbl.CreateField("DTE", dbBoolean)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.DTE", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdEntidad", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.IdEntidad", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("RutEntidad", dbText, 12)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.RutEntidad", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("NombreEntidad", dbText, 40)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.NombreEntidad", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("FechaOperacion", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.FechaOperacion", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Afecto", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.Neto", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IVA", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.IVA", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Exento", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.Exento", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("OtroImp", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.Exento", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Total", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.Total", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Pagado", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.Pagado", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Descrip", dbText, 50)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.Glosa", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("ConEntRel", dbBoolean)        'Con entidad relacionada
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.ConEntRel", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("OperDevengada", dbBoolean)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.OperDevengada", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("PagoAPlazo", dbBoolean)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.PagoAPlazo", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("FechaExigPago", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.FechaExigPago", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("Estado", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.Estado", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdUsuario", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.IdUsuario", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("FechaCreacion", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LibroCaja.FechaCreacion", vbExclamation
         lUpdOK = False
      End If


      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh

         Q1 = "CREATE UNIQUE INDEX Idx ON LibroCaja (IdLibroCaja) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)

         Q1 = "CREATE INDEX IdxDoc ON LibroCaja (IdDoc)"
         Rc = ExecSQL(DbMain, Q1, False)

         Q1 = "CREATE INDEX IdxNumDoc ON LibroCaja (TipoDoc, TipoLib, IdEntidad, NumDoc)"
         Rc = ExecSQL(DbMain, Q1, False)

      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla LibroCaja", vbExclamation
         lUpdOK = False

      End If

      Set Tbl = Nothing


      If lUpdOK Then
         lDbVer = 327
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V326 = lUpdOK

End Function

Public Function CorrigeBase_V325() As Boolean   'entregada 8/07/15
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
  

   On Error Resume Next
   
   '--------------------- Versión 325 -----------------------------------
   
   If lDbVer = 325 And lUpdOK = True Then
   
      'agregamos campo EsSupermercado a tabla Entidades
      Set Tbl = DbMain.TableDefs("Entidades")
       
      ERR.Clear
      Set Fld = Tbl.CreateField("EsSupermercado", dbBoolean)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Entidades.EsSupermercado", vbExclamation
         lUpdOK = False
      End If
            
      If lUpdOK Then
         lDbVer = 326
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V325 = lUpdOK

End Function

Public Function CorrigeBase_V324() As Boolean   'entregada 29/01/2015
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
  

   On Error Resume Next
   
   '--------------------- Versión 324 -----------------------------------
   
   'agregamos campos de arrastre a año siguiente a tabla ActFijoCompsFicha
   If lDbVer = 324 And lUpdOK = True Then
   
      Set Tbl = DbMain.TableDefs("ActFijoCompsFicha")
       
      'agregamos campos para el reporte IFRS
      ERR.Clear
      Set Fld = Tbl.CreateField("DepPeriodo", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.DepPeriodo", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("Factor", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.Factor", vbExclamation
         lUpdOK = False
      End If
        
      ERR.Clear
      Set Fld = Tbl.CreateField("Revalorizacion", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.Revalorizacion", vbExclamation
         lUpdOK = False
      End If
        
      If lUpdOK Then
         lDbVer = 325
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V324 = lUpdOK

End Function

Public Function CorrigeBase_V323() As Boolean   'entregada dic 2014
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
  

   On Error Resume Next
   
   '--------------------- Versión 323 -----------------------------------
   
   'agregamos campos de arrastre a año siguiente a tabla ActFijoCompsFicha
   If lDbVer = 323 And lUpdOK = True Then
   
      Set Tbl = DbMain.TableDefs("ActFijoCompsFicha")
       
      'agregamos campos para el reporte IFRS
      ERR.Clear
      Set Fld = Tbl.CreateField("NoExisteValRazonable", dbBoolean)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.NoExisteValRazonable", vbExclamation
         lUpdOK = False
      End If
        
      ERR.Clear
      Set Fld = Tbl.CreateField("OtrasDiferencias", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.OtrasDiferencias", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("DepAcum", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.DepAcum", vbExclamation
         lUpdOK = False
      End If
        
      ERR.Clear
      Set Fld = Tbl.CreateField("VidaUtilDep", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.VidaUtilDep", vbExclamation
         lUpdOK = False
      End If
        
      ERR.Clear
      Set Fld = Tbl.CreateField("ReservaAcum", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.ReservaAcum", vbExclamation
         lUpdOK = False
      End If
        
      ERR.Clear
      Set Fld = Tbl.CreateField("DepAcumuladaAnoAnt", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.DepAcumuladaAnoAnt", vbExclamation
         lUpdOK = False
      End If
        
      ERR.Clear
      Set Fld = Tbl.CreateField("VidaUtilYaDep", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.VidaUtilYaDep", vbExclamation
         lUpdOK = False
      End If
        
      ERR.Clear
      Set Fld = Tbl.CreateField("ReservaAcumAnt", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.ReservaAcumAnt", vbExclamation
         lUpdOK = False
      End If
             
      ERR.Clear
      Set Fld = Tbl.CreateField("IdCompFichaOldTmp", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.IdCompFichaOldTmp", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdCompFichaOld", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.IdCompFichaOld", vbExclamation
         lUpdOK = False
      End If
             
      Set Tbl = Nothing
  
      Set Tbl = DbMain.TableDefs("ActFijoFicha")
       
      'agregamos campo IdFichaOld a tabla ActFijoFicha
      ERR.Clear
      Set Fld = Tbl.CreateField("IdFichaOldTmp", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.IdFichaOldTmp", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdFichaOld", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.IdFichaOld", vbExclamation
         lUpdOK = False
      End If
  
      If lUpdOK Then
         lDbVer = 324
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V323 = lUpdOK

End Function

Public Function CorrigeBase_V322() As Boolean   'entregada 23  oct 2014
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
  

   On Error Resume Next
   
   '--------------------- Versión 322 -----------------------------------
   If lDbVer = 322 And lUpdOK = True Then
   
      'Agregamos campos de Dep. Instantánea, DécimaParte y VidaUtilAnos a tabla MovActivoFijo
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
       
      ERR.Clear
      Set Fld = Tbl.CreateField("DepInstant", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepInstant", vbExclamation
         lUpdOK = False
      End If
        
      ERR.Clear
      Set Fld = Tbl.CreateField("DepDecimaParte", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepDecimaParte", vbExclamation
         lUpdOK = False
      End If
        
      ERR.Clear
      Set Fld = Tbl.CreateField("DepInstantHist", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepInstantHist", vbExclamation
         lUpdOK = False
      End If
        
      ERR.Clear
      Set Fld = Tbl.CreateField("DepDecimaParteHist", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepDecimaParteHist", vbExclamation
         lUpdOK = False
      End If
        
      ERR.Clear
      Set Fld = Tbl.CreateField("VidaUtilAnos", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.VidaUtilAnos", vbExclamation
         lUpdOK = False
      End If
       
      Set Tbl = Nothing
  
      If lUpdOK Then
         lDbVer = 323
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V322 = lUpdOK

End Function

Public Function CorrigeBase_V321() As Boolean   'entregada 4 sept 2014
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
  

   On Error Resume Next
   
   '--------------------- Versión 321 -----------------------------------
   If lDbVer = 321 And lUpdOK = True Then
   
      'Agregamos campo FechaImportFile a tabla MovActivoFijo
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
       
      ERR.Clear
      Set Fld = Tbl.CreateField("FechaImportFile", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.FechaImport", vbExclamation
         lUpdOK = False
      End If
        
      Set Tbl = Nothing
  
      If lUpdOK Then
         lDbVer = 322
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V321 = lUpdOK

End Function

Public Function CorrigeBase_V320() As Boolean   'entregada 21 ago 2014
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
  

   On Error Resume Next
   
   '--------------------- Versión 320 -----------------------------------
   If lDbVer = 320 And lUpdOK = True Then
   
      'Agregamos campo SinDetComps en ActFijoFicha
      Set Tbl = DbMain.TableDefs("ActFijoFicha")
       
      ERR.Clear
      Set Fld = Tbl.CreateField("SinDetComps", dbBoolean)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.SinDetComps", vbExclamation
         lUpdOK = False
      End If
        
      Set Tbl = Nothing
  
      If lUpdOK Then
         lDbVer = 321
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V320 = lUpdOK

End Function

Public Function CorrigeBase_V319() As Boolean   'entregada 20 jun 2014
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
  

   On Error Resume Next
   
   '--------------------- Versión 319 -----------------------------------
   If lDbVer = 319 And lUpdOK = True Then
   
      'Agregamos campo Adq_OtrosConceptos y Gast_OtrosConceptos en ActFijoFicha
      Set Tbl = DbMain.TableDefs("ActFijoFicha")
       
      ERR.Clear
      Set Fld = Tbl.CreateField("AdquiOtrosConceptos", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.AdquiOtrosConceptos", vbExclamation
         lUpdOK = False
      End If
    
      ERR.Clear
      Set Fld = Tbl.CreateField("GastoOtrosConceptos", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.GastoOtrosConceptos", vbExclamation
         lUpdOK = False
      End If
    
      Set Tbl = Nothing
  
      If lUpdOK Then
         lDbVer = 320
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V319 = lUpdOK

End Function

Public Function CorrigeBase_V318() As Boolean   'entregada 7 jun 2014
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
  

   On Error Resume Next
   
   '--------------------- Versión 318 -----------------------------------
   If lDbVer = 318 And lUpdOK = True Then
   
      'Agregamos Tabla AFGrupos
      Set Tbl = New TableDef
      Tbl.Name = "AFGrupos"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdGrupo", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AFGrupos.IdGrupo", vbExclamation
         lUpdOK = False
      End If
            
      ERR.Clear
      Set Fld = Tbl.CreateField("NombGrupo", dbText, 15)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AFGrupos.NombGrupo", vbExclamation
         lUpdOK = False
      End If
      
      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX IdxGrupo ON AFGrupos (IdGrupo) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxNombre ON AFGrupos (NombGrupo)"
         Rc = ExecSQL(DbMain, Q1, False)
                  
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla AFGrupos", vbExclamation
         lUpdOK = False
         
      End If
      
      Set Tbl = Nothing
      
      Q1 = "INSERT INTO AFGrupos (NombGrupo) VALUES ('Maquinaria')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO AFGrupos (NombGrupo) VALUES ('Vehículos')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO AFGrupos (NombGrupo) VALUES ('Muebles')"
      Call ExecSQL(DbMain, Q1)
      
      
      'Agregamos Tabla AFComponentes
      Set Tbl = New TableDef
      Tbl.Name = "AFComponentes"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdComp", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AFComponentes.IdComp", vbExclamation
         lUpdOK = False
      End If
            
      ERR.Clear
      Set Fld = Tbl.CreateField("IdGrupo", dbLong)   'padre
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AFComponentes.IdGrupo", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("NombComp", dbText, 30)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "AFComponentes.NombComp", vbExclamation
         lUpdOK = False
      End If
      
      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX IdxComp ON AFComponentes (IdComp) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxNombre ON AFComponentes (NombComp)"
         Rc = ExecSQL(DbMain, Q1, False)
                 
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla AFComponentes", vbExclamation
         lUpdOK = False
         
      End If
      
      Set Tbl = Nothing
     
  
      'Agregamos Tabla ActFijoFicha
      Set Tbl = New TableDef
      Tbl.Name = "ActFijoFicha"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdFicha", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.IdFicha", vbExclamation
         lUpdOK = False
      End If
            
      ERR.Clear
      Set Fld = Tbl.CreateField("IdActFijo", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.IdActFijo", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdGrupo", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.IdGrupo", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("PrecioFactura", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.PrecioFactura", vbExclamation
         lUpdOK = False
      End If
    
      ERR.Clear
      Set Fld = Tbl.CreateField("DerechosIntern", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.DerechosIntern", vbExclamation
         lUpdOK = False
      End If
    
      ERR.Clear
      Set Fld = Tbl.CreateField("Transporte", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.Transporte", vbExclamation
         lUpdOK = False
      End If
    
      ERR.Clear
      Set Fld = Tbl.CreateField("ObrasAdapt", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.ObrasAdapt", vbExclamation
         lUpdOK = False
      End If
    
      ERR.Clear
      Set Fld = Tbl.CreateField("PrecioAdquis", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.PrecioAdquis", vbExclamation
         lUpdOK = False
      End If
    
      ERR.Clear
      Set Fld = Tbl.CreateField("IVARecuperable", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.IVARecuperable", vbExclamation
         lUpdOK = False
      End If
    
      ERR.Clear
      Set Fld = Tbl.CreateField("FormacionPers", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.FormacionPers", vbExclamation
         lUpdOK = False
      End If
    
      ERR.Clear
      Set Fld = Tbl.CreateField("ObrasReubic", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.ObrasReubic", vbExclamation
         lUpdOK = False
      End If
    
      ERR.Clear
      Set Fld = Tbl.CreateField("TotalGastos", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.TotalGastos", vbExclamation
         lUpdOK = False
      End If
    
      ERR.Clear
      Set Fld = Tbl.CreateField("FechaIncorporacion", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.FechaIncorporacion", vbExclamation
         lUpdOK = False
      End If
    
      ERR.Clear
      Set Fld = Tbl.CreateField("FechaDisponible", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoFicha.FechaDisponible", vbExclamation
         lUpdOK = False
      End If
    
      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON ActFijoFicha (IdFicha) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE INDEX IdxActFijo ON ActFijoFicha (IdActFijo)"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE INDEX IdxGrupo ON ActFijoFicha (IdGrupo)"
         Rc = ExecSQL(DbMain, Q1, False)
                 
                 
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla ActFijoFicha", vbExclamation
         lUpdOK = False
         
      End If
      
      Set Tbl = Nothing
  
      'Agregamos Tabla ActFijoCompsFicha
      Set Tbl = New TableDef
      Tbl.Name = "ActFijoCompsFicha"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdCompFicha", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.IdCompFicha", vbExclamation
         lUpdOK = False
      End If
            
      ERR.Clear
      Set Fld = Tbl.CreateField("IdActFijo", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.IdActFijo", vbExclamation
         lUpdOK = False
      End If
  
      ERR.Clear
      Set Fld = Tbl.CreateField("IdGrupo", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.IdGrupo", vbExclamation
         lUpdOK = False
      End If
  
      ERR.Clear
      Set Fld = Tbl.CreateField("IdComp", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.IdComp", vbExclamation
         lUpdOK = False
      End If
  
      ERR.Clear
      Set Fld = Tbl.CreateField("PjeDivComp", dbSingle)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.PjeDivComp", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("ValorCompra", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.ValorCompra", vbExclamation
         lUpdOK = False
      End If
  
      ERR.Clear
      Set Fld = Tbl.CreateField("ValorResidual", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.ValorResidual", vbExclamation
         lUpdOK = False
      End If
  
      ERR.Clear
      Set Fld = Tbl.CreateField("PjeAmortizacion", dbSingle)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.PjeAmortizacion", vbExclamation
         lUpdOK = False
      End If
  
      ERR.Clear
      Set Fld = Tbl.CreateField("VidaUtil", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.VidaUtil", vbExclamation
         lUpdOK = False
      End If
  
      ERR.Clear
      Set Fld = Tbl.CreateField("CostosAdicionales", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.CostosAdicionales", vbExclamation
         lUpdOK = False
      End If
  
      ERR.Clear
      Set Fld = Tbl.CreateField("TasaDesc", dbSingle)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.TasaDesc", vbExclamation
         lUpdOK = False
      End If
  
      ERR.Clear
      Set Fld = Tbl.CreateField("CostoDesmant", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.CostoDesmant", vbExclamation
         lUpdOK = False
      End If
 
      ERR.Clear
      Set Fld = Tbl.CreateField("ValActCostoDesmant", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.ValActCostoDesmant", vbExclamation
         lUpdOK = False
      End If
 
      ERR.Clear
      Set Fld = Tbl.CreateField("ValorBien", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.ValorBien", vbExclamation
         lUpdOK = False
      End If
 
      ERR.Clear
      Set Fld = Tbl.CreateField("ValorRazonable_31_12", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ActFijoCompsFicha.ValorRazonable_31_12", vbExclamation
         lUpdOK = False
      End If
 
      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON ActFijoCompsFicha (IdCompFicha) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE INDEX IdxActFijo ON ActFijoCompsFicha (IdActFijo)"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE INDEX IdxComp ON ActFijoCompsFicha (IdGrupo, IdComp)"
         Rc = ExecSQL(DbMain, Q1, False)
                                  
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla ActFijoCompsFicha", vbExclamation
         lUpdOK = False
         
      End If
      
      Set Tbl = Nothing
  
  
      If lUpdOK Then
         lDbVer = 319
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V318 = lUpdOK

End Function

Public Function CorrigeBase_V317() As Boolean   'entregada el 21 de enero 2014
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
  

   On Error Resume Next
   
   '--------------------- Versión 317 -----------------------------------
   If lDbVer = 317 And lUpdOK = True Then
   
      'Agregamos campo TipoAjusteComp en LogComprobantes
      Set Tbl = DbMain.TableDefs("LogComprobantes")

      ERR.Clear
      Set Fld = Tbl.CreateField("TipoAjusteComp", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh

      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LogComprobantes.TipoAjusteComp", vbExclamation
         lUpdOK = False

      End If
      
      Set Tbl = Nothing
  
      If lUpdOK Then
         lDbVer = 318
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V317 = lUpdOK

End Function

Public Function CorrigeBase_V316() As Boolean   'entregada 10 enero 2014
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
  

   On Error Resume Next
   
   '--------------------- Versión 316 -----------------------------------
   If lDbVer = 316 And lUpdOK = True Then
   
      'Agregamos campo CodIFRS a Cuentas
      Set Tbl = DbMain.TableDefs("Cuentas")

      ERR.Clear
      Set Fld = Tbl.CreateField("CodIFRS", dbText, 15)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh

      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.CodIFRS", vbExclamation
         lUpdOK = False

      End If
      
      Set Tbl = Nothing
  
      'linkeamos nuevamente la tabla PlanAvanzado para que tome el nuevo campo CodIFRS
      DbComun = gDbPath & "\" & BD_COMUN
      Call LinkMdbTable(DbMain, DbComun, "PlanAvanzado", , , , gComunConnStr, True)
      
      
      'ahora actualizamos el campo CodIFRS con lo que está en los planes predefinidos
      Q1 = "UPDATE Cuentas INNER JOIN PlanAvanzado ON Cuentas.Codigo = PlanAvanzado.Codigo"
      Q1 = Q1 & " SET Cuentas.CodIFRS = PlanAvanzado.CodIFRS "
      Call ExecSQL(DbMain, Q1)
  
      
      If lUpdOK Then
         lDbVer = 317
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V316 = lUpdOK

End Function

Public Function CorrigeBase_V315() As Boolean   'entregada 13 nov 2013
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
  

   On Error Resume Next
   
   '--------------------- Versión 315 -----------------------------------
   If lDbVer = 315 And lUpdOK = True Then
   
      Q1 = "UPDATE Comprobante SET TipoAjuste = " & TAJUSTE_AMBOS & " WHERE TipoAjuste IS NULL "
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVer = 316
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V315 = lUpdOK

End Function

Public Function CorrigeBase_V314() As Boolean   'agregada 29 oct 2013     entregada 8 nov 2013
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
  

   On Error Resume Next
   
   '--------------------- Versión 314 -----------------------------------
   If lDbVer = 314 And lUpdOK = True Then
   
      'Agregamos campo TipoAjuste a Comprobante
      Set Tbl = DbMain.TableDefs("Comprobante")

      ERR.Clear
      Set Fld = Tbl.CreateField("TipoAjuste", dbByte)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh

      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Comprobante.TipoAjuste", vbExclamation
         lUpdOK = False

      End If
      
      Set Tbl = Nothing
      
      'Agregamos campos DebeTrib y HaberTrib a tabla Cuentas
      Set Tbl = DbMain.TableDefs("Cuentas")

      ERR.Clear
      Set Fld = Tbl.CreateField("DebeTrib", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh

      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.DebeTrib", vbExclamation
         lUpdOK = False

      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("HaberTrib", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh

      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.HaberTrib", vbExclamation
         lUpdOK = False

      End If
      
      Set Tbl = Nothing
      
      If lUpdOK Then
         lDbVer = 315
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V314 = lUpdOK

End Function

Public Function CorrigeBase_V313() As Boolean   'entregada 14 sep 2013
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
  

   On Error Resume Next
   
   '--------------------- Versión 313 -----------------------------------
   If lDbVer = 313 And lUpdOK = True Then
   
      'Agregamos campo ValIVAIrrec a Documento
      Set Tbl = DbMain.TableDefs("Documento")

      ERR.Clear
      Set Fld = Tbl.CreateField("ValIVAIrrec", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh

      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.ValIVAIrrec", vbExclamation
         lUpdOK = False

      End If
      
      Set Tbl = Nothing
      
      
      If lUpdOK Then
         lDbVer = 314
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V313 = lUpdOK

End Function

Public Function CorrigeBase_V312() As Boolean   'entregada 10 sep 2013
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
  

   On Error Resume Next
   
   '--------------------- Versión 312 -----------------------------------
   If lDbVer = 312 And lUpdOK = True Then
   
      'Agregamos campo PropIVA a Documento
      Set Tbl = DbMain.TableDefs("Documento")

      ERR.Clear
      Set Fld = Tbl.CreateField("PropIVA", dbInteger)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh

      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.PropIVA", vbExclamation
         lUpdOK = False

      End If
      
      'Agregamos tabla PropIVA_TotMensual
      
      Set Tbl = New TableDef
      Tbl.Name = "PropIVA_TotMensual"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("Mes", dbInteger)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "PropIVA_TotMensual.Mes", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("TotalAfecto", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "PropIVA_TotMensual.TotalAfecto", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("TotalExento", dbDouble)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "PropIVA_TotMensual.TotalExento", vbExclamation
         lUpdOK = False
      End If
      
      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX IdxMes ON PropIVA_TotMensual (Mes)"
         Rc = ExecSQL(DbMain, Q1, False)
                  
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla PropIVA_TotMensual", vbExclamation
         lUpdOK = False
         
      End If
      
      Set Tbl = Nothing
      
      
      
      If lUpdOK Then
         lDbVer = 313
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V312 = lUpdOK

End Function

Public Function CorrigeBase_V311() As Boolean   '13 ago 13
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
  

   On Error Resume Next
   
   '--------------------- Versión 311 -----------------------------------
   If lDbVer = 311 And lUpdOK = True Then
   
      'Agregamos campo a comprobante para la fecha de importación
      Set Tbl = DbMain.TableDefs("Comprobante")

      ERR.Clear
      Set Fld = Tbl.CreateField("FechaImport", dbLong)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh

      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Comprobante.FechaImport", vbExclamation
         lUpdOK = False

      End If
      
      
      If lUpdOK Then
         lDbVer = 312
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V311 = lUpdOK

End Function

Public Function CorrigeBase_V310() As Boolean   '10 jul 13
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
  

   On Error Resume Next
   
   '--------------------- Versión 310 -----------------------------------
   If lDbVer = 310 And lUpdOK = True Then
   
      'Agregamos tabla de log de operaciones con comprobantes
      
      Set Tbl = New TableDef
      Tbl.Name = "LogComprobantes"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdLog", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LogComprobantes.IdLog", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdComp", dbLong)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LogComprobantes.IdComp", vbExclamation
         lUpdOK = False
      End If
           
      ERR.Clear
      Set Fld = Tbl.CreateField("IdUsuario", dbLong)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LogComprobantes.IdUsuario", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("Fecha", dbDouble)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LogComprobantes.Fecha", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdOper", dbInteger)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LogComprobantes.IdOper", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("Estado", dbInteger)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LogComprobantes.Estado", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("CorrComp", dbLong)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LogComprobantes.CorrComp", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("FechaComp", dbLong)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LogComprobantes.FechaComp", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("TipoComp", dbByte)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LogComprobantes.TipoComp", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("EstadoComp", dbByte)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "LogComprobantes.EstadoComp", vbExclamation
         lUpdOK = False
      End If
      
           
           
      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX IdLog ON LogComprobantes (IdLog)"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE  INDEX Fecha ON LogComprobantes (Fecha)"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla LogComprobantes", vbExclamation
         lUpdOK = False
         
      End If
      
      Set Tbl = Nothing
      
      
      If lUpdOK Then
         lDbVer = 311
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V310 = lUpdOK

End Function

Public Function CorrigeBase_V309() As Boolean   '12 dic 12
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 309 -----------------------------------
   If lDbVer = 309 And lUpdOK = True Then
   
      'Agregamos Cred 33 bis a ParamEmpresa
      
      If gEmpresa.Ano >= 2012 Then
         Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('CREDART33',0,'0.04')"
         Call ExecSQL(DbMain, Q1)
      End If
   
      If lUpdOK Then
         lDbVer = 310
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V309 = lUpdOK

End Function

Public Function CorrigeBase_V308() As Boolean   '12 dic 12
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 308 -----------------------------------
   If lDbVer = 308 And lUpdOK = True Then
   
      'Agregamos campo a comprobante para marcar C.M.
      Set Tbl = DbMain.TableDefs("Comprobante")

      ERR.Clear
      Set Fld = Tbl.CreateField("EsCCMM", dbBoolean)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh

      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Comprobante.EsCCMM", vbExclamation
         lUpdOK = False

      End If
   
      If lUpdOK Then
         lDbVer = 309
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V308 = lUpdOK

End Function

Public Function CorrigeBase_V307() As Boolean   'Version 4.0, 26 sept 2012
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 307 -----------------------------------
   If lDbVer = 307 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
      
      'agregamos campo Cred4PorcAnoInit
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Cred4PorcAnoInit", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.Cred4PorcAnoInit", vbExclamation
         lUpdOK = False
      End If
                 
      If lUpdOK Then
         lDbVer = 308
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V307 = lUpdOK

End Function



Public Function CorrigeBase_V306() As Boolean   ''Version 4.0, ???? agosto 2012
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 306 -----------------------------------
   If lDbVer = 306 And lUpdOK = True Then
   
      'agregamos los campos de IFRS en la tabla Cuentas

      ERR.Clear
      Set Tbl = DbMain.TableDefs("Cuentas")

      ERR.Clear
      Set Fld = Tbl.CreateField("CodIFRS_EstRes", dbText, 15)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh

      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.CodIFRS_EstRes", vbExclamation
         lUpdOK = False

      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("CodIFRS_EstFin", dbText, 15)
      Tbl.Fields.Append Fld

      If ERR = 0 Then
         Tbl.Fields.Refresh

      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.CodIFRS_EstRes", vbExclamation
         lUpdOK = False

      End If
      
      'Asignamos los códigos IFRS en el plan de cuentas del cliente, si tiene plan básico, intermedio o avanzado
      
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
      
         If vFld(Rs("Valor")) = "BÁSICO" Or vFld(Rs("Valor")) = "INTERMEDIO" Or vFld(Rs("Valor")) = "AVANZADO" Then
         
            'linkeamos la tabla PlanAvanzado porque el LinkMdbAdm se llama después del corrige base y puede no estar linkeada
            Call LinkMdbTable(DbMain, gDbPath & "\" & BD_COMUN, "PlanAvanzado", , , , gComunConnStr, True)
            
            Q1 = "UPDATE Cuentas INNER JOIN PlanAvanzado "
            Q1 = Q1 & " ON Cuentas.Codigo = PlanAvanzado.Codigo AND Cuentas.Descripcion = PlanAvanzado.Descripcion "
            Q1 = Q1 & " SET Cuentas.CodIFRS_EstFin = PlanAvanzado.CodIFRS_EstFin, "
            Q1 = Q1 & "     Cuentas.CodIFRS_EstRes = PlanAvanzado.CodIFRS_EstRes "
         
            Call ExecSQL(DbMain, Q1)
         End If
      End If
      
      Call CloseRs(Rs)
      
      If lUpdOK Then
         lDbVer = 307
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V306 = lUpdOK


End Function

Public Function CorrigeBase_V305() As Boolean   'Version 3.0, 16 abr 2012
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String


   On Error Resume Next

   '--------------------- Versión 305 -----------------------------------
   If lDbVer = 305 And lUpdOK = True Then
   
      'eliminamos los códigos F22 637 y 638 del plan de cuentas
      Q1 = "UPDATE Cuentas SET CodF22 =  0 WHERE CodF22 IN(637, 638)"
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVer = 306
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V305 = lUpdOK


End Function

Public Function CorrigeBase_V304() As Boolean   'Version 3.0, 16 nov. 2011
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 304 -----------------------------------
   If lDbVer = 304 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("Empresa")
      
      'agregamos campo Franq14ter
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Franq14ter", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.Franq14ter", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo Franq14quater
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Franq14quater", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.Franq14quater", vbExclamation
         lUpdOK = False
      End If
    
           
      If lUpdOK Then
         lDbVer = 305
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V304 = lUpdOK

End Function

Public Function CorrigeBase_V303() As Boolean   'Version 3.0, 26 Oct. 2011
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 303 -----------------------------------
   If lDbVer = 303 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
      
      'agregamos campo FImported a tabla MovActivoFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FImported", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.FImported", vbExclamation
         lUpdOK = False
      End If
      
    
      'agregamos campo ValorReajustado a tabla MovActivoFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("ValReajustadoNetoAnt", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.ValReajustadoNetoAnt", vbExclamation
         lUpdOK = False
      End If
            
      If lUpdOK Then
         lDbVer = 304
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V303 = lUpdOK

End Function


Public Function CorrigeBase_V302() As Boolean   'Version 3.0, 7 sept. 2011
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Valor As Double
   Dim IdDoc As Long
   Dim ValIVA As Double
  

   On Error Resume Next
   
   '--------------------- Versión 302 -----------------------------------
   If lDbVer = 302 And lUpdOK = True Then
   
      ERR.Clear
      
      'recálculo de Docs con IVA Irrecuperable por cambio en detalle de otros impuestos en libro de compras
      Call LinkMdbTable(DbMain, gDbPath & "\" & BD_COMUN, "TipoDocs", , , , gComunConnStr, True)

      Q1 = "SELECT MovDocumento.idDoc, IdTipoValLib, EsRebaja, Sum(Debe) as SumDebe, Sum(Haber) as SumHaber"
      Q1 = Q1 & " FROM (Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc)"
      Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc"
      Q1 = Q1 & " WHERE Documento.TipoLib = " & LIB_COMPRAS
      Q1 = Q1 & " AND Documento.Estado <> " & ED_ANULADO
      Q1 = Q1 & " AND IdTipoValLib IN (" & LIBCOMPRAS_IVACREDFISC & "," & LIBCOMPRAS_IVAIRREC & "," & LIBCOMPRAS_IVAACTFIJO & ")"
      Q1 = Q1 & " AND MovEdited <> 0"
      Q1 = Q1 & " GROUP BY MovDocumento.idDoc, IdTipoValLib, EsRebaja"
      Q1 = Q1 & " ORDER BY MovDocumento.idDoc"
    
      Set Rs = OpenRs(DbMain, Q1)
      IdDoc = 0
      Valor = 0
      
      Do While Not Rs.EOF
      
         If IdDoc <> vFld(Rs("IdDoc")) Then    'cambió el documento
            If IdDoc > 0 Then
               Q1 = "UPDATE Documento SET IVA = " & ValIVA & " WHERE IdDoc = " & IdDoc
               Call ExecSQL(DbMain, Q1)
               
               Valor = 0
               ValIVA = 0
            End If
            IdDoc = vFld(Rs("IdDoc"))
         End If
         
         If vFld(Rs("EsRebaja")) Then
            Valor = vFld(Rs("SumHaber")) - vFld(Rs("SumDebe"))
         Else
            Valor = vFld(Rs("SumDebe")) - vFld(Rs("SumHaber"))
         End If

         ValIVA = ValIVA + Valor
         
         Rs.MoveNext
         
      Loop
      
      If IdDoc <> 0 Then    'el último  documento
         Q1 = "UPDATE Documento SET IVA = " & Valor & " WHERE IdDoc = " & IdDoc
         Call ExecSQL(DbMain, Q1)
      End If
      
      Call CloseRs(Rs)
         
      
      If lUpdOK Then
         lDbVer = 303
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V302 = lUpdOK

End Function

Public Function CorrigeBase_V301() As Boolean   'Version 3.0, 31 agosto 2011
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 301 -----------------------------------
   If lDbVer = 301 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("Documento")
      
      'agregamos campo IdDocAsoc a tabla Documento
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdDocAsoc", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.IdDocAsoc", vbExclamation
         lUpdOK = False
      End If
      
    
      
      If lUpdOK Then
         lDbVer = 302
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V301 = lUpdOK

End Function

Public Function CorrigeBase_V300() As Boolean   'Version 3.0, 29 julio 2011
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
  

   On Error Resume Next
   
   '--------------------- Versión 300 -----------------------------------
   If lDbVer <= 300 And lUpdOK = True Then
   
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("MovComprobante")
      
      'agregamos campo Nota a tabla MovComprobante
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Nota", dbText, 120)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovComprobante.Nota", vbExclamation
         lUpdOK = False
      End If
      
      'agrandamos en campo orden para dar más espacio a Orden cuando se centralizan documentos
      Call AlterField(DbMain, "MovComprobante", "Orden", dbLong)
      
      'agrandamos el campo NumDoc para dar más espacio a para documentos de importación
      Call AlterField(DbMain, "Documento", "NumDoc", dbText, 20)
      Call AlterField(DbMain, "Documento", "NumDocHasta", dbText, 20)
      Call AlterField(DbMain, "Documento", "NumDocRef", dbText, 20)
      
      'agregamos capos a tabla Docemnto para manejas el Doc Maquina Registradora
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
      
      'campo NumFiscImpr (n° fiscal impresora)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NumFiscImpr", dbText, 20)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.NumFiscImpr", vbExclamation
         lUpdOK = False
      End If
      
      'campo NumInformeZ (n° informe Z)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NumInformeZ", dbText, 20)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.NumInformeZ", vbExclamation
         lUpdOK = False
      End If
      
      'campo CantBoletas (cantidad de boletas emitidas)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CantBoletas", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.CantBoletas", vbExclamation
         lUpdOK = False
      End If
      
      'campo VentasAcumInfZ (ventas acumuladas según informe Z)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("VentasAcumInfZ", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.VentasAcumInfZ", vbExclamation
         lUpdOK = False
      End If
      
      
      
      If lUpdOK Then
         lDbVer = 301
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V300 = lUpdOK

End Function

Public Function CorrigeBase_V299() As Boolean   'Version 3.0, 13 julio 2011
   Dim Q1 As String

   On Error Resume Next
   
   '--------------------- Versión 299 -----------------------------------
   If lDbVer <= 299 And lUpdOK = True Then
   
      Call CrearTblCtasRazFin

      If lUpdOK Then
         lDbVer = 300
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V299 = lUpdOK

End Function

Private Function CorrigeBase_V50() As Boolean     '7 abr 2011 v2.0.11
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer
   Dim StrCod As String

   On Error Resume Next
   
   '--------------------- Versión 50 -----------------------------------
   If lDbVer = 50 And lUpdOK = True Then
         
      ERR.Clear
      
      'eliminamos algunos códigos de exportación a Form 22, de acuerdo a lo solicitado por Victor Morales
      'en reporte 54-B, 7 abr 2011
      
      StrCod = "239, 240, 778, 779, 816, 817, 857, 858, 861"
      
      Q1 = "UPDATE Cuentas SET CodF22 = 0 WHERE CodF22 IN(" & StrCod & ")"
      Call ExecSQL(DbMain, Q1)
   
      
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 51
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V50 = lUpdOK

End Function

Private Function CorrigeBase_V49() As Boolean     '21 Enero 2011 v.2.0.9
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 49 -----------------------------------
   If lDbVer = 49 And lUpdOK = True Then
         
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
      
      'agregamos campo ValorLibro a tabla MovActivoFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("ValorLibro", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.ValorLibro", vbExclamation
         lUpdOK = False
      End If
      
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 50
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V49 = lUpdOK

End Function

Private Function CorrigeBase_V48() As Boolean     '13 dic 2010 v.2.0.9
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 48 -----------------------------------
   'No hace nada, sólo aumentar la versión
   If lDbVer = 48 And lUpdOK = True Then
         
      
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 49
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V48 = lUpdOK

End Function

Private Function CorrigeBase_V47() As Boolean     '7 ene 2010 (v 2.0.7)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 47 -----------------------------------
   If lDbVer = 47 And lUpdOK = True Then
         
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("MovComprobante")
      
      'agregamos campo DeRemu a tabla Cuentas
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DeRemu", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovComprobante.DeRemu", vbExclamation
         lUpdOK = False
      End If
      
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 48
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V47 = lUpdOK

End Function

Private Function CorrigeBase_V46() As Boolean     '30 nov.2009 (v 2.0.7)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 46 -----------------------------------
   If lDbVer = 46 And lUpdOK = True Then
         
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("Cuentas")
      
      'agregamos campo CorrelativoCheque a tabla Cuentas
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CorrelativoCheque", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.CorrelativoCheque", vbExclamation
         lUpdOK = False
      End If
      
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 47
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V46 = lUpdOK

End Function

Private Function CorrigeBase_V45() As Boolean     '23 oct.2009 (v 2.0.7)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 45 -----------------------------------
   If lDbVer = 45 And lUpdOK = True Then
         
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
      
      'agregamos campo TotalmenteDepreciado a tabla MovActFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TotalmenteDepreciado", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActFijo.TotalmenteDepreciado", vbExclamation
         lUpdOK = False
      End If
      
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 46
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V45 = lUpdOK

End Function

Private Function CorrigeBase_V44() As Boolean     '11 jun 2009 (v 2.0.5)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 44 -----------------------------------
   If lDbVer = 44 And lUpdOK = True Then
         
      ERR.Clear
      
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
      
      'agregamos campo ValReajustadoNeto a tabla MovActFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("ValReajustadoNeto", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActFijo.ValReajustadoNeto", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo IdActFijoOld a tabla MovActFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdActFijoOld", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActFijo.IdActFijoOld", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo IdActFijoOldTmp a tabla MovActFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdActFijoOldTmp", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActFijo.IdActFijoOldTmp", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo OldIdDocTmp a tabla Documento
      Set Tbl = DbMain.TableDefs("Documento")
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("OldIdDocTmp", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActFijo.OldIdDocTmp", vbExclamation
         lUpdOK = False
      End If
   
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 45
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V44 = lUpdOK

End Function

Private Function CorrigeBase_V43() As Boolean     '21 abr. 2009 (v 2.0.4)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 43 -----------------------------------
   If lDbVer = 43 And lUpdOK = True Then
         
      ERR.Clear
      
'      If gFunciones.ExpFUT Then   'se elimina esta verificación porque crea problemas con empresas nuevas hasta que saquemos actualización que ya no usa campo (FCA 29/11/2017)
         Call CrearTblCtasFUT
'      End If
      
      If gFunciones.DetSaldoApertura Then
         Call CreateTblDetSaldosAp
      End If
      
      'esto no se hace porque los que lo modificaron después de la versión 42, lo perderían.
      
'      'limpiamos tipo contrib de FUT y general
'      'ya no se usa el campo TFontribFUT porque se igauló al campo TipoContrib
'      Q1 = "UPDATE Empresa SET TipoContrib = 0, TContribFUT = 0 WHERE Id = " & gEmpresa.Id
'      Call ExecSQL(DbMain, Q1)
   
      MsgBox1 "Por favor verifique el tipo de contribuyente que tiene seleccionado para la empresa.", vbInformation
      
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = 44
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V43 = lUpdOK

End Function

Private Function CorrigeBase_V42() As Boolean     '25 nov. 2008 (v 2.0.4)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 42 -----------------------------------
   If lDbVer = 42 And lUpdOK = True Then
         
      ERR.Clear
      
      'limpiamos tipo contrib de FUT y general
      'ya no se usa el campo TFontribFUT porque se igauló al campo TipoContrib
      Q1 = "UPDATE Empresa SET TipoContrib = 0, TContribFUT = 0 WHERE Id = " & gEmpresa.id
      Call ExecSQL(DbMain, Q1, False)    'se cae! Por eso se agregó CorrigeBase_43 que crea el campo TContribFUT antes de actualizarlo
         
      ' MsgBox1 "Por favor verifique el tipo de contribuyente que tiene seleccionado para la empresa.", vbInformation
      
      
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = lDbVer + 1
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V42 = lUpdOK

End Function


Private Function CorrigeBase_V41() As Boolean     '9 oct. 2008 (v 2.0.4)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 41 -----------------------------------
   If lDbVer = 41 And lUpdOK = True Then
         
      ERR.Clear
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
      
      'agregamos campo NoDepreciable a tabla MovActFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NoDepreciable", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActFijo.NoDepreciable", vbExclamation
         lUpdOK = False
      End If
            
      'agregamos campo ValCred33 a tabla MovActFijo
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("ValCred33", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActFijo.ValCred33", vbExclamation
         lUpdOK = False
      End If
           
           
       'cambiamos largo campo Giro a tabla Empresa
      ERR.Clear
      
      Call AlterField(DbMain, "Empresa", "Giro", dbText, 80)
          
            
   '--------------------- Actualización Versión -------------------------
   
      If lUpdOK Then
         lDbVer = lDbVer + 1
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V41 = lUpdOK

End Function
Private Function CorrigeBase_V40() As Boolean     '21 jul 2008 (v 2.0.4)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 39 -----------------------------------
   If lDbVer = 40 And lUpdOK = True Then
         
      'cambiamos tipo campo IVAIrrecuperable a tabla Documento
      ERR.Clear
      
      Call AlterField(DbMain, "Documento", "IVAIrrecuperable", dbInteger)
      
   '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = lDbVer + 1
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V40 = lUpdOK

End Function
Private Function CorrigeBase_V39() As Boolean     '15 jul 2008 (v 2.0.4)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 39 -----------------------------------
   If lDbVer = 39 And lUpdOK = True Then
   
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
      
      'agregamos campo IVAIrrecuperable a tabla Documento
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IVAIrrecuperable", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.IVAIrrecuperable", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo DocOtrosEnAnalitico a tabla Documento
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DocOtrosEnAnalitico", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.DocOtrosEnAnalitico", vbExclamation
         lUpdOK = False
      End If
      
      
   '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = lDbVer + 1
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V39 = lUpdOK

End Function
Private Function CorrigeBase_V38() As Boolean     '3 jul 2008 (v 2.0.1)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 38 -----------------------------------
   If lDbVer = 38 And lUpdOK = True Then
   
      ERR.Clear
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FechaUtilizacion", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.FechaUtilizacion", vbExclamation
         lUpdOK = False
      End If
            
      
   '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 39
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V38 = lUpdOK

End Function
Private Function CorrigeBase_V37() As Boolean     '17 jun 2008 (v 2.0.1)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 37 -----------------------------------
   If lDbVer = 37 And lUpdOK = True Then
   
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Giro", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.Giro", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FacCompraRetParcial", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.FacCompraRetParcial", vbExclamation
         lUpdOK = False
      End If
      
      
   '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 38
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V37 = lUpdOK

End Function
Private Function CorrigeBase_V36() As Boolean     '11 mar 2008
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 36 -----------------------------------
   If lDbVer = 36 And lUpdOK = True Then
   
      Call CorrigeCodF22_2("Cuentas")
      
   '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 37
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V36 = lUpdOK

End Function
Private Function CorrigeBase_V35() As Boolean     '01 feb 2008
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 35 -----------------------------------
   If lDbVer = 35 And lUpdOK = True Then
   
      ERR.Clear
      Call CreateTableLockAction(DbMain, False)
      
   '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 36
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V35 = lUpdOK

End Function
Private Function CorrigeBase_V34() As Boolean     '27 Nov 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next
   
   '--------------------- Versión 34 -----------------------------------
   If lDbVer = 34 And lUpdOK = True Then
   
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Cartola")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("SaldoIni", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cartola.SaldoIni", vbExclamation
         lUpdOK = False
      End If
      
   '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 35
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V34 = lUpdOK

End Function
Private Function CorrigeBase_V33() As Boolean     '15 Junio 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
'   Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)
   
   '--------------------- Versión 33 -----------------------------------
   If lDbVer = 33 And lUpdOK = True Then    '15 Junio 2006
   
      Q1 = "SELECT IdDoc, Sum(MovComprobante.Debe) AS Debe, Sum(MovComprobante.Haber) AS Haber"
      Q1 = Q1 & " FROM MovComprobante GROUP BY IdDoc"
      Call DbMain.CreateQueryDef("vMovCompIdDoc", Q1)
      
      
   '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 34
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V33 = lUpdOK

End Function
Private Function CorrigeBase_V32() As Boolean     '24 Abril 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
'   Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)
   
   '--------------------- Versión 32 -----------------------------------
   If lDbVer = 32 And lUpdOK = True Then    '24 Abril 2006
   
      Call CorrigeTipoCapPropio_1("Cuentas")
      
   '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 33
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V32 = lUpdOK

End Function
Private Function CorrigeBase_V31() As Boolean     '5 Abril 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
'   Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)
   
   '--------------------- Versión 31 -----------------------------------
   If lDbVer = 31 And lUpdOK = True Then    '5 Abril 2006
   
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('VALORIVA',0,'0.19')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      
   '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 32
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V31 = lUpdOK

End Function
Private Function CorrigeBase_V30() As Boolean     '27 Marzo 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
'   Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)
   
   '--------------------- Versión 30 -----------------------------------
   If lDbVer = 30 And lUpdOK = True Then    '27 Mar 2006
   
      '--------------------- Documento -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FImportSuc", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.FImportSuc", vbExclamation
         lUpdOK = False
      End If
      
      Call CorrigeCodF22_1("Cuentas")
     
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 31
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V30 = lUpdOK

End Function
Private Function CorrigeBase_V29() As Boolean     '17 Marzo 2006 PS
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
'   Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)
   
   '--------------------- Versión 29 -----------------------------------
   If lDbVer = 29 And lUpdOK = True Then    '17 Mar 2006
   
     Call AlterField(DbMain, "Empresa", "CodActEconom", dbText, 8)
     
    '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 30
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V29 = lUpdOK

End Function

Private Function CorrigeBase_V28() As Boolean     '08 Marzo 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
'   Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)

   '--------------------- Versión 28 -----------------------------------
   If lDbVer = 28 And lUpdOK = True Then    '08 Mar 2006
         
      '--------------------- Tabla Documento -----------------
                  
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
                               
      'agregamos campo TotPagadoAnoAnt
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TotPagadoAnoAnt", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.TotPagadoAnoAnt", vbExclamation
         lUpdOK = False
      End If
      
           
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 29
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V28 = lUpdOK

End Function
Private Function CorrigeBase_V27() As Boolean     '23 Enero 2006
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
'   Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)

   '--------------------- Versión 27 -----------------------------------
   If lDbVer = 27 And lUpdOK = True Then    '23 Ene 2006
         
      '--------------------- Codigos F22 en Tabla Cuentas -----------------
                  
      ERR.Clear
      'eliminamos algunos códigos de exportación a Form 22:
      '  628 porque es un campo que se ingresa con detalles, no ingreso directo
      '  366 y 384 porque son campos calculados por Form 22
      Q1 = "UPDATE Cuentas SET CodF22 = 0 WHERE CodF22 IN( 628, 366, 384)"
      Call ExecSQL(DbMain, Q1)
      
           
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 28
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V27 = lUpdOK

End Function
Private Function CorrigeBase_V26() As Boolean     '23 Diciembre 2005
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   'Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)

   '--------------------- Versión 26 -----------------------------------
   If lDbVer = 26 And lUpdOK = True Then    '23 Dic 2005
         
      '--------------------- Notas -----------------
                  
      ERR.Clear
      Call AlterField(DbMain, "Notas", "Incluir", dbInteger)
           
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 27
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V26 = lUpdOK

End Function
Private Function CorrigeBase_V25() As Boolean     '14 Diciembre 2005
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   'Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)

   '--------------------- Versión 25 -----------------------------------
   If lDbVer = 25 And lUpdOK = True Then    '14 Dic 2005
         
      '--------------------- Comprobante Tipo -----------------
                  
      ERR.Clear
      Set Tbl = DbMain.TableDefs("CT_Comprobante")
                               
      'agregamos campo Imprimir Resumido (para versión 2006)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("ImpResumido", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "ImpResumido", vbExclamation
         lUpdOK = False
      End If
           
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 26
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V25 = lUpdOK

End Function
Private Function CorrigeBase_V24() As Boolean     '15 Octubre 2005
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   'Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)

   '--------------------- Versión 24 -----------------------------------
   If lDbVer = 24 And lUpdOK = True Then    '15 Oct. 2005
         
      '--------------------- Comprobante -----------------
                  
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Comprobante")
                               
      'agregamos campo Imprimir Resumido (para versión 2006)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("ImpResumido", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Comprobante.ImpResumido", vbExclamation
         lUpdOK = False
      End If
           
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 25
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V24 = lUpdOK

End Function

Private Function CorrigeBase_V23() As Boolean     '25 Agosto 2005
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   'Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)

   '--------------------- Versión 23 -----------------------------------
   If lDbVer = 23 And lUpdOK = True Then    '25 Agosto 2005
         
      '--------------------- MovActivoFijo -----------------
                  
      ERR.Clear
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
             
                  
      'cambiamos nobres de campos de depreciación en MovActivoFijo
      ERR.Clear
      Tbl.Fields("DepFin").Name = "DepNormal"
      Tbl.Fields("DepTrib").Name = "DepAcelerada"
      Tbl.Fields("DepFinHist").Name = "DepNormalHist"
      Tbl.Fields("DepTribHist").Name = "DepAceleradaHist"
      
      'agregamos campo Tipo depreciación
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoDep", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.TipoDep", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo Tipo Depreciación Histórica
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoDepHist", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.TipoDepHist", vbExclamation
         lUpdOK = False
      End If

      'agregamos campo Depreciación Acumulada Histórica
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DepAcumHist", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepAcumHist", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo Vida útil
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("VidaUtil", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.VidaUtil", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo Depreciación Acumulada Final (a final del año actual)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DepAcumFinal", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepAcumFinal", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo Vida útil Residual (a final año actual)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("VidaUtilResidual", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.VidaUtilResidual", vbExclamation
         lUpdOK = False
      End If
      
      'agregamos campo FExported (Fecha exportación a año siguiente o importación desde año anterior)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FExported", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.FExported", vbExclamation
         lUpdOK = False
      End If
      
      'eliminamos campos que no se usan en MovActivoFijo
      ERR.Clear
      Tbl.Indexes.Delete ("IdDocVenta")
      Tbl.Fields.Delete ("IdDocVenta")
      ERR.Clear
      Tbl.Indexes.Delete ("IdCompVenta")
      Tbl.Fields.Delete ("IdCompVenta")
      ERR.Clear
      Tbl.Indexes.Delete ("IdMovCompVenta")
      Tbl.Fields.Delete ("IdMovCompVenta")
      ERR.Clear
      Tbl.Indexes.Delete ("IdCuentaVenta")
      Tbl.Fields.Delete ("IdCuentaVenta")
      
      'eliminamos campo IdActFijo de MovComprobante (no se usa)
      ERR.Clear
      Set Tbl = DbMain.TableDefs("MovComprobante")
      
      ERR.Clear
      Tbl.Fields.Delete ("IdActFijo")
           
           
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 24
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V23 = lUpdOK

End Function

Private Function CorrigeBase_V22() As Boolean      ' 12 Julio 2005
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
'   Dim DbVer As Integer, UpdOK As Boolean
   Dim i As Integer
   Dim Rc As Integer

   On Error Resume Next

'   UpdOK = True
'   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
'   Set Rs = OpenRs(DbMain, Q1)
'   If Rs.EOF Then
'      Call CloseRs(Rs)
'      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
'      Call ExecSQL(DbMain, Q1)
'      DbVer = 0
'   Else
'      DbVer = Val(vFld(Rs("Valor")))
'   End If
'
'   Call CloseRs(Rs)

   '--------------------- Versión 22 -----------------------------------
   If lDbVer = 22 And lUpdOK = True Then   ' 12 Julio 2005
         
      '--------------------- Creamos tabla Sucursales -----------------
        
      Set Tbl = New TableDef
      Tbl.Name = "Sucursales"
      
      ERR.Clear
      Set Fld = Tbl.CreateField("IdSucursal", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Sucursales.IdSucursal", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Set Fld = Tbl.CreateField("Codigo", dbText, 15)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Sucursales.Codigo", vbExclamation
         lUpdOK = False
      End If
           
      ERR.Clear
      Set Fld = Tbl.CreateField("Descripcion", dbText, 30)
      Tbl.Fields.Append Fld
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Sucursales.Descripcion", vbExclamation
         lUpdOK = False
      End If
           
      DbMain.TableDefs.Append Tbl
      If ERR = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX IdSucursal ON Sucursales (IdSucursal)"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla Sucursales", vbExclamation
         lUpdOK = False
         
      End If
      
      Set Tbl = Nothing
      
      '--------------------- Add Campo Sucursal a Docs ---------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdSucursal", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.IdSucursal", vbExclamation
         lUpdOK = False
      End If
           
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 23
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V22 = lUpdOK
   
End Function

Private Function CorrigeBase_2005_01() As Boolean
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rs As Recordset
   Dim id As Integer
   Dim i As Integer, Rc As Long
   Dim MaxTipoDoc As Integer
'   Dim DbVer As Integer, UpdOK As Boolean

   On Error Resume Next

   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      lDbVer = 0
   Else
      lDbVer = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)
   
'   If lDbVer = 709 Then    'parche por cambio de IdEmpresa y Año
'      lDbVer = 365
'   End If

   '--------------------- Versión 0 -----------------------------------

   If lDbVer = 0 Then

      '--------------------- LogImpreso -----------------------------------

      Set Tbl = DbMain.TableDefs("LogImpreso")

      ERR.Clear
      Tbl.Fields.Delete "id"
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3265 Then
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LogImpreso.Id", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Set Fld = Tbl.CreateField("IdLog", dbLong)
      Fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append Fld
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LogImpreso.IdLog", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdUsuario", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LogImpreso.IdUsuario", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Mes", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LogImpreso.Mes", vbExclamation
         lUpdOK = False
      End If
      
      '--------------------- Notas -----------------------------------

      Set Tbl = DbMain.TableDefs("Notas")

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IncluirInfo", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Notas.IncluirInfo", vbExclamation
         lUpdOK = False
      End If

      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 1
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If

   '--------------------- Versión 1 -----------------------------------
   
   If lDbVer = 1 And lUpdOK = True Then

      '--------------------- Empresa -----------------------------------

      ERR.Clear
      
      Call AlterField(DbMain, "Empresa", "RutRepLegal1", dbText, 12)
      Call AlterField(DbMain, "Empresa", "RutRepLegal2", dbText, 12)

      
      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 2
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 2 -----------------------------------
   
   If lDbVer = 2 And lUpdOK = True Then

      '--------------------- Documento -----------------------------------

      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FEmisionOri", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "FEmisionOri.FEmisionOri", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CorrInterno", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "FEmisionOri.CorrInterno", vbExclamation
         lUpdOK = False
      End If
      
      
      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 3
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 3 -----------------------------------
   
   If lDbVer = 3 And lUpdOK = True Then

      '--------------------- Documento -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("SaldoDoc", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.SaldoDoc", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Call ExecSQL(DbMain, "UPDATE Documento SET SaldoDoc=NULL")
      
      
      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 4
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 4 -----------------------------------
   
   If lDbVer = 4 And lUpdOK = True Then

      '--------------------- LogImpreso -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("LogImpreso")

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FDesde", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LogImpreso.FDesde", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FHasta", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "LogImpreso.FHasta", vbExclamation
         lUpdOK = False
      End If
      
      '--------------------- Documento -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FExported", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.FExported", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("OldIdDoc", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.OldIdDoc", vbExclamation
         lUpdOK = False
      End If
      
      If lUpdOK Then
         lDbVer = 5
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
            
   End If
   
   '--------------------- Versión 5 -----------------------------------
   
   If lDbVer = 5 And lUpdOK = True Then   '13 Dic 2004
   
      '--------------------- MovActivoFijo -----------------------------------
      
      ERR.Clear
      Call AlterField(DbMain, "MovActivoFijo", "Cred4Porc", dbBoolean)
      
      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 6
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   '--------------------- Versión 6 -----------------------------------
   
   If lDbVer = 6 And lUpdOK = True Then   '16 Dic 2004
   
      '--------------------- Entidades -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Entidades")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Giro", dbText, 50)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.Giro", vbExclamation
         lUpdOK = False
      End If
      
      '--------------------- MovActivoFijo -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DepFinHist", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepFinHist", vbExclamation
         lUpdOK = False
      End If
            
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DepTribHist", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.DepTribHist", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NetoVenta", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.NetoVenta", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IVAVenta", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.IVAVenta", vbExclamation
         lUpdOK = False
      End If
      
      
      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 7
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If

   '--------------------- Versión 7 -----------------------------------
   
   If lDbVer = 7 And lUpdOK = True Then   '17 Dic 2004
         
      '--------------------- MovActivoFijo -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("MovActivoFijo")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FechaVentaBaja", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.FechaVentaBaja", vbExclamation
         lUpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdDocVenta", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovActivoFijo.IdDocVenta", vbExclamation
         lUpdOK = False
      End If
      
      '--------------------- MovComprobante -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("MovComprobante")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdActFijo", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "MovComprobante.IdActFijo", vbExclamation
         lUpdOK = False
      End If
      
      
      '--------------------- Documento -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
      
      'agergamos campo es DTE (Doc Trib Electrónico)
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DTE", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.DTE", vbExclamation
         lUpdOK = False
      End If
      
      'porcentaje de retención: 10% o 20%  (nacional o extranjero)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("PorcentRetencion", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.PorcentRetencion", vbExclamation
         lUpdOK = False
      End If
      
      'Tipo de retención: Dieta u Honorarios
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoRetencion", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.TipoRetencion", vbExclamation
         lUpdOK = False
      End If
      
           
      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 8
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   
   '--------------------- Versión 8 -----------------------------------
   
   If lDbVer = 8 And lUpdOK = True Then   '27 Dic 2004
         
     
      '--------------------- Cuentas -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Cuentas")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Cuentas.CodF29", vbExclamation
         lUpdOK = False
      End If

      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 9
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   '--------------------- Versión 9 -----------------------------------
   
   If lDbVer = 9 And lUpdOK = True Then   '4 Ene 2005
         
     
      '--------------------- Empresa -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Empresa")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoContrib", dbInteger)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.TipoContrib", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TransaBolsa", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.TransaBolsa", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Franq14bis", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.Franq14bis", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqLey18392", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqLey18392", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqDL600", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqDL600", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqDL701", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqDL701", vbExclamation
         lUpdOK = False
      End If

      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FranqDS341", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.FranqDS341", vbExclamation
         lUpdOK = False
      End If


      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 10
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   
   '--------------------- Versión 10 -----------------------------------
   
   If lDbVer = 10 And lUpdOK = True Then   '13 Ene 2005
         
         
      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 11
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   '--------------------- Versión 11 -----------------------------------
   
   If lDbVer = 11 And lUpdOK = True Then   '17 Ene 2005
         
      Q1 = "CREATE UNIQUE INDEX TipoCod ON ParamEmpresa (Tipo, Codigo ) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "DROP TABLE TipoDocs"
      Rc = ExecSQL(DbMain, Q1, False)
      
      DbMain.TableDefs.Refresh
               
      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 12
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
     
   '--------------------- Versión 12 -----------------------------------
   
   If lDbVer = 12 And lUpdOK = True Then   '31 Ene 2005
         
      Q1 = "DROP INDEX NumDoc ON Documento "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX NumDoc ON Documento (TipoLib, TipoDoc, IdEntidad, NumDoc ) "
      Rc = ExecSQL(DbMain, Q1, False)
                     
      '--------------------- Documento -----------------------------------
      
      'estos campos ya habían sido agregados en la versión 7 pero algunos usuarios por alguna razón no los tienen ????
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
      
      'agregamos campo es DTE (Doc Trib Electrónico)
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DTE", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.DTE", vbExclamation
         lUpdOK = False
      End If
      
      'porcentaje de retención: 10% o 20%  (nacional o extranjero)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("PorcentRetencion", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.PorcentRetencion", vbExclamation
         lUpdOK = False
      End If
      
      'Tipo de retención: Dieta u Honorarios
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoRetencion", dbByte)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.TipoRetencion", vbExclamation
         lUpdOK = False
      End If
      
      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 13
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   '--------------------- Versión 13 -----------------------------------
   
   If lDbVer = 13 And lUpdOK = True Then   '7 Mar 2005
                              
      '--------------------- Documento -----------------------------------
            
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
            
      'agregamos campo MovEdited que indica si los movimientos han sido editados en la ventana de detalle del doc.
      'para evitar que el libro de compras, ventas o retenciones pise los movimientos definidos por el usuario
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("MovEdited", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.MovEdited", vbExclamation
         lUpdOK = False
      End If
     
      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 14
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If

   '--------------------- Versión 14 -----------------------------------
   
   If lDbVer = 14 And lUpdOK = True Then   '17 Mar 2005
                              
      '--------------------- Documento -----------------------------------
            
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
            
      'agregamos campo OtrosVal para acumular total de otros valores que no sean Afecto, Exento, IVA, OtrosImp
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("OtrosVal", dbDouble)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.OtrosVal", vbExclamation
         lUpdOK = False
      End If
     
      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 15
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If


   '--------------------- Versión 15 -----------------------------------
   
   If lDbVer = 15 And lUpdOK = True Then   '28 Mar 2005
                              
      '--------------------- Documento -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
      
      'Fecha de importación desde F29
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FImporF29", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.FImporF29", vbExclamation
         lUpdOK = False
      End If
            
      'número del documento asociado (por ejemplo una factura asociada a una nota de crédito)
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NumDocRef", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.NumDocRef", vbExclamation
         lUpdOK = False
      End If

      ' Id de la cuenta contable correspondiente al cheque con el que se paga este documento
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdCtaBanco", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.IdCtaBanco", vbExclamation
         lUpdOK = False
      End If

      '---------------------  Entidades -----------------------------------
      
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Entidades")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NotValidRut", dbBoolean)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Entidades.NotValidRut", vbExclamation
         lUpdOK = False
      End If


      '--------------------- Actualización Versión -------------------------

      If lUpdOK Then
         lDbVer = 16
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   '--------------------- Versión 16 -----------------------------------
   
   If lDbVer = 16 And lUpdOK = True Then   '13 Abril 2005
   
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Empresa")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Opciones", dbLong)

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.Opciones", vbExclamation
         lUpdOK = False
      End If
      
      'Funcion creada por PAM
      Call CreateTableLockAction(DbMain)
      
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 17
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If

   '--------------------- Versión 17 -----------------------------------
   
   If lDbVer = 17 And lUpdOK = True Then   '28 Abril 2005
   
      Q1 = "UPDATE Cuentas SET Atrib" & ATRIB_CAJA & "=1 WHERE " & GenLike(DbMain, "Caja", "Descripcion", GL_RWILD)
      Call ExecSQL(DbMain, Q1)
      
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 18
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If

   '--------------------- Versión 18 -----------------------------------
   
   If lDbVer = 18 And lUpdOK = True Then   ' 13 may 2005
         
      'actualizamos el idPadre de una cuenta del plan intermedio
      Q1 = "SELECT Codigo, IdCuenta FROM Cuentas WHERE Codigo IN ( '3050100', '3050101' ) ORDER BY Codigo"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
      
         If vFld(Rs("Codigo")) = "3050100" Then ' existe el padre
            Q1 = "UPDATE Cuentas SET IdPadre = " & vFld(Rs("IdCuenta")) & " WHERE Codigo = '3050101'"
            Call CloseRs(Rs)
           
         Else ' no existe el padre
            Q1 = "DELETE * FROM Cuentas WHERE Codigo = '3050101'"
            Call CloseRs(Rs)
         End If
         
         Call ExecSQL(DbMain, Q1) ' se actualiza después de cerrar el Recordset porque podría estar bloqueado
      End If
      Call CloseRs(Rs)
      
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 19
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If

   '--------------------- Versión 19 -----------------------------------
   
   If lDbVer = 19 And lUpdOK = True Then   ' 7 Junio 2005
         
      'corregimos Comprobantes Tipo para planes Intermedio y Basico
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'PLANCTAS'"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         If vFld(Rs("Valor")) = "BÁSICO" Or vFld(Rs("Valor")) = "INTERMEDIO" Then
            Call UpdateComprobantesTipo
         End If
      End If
      
      Call CloseRs(Rs)
      
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 20
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If

   '--------------------- Versión 20 -----------------------------------
   
   If lDbVer = 20 And lUpdOK = True Then   ' 14 Junio 2005, después de generar Test del día
         
      'agregamos campo TipoRelEnt a tabla documentos para almacenar si la entidad relacionada es: Emisor, Receptor u Otro
      ERR.Clear
      Set Tbl = DbMain.TableDefs("Documento")
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoRelEnt", dbInteger)  'tipo de relación de la entidad con el doc

      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error & vbLf & "Documento.TipoRelEnt", vbExclamation
         lUpdOK = False
      End If
      
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 21
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   '--------------------- Versión 21 -----------------------------------
   
   If lDbVer = 21 And lUpdOK = True Then   ' 20 Junio 2005
         
      Call AlterField(DbMain, "MovComprobante", "Orden", dbInteger)
      
      '--------------------- Actualización Versión -------------------------
      If lUpdOK Then
         lDbVer = 22
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_2005_01 = lUpdOK
   
End Function
'2699584 Crea tabla Percepciones
Public Function CreateTblPercepciones() As Boolean
   Dim Q1 As String
   Dim Rc As Long
   Dim Tbl As TableDef, Fld As Field
   
   On Error Resume Next
   
   'Creamos tabla Percepciones
   
   CreateTblPercepciones = True
   
   Set Tbl = New TableDef
   Tbl.Name = "Percepciones"
   
   ERR.Clear
   Set Fld = Tbl.CreateField("IDPerc", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.IDPerc", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("IDComp", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.IDComp", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("Orden", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.Orden", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("IdCuenta", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.IdCuenta", vbExclamation
   End If
   
    ERR.Clear
    Set Fld = Tbl.CreateField("IdEmpresa", dbLong)
    Tbl.Fields.Append Fld

    If ERR = 0 Then
       Tbl.Fields.Refresh
    ElseIf ERR <> 3191 Then ' ya existe
       MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.IdEmpresa", vbExclamation
       lUpdOK = False
    End If
    
    ERR.Clear
    Set Fld = Tbl.CreateField("Ano", dbInteger)
    Tbl.Fields.Append Fld

    If ERR = 0 Then
       Tbl.Fields.Refresh
    ElseIf ERR <> 3191 Then ' ya existe
       MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.Ano", vbExclamation
       lUpdOK = False
    End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("Fecha", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.Fecha", vbExclamation
   End If
      
   ERR.Clear
   Set Fld = Tbl.CreateField("NumCertificado", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.NumCertificado", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("RutEmpresa", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.RutEmpresa", vbExclamation
   End If
      
   ERR.Clear
   Set Fld = Tbl.CreateField("Regimen", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.Regimen", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("Contabilizacion", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.Contabilizacion", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("TasaTef", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.TasaTef", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("TasaTex", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.TasaTex", vbExclamation
   End If
      
   ERR.Clear
   Set Fld = Tbl.CreateField("Percepciones", dbDouble)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.Percepciones", vbExclamation
   End If
      
      
   DbMain.TableDefs.Append Tbl
   Set Tbl = Nothing

   DbMain.TableDefs.Refresh

'   Q1 = "CREATE UNIQUE INDEX Id ON DetSaldosAp (Id ) WITH PRIMARY"
'   Rc = ExecSQL(DbMain, Q1, False)
'
'   Q1 = "CREATE UNIQUE INDEX IdEntidad ON DetSaldosAp (IdCuenta, IdEntidad)"
'   Rc = ExecSQL(DbMain, Q1, False)

End Function
Public Function CreateTblDetPercepciones() As Boolean
   Dim Q1 As String
   Dim Rc As Long
   Dim Tbl As TableDef, Fld As Field
   
   On Error Resume Next
   
   'Creamos tabla Percepciones
   
   CreateTblDetPercepciones = True
   
   Set Tbl = New TableDef
   Tbl.Name = "DetPercepciones"
   
   ERR.Clear
   Set Fld = Tbl.CreateField("IDPerc", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblDetPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetPercepciones.IDPerc", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("CodDet", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblDetPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetPercepciones.CodDet", vbExclamation
   End If
      
   ERR.Clear
   Set Fld = Tbl.CreateField("Valor", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblDetPercepciones = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetPercepciones.Valor", vbExclamation
   End If
   
'   ERR.Clear
'   Set Fld = Tbl.CreateField("RutEmpresa", dbLong)
'   Tbl.Fields.Append Fld
'
'   If ERR = 0 Then
'      Tbl.Fields.Refresh
'   ElseIf ERR <> 3191 Then ' ya existe
'      CreateTblDetPercepciones = False
'      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Percepciones.RutEmpresa", vbExclamation
'   End If
      

   DbMain.TableDefs.Append Tbl
   Set Tbl = Nothing

   DbMain.TableDefs.Refresh

'   Q1 = "CREATE UNIQUE INDEX Id ON DetSaldosAp (Id ) WITH PRIMARY"
'   Rc = ExecSQL(DbMain, Q1, False)
'
'   Q1 = "CREATE UNIQUE INDEX IdEntidad ON DetSaldosAp (IdCuenta, IdEntidad)"
'   Rc = ExecSQL(DbMain, Q1, False)

End Function
' Fin 2699584
'Crea tabla Detalle de Saldos de Apertura
Private Function CreateTblDetSaldosAp() As Boolean    '04 Oct. 2005
   Dim Q1 As String
   Dim Rc As Long
   Dim Tbl As TableDef, Fld As Field
   
   On Error Resume Next
   
   'Creamos tabla DetSaldosAp
   
   CreateTblDetSaldosAp = True
   
   Set Tbl = New TableDef
   Tbl.Name = "DetSaldosAp"
   
   ERR.Clear
   Set Fld = Tbl.CreateField("Id", dbLong)
   Fld.Attributes = dbAutoIncrField
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblDetSaldosAp = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetSaldosAp.Id", vbExclamation
   End If
      
   ERR.Clear
   Set Fld = Tbl.CreateField("IdCuenta", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblDetSaldosAp = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetSaldosAp.IdCuenta", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("IdEntidad", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblDetSaldosAp = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetSaldosAp.IdEntidad", vbExclamation
   End If
      
   ERR.Clear
   Set Fld = Tbl.CreateField("Debe", dbDouble)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblDetSaldosAp = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetSaldosAp.Debe", vbExclamation
   End If
      
   ERR.Clear
   Set Fld = Tbl.CreateField("Haber", dbDouble)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblDetSaldosAp = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetSaldosAp.Haber", vbExclamation
   End If
      
   ERR.Clear
   Set Fld = Tbl.CreateField("Saldo", dbDouble)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblDetSaldosAp = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "DetSaldosAp.Saldo", vbExclamation
   End If
      
   DbMain.TableDefs.Append Tbl
   Set Tbl = Nothing

   DbMain.TableDefs.Refresh

   Q1 = "CREATE UNIQUE INDEX Id ON DetSaldosAp (Id ) WITH PRIMARY"
   Rc = ExecSQL(DbMain, Q1, False)
   
   Q1 = "CREATE UNIQUE INDEX IdEntidad ON DetSaldosAp (IdCuenta, IdEntidad)"
   Rc = ExecSQL(DbMain, Q1, False)

End Function

Public Function CrearTblCtasFUT() As Boolean
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Q1 As String
   Dim Rc As Integer

   On Error Resume Next

   CrearTblCtasFUT = True

   '--------------------- Agregamos campo TContribFUT a tabla Empresa -----------------

   Set Tbl = DbMain.TableDefs("Empresa")

   ERR.Clear
   Set Fld = Tbl.CreateField("TContribFUT", dbLong)
   Tbl.Fields.Append Fld
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then
      MsgBeep vbExclamation
      MsgBox "Error " & ERR & ", " & Error & vbLf & "Empresa.TContribFUT", vbExclamation
      CrearTblCtasFUT = False
   End If
      
   '--------------------- Crear tabla CuentasFUT -----------------
   
   Set Tbl = New TableDef
   Tbl.Name = "CuentasFUT"
   
   ERR.Clear
   Set Fld = Tbl.CreateField("Id", dbLong)
   Fld.Attributes = dbAutoIncrField
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CuentasFUT.Id", vbExclamation
      CrearTblCtasFUT = False
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("TipoIngGas", dbInteger)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CuentasFUT.TipoIngGas", vbExclamation
      CrearTblCtasFUT = False
   End If
      
   ERR.Clear
   Set Fld = Tbl.CreateField("IdItem", dbInteger)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CuentasFUT.IdItem", vbExclamation
      CrearTblCtasFUT = False
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("IdCuenta", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CuentasFUT.IdCuenta", vbExclamation
      CrearTblCtasFUT = False
   End If

   ERR.Clear
   Set Fld = Tbl.CreateField("CodCuenta", dbText, 15)  'se almacena el código para agilizar el query sobre los movimientos, dado que pueden ser cuentas no de último nivel
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CuentasFUT.CodCuenta", vbExclamation
      CrearTblCtasFUT = False
   End If
   
   DbMain.TableDefs.Append Tbl
   If ERR = 0 Then
      DbMain.TableDefs.Refresh
      
      Q1 = "CREATE UNIQUE INDEX Id ON CuentasFUT (Id) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE UNIQUE INDEX IdItem ON CuentasFUT (TipoIngGas, IdItem)"
      Rc = ExecSQL(DbMain, Q1)
      
   ElseIf ERR <> 3010 Then ' ya existe
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla CuentasFUT", vbExclamation
      CrearTblCtasFUT = False
      
   End If
   
   Set Tbl = Nothing
   
End Function

'Esta función debe ser invocada desde CorrigeBaseAdm y
'Se debe agregar en link de la tabla en las db empresa-año
Private Function CrearTblCtasRazFin() As Boolean
   Dim Q1 As String
   Dim Rc As Long
   Dim Tbl As TableDef, Fld As Field
   
   
   On Error Resume Next
   
   'Creamos tabla CuentasRazon

   CrearTblCtasRazFin = True
   
   Set Tbl = New TableDef
   Tbl.Name = "CuentasRazon"
   
   ERR.Clear
   Set Fld = Tbl.CreateField("IdRazon", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CrearTblCtasRazFin = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CuentasRazon.IdRazon", vbExclamation
   End If
      
   ERR.Clear
   Set Fld = Tbl.CreateField("NumDenom", dbInteger)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CrearTblCtasRazFin = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CuentasRazon.NumDenom", vbExclamation
   End If
      
   ERR.Clear
   Set Fld = Tbl.CreateField("CodCuenta", dbText, 15)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CrearTblCtasRazFin = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CuentasRazon.IdCuenta", vbExclamation
   End If
            
   ERR.Clear
   Set Fld = Tbl.CreateField("Operador", dbText, 1)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CrearTblCtasRazFin = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "CuentasRazon.Operador", vbExclamation
   End If
   
   DbMain.TableDefs.Append Tbl
   Set Tbl = Nothing

   DbMain.TableDefs.Refresh

   Q1 = "CREATE UNIQUE INDEX IdCtaRazon ON CuentasRazon (IdRazon, NumDenom, IdCuenta) WITH PRIMARY"
   Rc = ExecSQL(DbMain, Q1, False)
   
   
   'tabla de parámetros razones financieras
   Set Tbl = New TableDef
   Tbl.Name = "ParamRazon"
   
   ERR.Clear
   Set Fld = Tbl.CreateField("IdRazon", dbLong)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CrearTblCtasRazFin = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "ParamRazon.IdRazon", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("CantDias", dbLong)
   Tbl.Fields.Append Fld

   If ERR = 0 Then
      Tbl.Fields.Refresh
   
   ElseIf ERR <> 3191 Then ' ya existe
      MsgBeep vbExclamation
      MsgBox "Error " & ERR & ", " & Error & vbLf & "ParamRazon.CantDias", vbExclamation
      CrearTblCtasRazFin = False
   
   End If

   DbMain.TableDefs.Append Tbl
   Set Tbl = Nothing
   
   DbMain.TableDefs.Refresh

   Q1 = "CREATE UNIQUE INDEX IdxParamRazon ON ParamRazon (IdRazon) WITH PRIMARY"
   Rc = ExecSQL(DbMain, Q1, False)


End Function

Private Sub UpdImpAdicionales2016()
   Dim Q1 As String

   'Compras
   
   If gEmpresa.Ano < 2016 Then
   
      'Cambiamos todos los otros impuestos asociados a Diesel y Transporte a Específico Diesel y Específico Transporte
      Q1 = "UPDATE MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc"
      Q1 = Q1 & " SET IdTipoValLib = " & LIBCOMPRAS_IMPESPDIESEL
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND IdTipoValLib IN(" & LIBCOMPRAS_IMPESPPETRGRAL & "," & LIBCOMPRAS_IMPESPPETRGENCF & "," & LIBCOMPRAS_IMPESPPETRGENSINCF & ")"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc"
      Q1 = Q1 & " SET IdTipoValLib = " & LIBCOMPRAS_IMPESPDIESELTRANS
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND IdTipoValLib IN(" & LIBCOMPRAS_IMPESPPETRCARGACF & "," & LIBCOMPRAS_IMPESPPETRCARGASINCF & ")"
      Call ExecSQL(DbMain, Q1)
         
      'Cambiamos ILA por Bebidas Analcoholicas con elevado cont. Azúcar por Imtp. Bebidas analcoholicas con edulcorante
      Q1 = "UPDATE MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc"
      Q1 = Q1 & " SET IdTipoValLib = " & LIBCOMPRAS_IMPBEBANALC
      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND IdTipoValLib =" & LIBCOMPRAS_ILABEDANALCAZUCAR
      Call ExecSQL(DbMain, Q1)
      
      'Cambiamos IVA por Adq. o Const. Inmuebles por IVA general OJO ESTO NO LO HACEMOS PORQUE PUEDE CREAR PROBLEMAS CON EL ENCABEZADO DEL DOCUMENTO
'      Q1 = "UPDATE MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc"
'      Q1 = Q1 & " SET IdTipoValLib = " & LIBCOMPRAS_IVACREDFISC
'      Q1 = Q1 & " WHERE TipoLib = " & LIB_COMPRAS & " AND IdTipoValLib =" & LIBCOMPRAS_IVAADQCONSTINMUEBLES
'      Call ExecSQL(DbMain, Q1)
   
   End If

End Sub

Private Sub AppendIdEmpresaAno(ByVal Tabla As String, Optional ByVal AppendAno As Boolean = True, Optional ByVal CreateIndex As Boolean = True)
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim Q1 As String
      
'   If Not gEmprSeparadas Then
'      Exit Sub
'   End If
   
   On Error Resume Next
   
   Set Tbl = DbMain.TableDefs(Tabla)
  
   ERR.Clear
   Tbl.Fields.Append Tbl.CreateField("IdEmpresa", dbLong)

   If ERR = 0 Then
      Tbl.Fields.Refresh
   End If
   
   Q1 = "UPDATE " & Tabla & " SET IdEmpresa = " & lEmpIdEnArchivo
   Call ExecSQL(DbMain, Q1)
   
   
   If AppendAno Then
   
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Ano", dbInteger)

      If ERR = 0 Or ERR = 3191 Then
         Tbl.Fields.Refresh
                       
         Q1 = "UPDATE " & Tabla & " SET Ano = " & lEmpAnoEnArchivo
         Call ExecSQL(DbMain, Q1)
                      
         If CreateIndex Then
            Q1 = "CREATE INDEX IdxEmpresa ON " & Tabla & " (IdEmpresa, Ano) "
            Rc = ExecSQL(DbMain, Q1, False)
         End If
         
      End If
      
   Else
      
      DbMain.TableDefs.Refresh
            
      If CreateIndex Then
         Q1 = "CREATE INDEX IdxEmpresa ON " & Tabla & " (IdEmpresa) "
         Rc = ExecSQL(DbMain, Q1, False)
      End If
      
   End If

   On Error GoTo 0
   
End Sub

'esta función verifica que el IdEMpresa y el año en cada una de las tablas corresponda y si no es así, lo actualiza
Private Sub VerificaTblEmpAno()
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim BaseConEmpAnoMalo As Boolean
   
   Q1 = "SELECT Count(*) FROM Empresa"    'si este registro está repetido, indica que la base tiene los años malos
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      If vFld(Rs(0)) > 1 Then
      
         Q1 = "DELETE * FROM Empresa WHERE Ano = " & lEmpAnoEnArchivo
         Call ExecSQL(DbMain, Q1)
         
      End If
   End If
   
   Call CloseRs(Rs)
   
'   If Not BaseConAnoMalo Then
'      Exit Sub
'   End If
   
   'aprovechamos de poner índice único
   Q1 = "CREATE UNIQUE INDEX IdxEmpresa ON Empresa (Id) "
   Rc = ExecSQL(DbMain, Q1, False)

   Q1 = "UPDATE Empresa SET Id = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   
   
   BaseConEmpAnoMalo = True
   Q1 = "SELECT IdEmpresa, Ano FROM Cuentas "    'si no hay cuentas de este año, hay que corregir
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      If vFld(Rs("IdEmpresa")) = lEmpIdEnArchivo And vFld(Rs("Ano")) = lEmpAnoEnArchivo Then
         BaseConEmpAnoMalo = False
      End If
   End If
   Call CloseRs(Rs)

   'Ahora actualizamos las tablas
   
   If Not BaseConEmpAnoMalo Then
      Exit Sub
   End If
  
   
   Q1 = "UPDATE ActFijoCompsFicha SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE ActFijoFicha SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE AjustesExtLibCaja SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE AsistImpPrimCat SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE BaseImponible14Ter SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE Cartola SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE Comprobante SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE CtasAjustesExCont SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE Cuentas SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE CuentasBasicas SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE DetCartola SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE DetSaldosAp SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE DocCuotas SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE Documento SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE ImpAdic SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE InfoAnualDJ1847 SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE LibroCaja SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE LockAction SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE LogComprobantes SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE LogImpreso SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE MovActivoFijo SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE MovComprobante SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE MovDocumento SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE PropIVA_TotMensual SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   Q1 = "UPDATE Socios SET IdEmpresa = " & lEmpIdEnArchivo & ", Ano = " & lEmpAnoEnArchivo
   Call ExecSQL(DbMain, Q1)
   

End Sub

Private Sub DelCompAperturaDuplicadosSY()     'Caso Soledad Yañez (7 sept 2020)
   Dim Q1 As String
   Dim Rs As Recordset, Rs2 As Recordset
   Dim idcomp As Long
   
   'Primero eliminamos los comprobantes de apertura financieros duplicados
   Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo = " & TC_APERTURA
   Q1 = Q1 & " AND TipoAjuste = " & TAJUSTE_FINANCIERO
   Q1 = Q1 & " ORDER BY Fecha asc"
   
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      
      idcomp = vFld(Rs("IdComp"))
   
      Q1 = "UPDATE EmpresasAno SET IdCompAper = " & idcomp
      Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
          
      Rs.MoveNext    'saltamos el primero
   
      Do While Not Rs.EOF
      
         idcomp = vFld(Rs("IdComp"))
         
         'Actualizamos IdCompAper en la tabla EmpresasAno
         Q1 = "DELETE * FROM MovComprobante WHERE IdComp = " & idcomp
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "DELETE * FROM Comprobante WHERE IdComp = " & idcomp
         Call ExecSQL(DbMain, Q1)
         
         AddLog ("Se elimina comprobante de Apertura Financiero duplicado RUTEmp=" & FmtRut(gEmpresa.Rut) & " Año=" & gEmpresa.Ano & " IdComp=" & idcomp)
         
         Rs.MoveNext
         
      Loop
         
   End If
   
   Call CloseRs(Rs)
   
   'Segundo eliminamos los comprobantes de apertura financieros duplicados
   Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo = " & TC_APERTURA
   Q1 = Q1 & " AND TipoAjuste = " & TAJUSTE_TRIBUTARIO
   Q1 = Q1 & " ORDER BY Fecha asc"
   
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
   
      idcomp = vFld(Rs("IdComp"))
   
      'Actualizamos IdCompAperTrib en la tabla EmpresasAno
      Q1 = "UPDATE EmpresasAno SET IdCompAperTrib = " & idcomp
      Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
      

      Rs.MoveNext    'saltamos el primero
   
      Do While Not Rs.EOF
      
         idcomp = vFld(Rs("IdComp"))
         
         Q1 = "DELETE * FROM MovComprobante WHERE IdComp = " & idcomp
         Call ExecSQL(DbMain, Q1)
         
         Q1 = "DELETE * FROM Comprobante WHERE IdComp = " & idcomp
         Call ExecSQL(DbMain, Q1)
         
         AddLog ("Se elimina comprobante de Apertura Tributario duplicado RUTEmp=" & FmtRut(gEmpresa.Rut) & " Año=" & gEmpresa.Ano & " IdComp=" & idcomp)
                     
         Rs.MoveNext
         
      Loop
         
   End If
   
   Call CloseRs(Rs)
     
   'Tercero eliminamos los comprobantes con valor cero y ningún movimiento
   Q1 = "SELECT IdComp FROM Comprobante WHERE Tipo <> " & TC_APERTURA
   Q1 = Q1 & " AND TotalDebe = 0 AND TotalHaber = 0"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
   
      idcomp = vFld(Rs("IdComp"))
      
      Q1 = "SELECT Count(*) FROM MovComprobante WHERE IdComp = " & idcomp
      Set Rs2 = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         If vFld(Rs(0)) = 0 Then     'no tiene movimientos, lo eliminamos
      
            Q1 = "DELETE * FROM Comprobante WHERE IdComp = " & idcomp
            Call ExecSQL(DbMain, Q1)
            
            AddLog ("Se elimina comprobante de con valor cero y sin movimeintos RUTEmp=" & FmtRut(gEmpresa.Rut) & " Año=" & gEmpresa.Ano & " IdComp=" & idcomp)
         End If
         
      End If
      
      Call CloseRs(Rs2)
      
      Rs.MoveNext
      
   Loop
           
   Call CloseRs(Rs)
   
   'Eliminamos los saldos de apertura tributarios, si no hay añbo anterior
   If gEmpresa.TieneAnoAnt = False Then
      Q1 = "UPDATE Cuentas SET DebeTrib = 0, HaberTrib = 0"
      Call ExecSQL(DbMain, Q1)
   End If
     
End Sub


'2860036 Crea tabla Membrete
Public Function CreateTblMembrete() As Boolean
   Dim Q1 As String
   Dim Rc As Long
   Dim Tbl As TableDef, Fld As Field
   
   On Error Resume Next
   
   'Creamos tabla Membrete
   
   CreateTblMembrete = True
   
   Set Tbl = New TableDef
   Tbl.Name = "Membrete"
   
   ERR.Clear
   Set Fld = Tbl.CreateField("TituloMembrete1", dbText)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblMembrete = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Membrete.TituloMembrete1", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("TituloMembrete2", dbText)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblMembrete = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Membrete.TituloMembrete2", vbExclamation
   End If
   
   ERR.Clear
   Set Fld = Tbl.CreateField("Texto1", dbText)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblMembrete = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Membrete.Texto1", vbExclamation
   End If
   
    ERR.Clear
    Set Fld = Tbl.CreateField("Texto2", dbText)
    Tbl.Fields.Append Fld

    If ERR = 0 Then
       Tbl.Fields.Refresh
    ElseIf ERR <> 3191 Then ' ya existe
       MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Membrete.Texto2", vbExclamation
       lUpdOK = False
    End If
    
    ERR.Clear
    Set Fld = Tbl.CreateField("IdEmpresa", dbInteger)
    Tbl.Fields.Append Fld

    If ERR = 0 Then
       Tbl.Fields.Refresh
    ElseIf ERR <> 3191 Then ' ya existe
       MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Membrete.IdEmpresa", vbExclamation
       lUpdOK = False
    End If
      
   DbMain.TableDefs.Append Tbl
   Set Tbl = Nothing

   DbMain.TableDefs.Refresh

End Function
'fin 2860036

'2861570
Public Function CreateTblFirmas() As Boolean
   Dim Q1 As String
   Dim Rc As Long
   Dim Tbl As TableDef, Fld As Field
   
   On Error Resume Next
   
   'Creamos tabla Membrete
   
   CreateTblFirmas = True
   
   Set Tbl = New TableDef
   Tbl.Name = "Firmas"
   
   ERR.Clear
   Set Fld = Tbl.CreateField("patch", dbText)
   Tbl.Fields.Append Fld
   
   If ERR = 0 Then
      Tbl.Fields.Refresh
   ElseIf ERR <> 3191 Then ' ya existe
      CreateTblFirmas = False
      MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Firmas.patch", vbExclamation
   End If
   
    ERR.Clear
    Set Fld = Tbl.CreateField("IdEmpresa", dbInteger)
    Tbl.Fields.Append Fld

    If ERR = 0 Then
       Tbl.Fields.Refresh
    ElseIf ERR <> 3191 Then ' ya existe
       MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Firmas.IdEmpresa", vbExclamation
       lUpdOK = False
    End If
    
     ERR.Clear
    Set Fld = Tbl.CreateField("Tipo", dbText)
    Tbl.Fields.Append Fld

    If ERR = 0 Then
       Tbl.Fields.Refresh
    ElseIf ERR <> 3191 Then ' ya existe
       MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Firmas.Tipo", vbExclamation
       lUpdOK = False
    End If
    
    ERR.Clear
    Set Fld = Tbl.CreateField("ano", dbText)
    Tbl.Fields.Append Fld

    If ERR = 0 Then
       Tbl.Fields.Refresh
    ElseIf ERR <> 3191 Then ' ya existe
       MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Firmas.ano", vbExclamation
       lUpdOK = False
    End If
      
   DbMain.TableDefs.Append Tbl
   Set Tbl = Nothing

   DbMain.TableDefs.Refresh

End Function
'fin 2861570
