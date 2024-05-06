Attribute VB_Name = "MCorrigeBaseAdmSQLServer"
Option Explicit
Private lDbVerAdm As Integer
Dim lUpdOK As Boolean

Public Sub CorrigeBaseAdmSQLServer()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Rc As Long
   
   lDbVerAdm = 0
   lUpdOK = True
   
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
      lDbVerAdm = 0
   Else
      lDbVerAdm = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)

                  
   If Not CorrigeBaseAdmSQLServer_V355() Then     'agregada 2 sep 2019
      Exit Sub
   End If
            
   If Not CorrigeBaseAdmSQLServer_V356() Then     'agregada 24 dic 2019
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V357() Then     'agregada 16 mar 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V358() Then     'agregada 31 mar 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V359() Then     'agregada 30 jul 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V360() Then     'agregada 24 ago 2020
      Exit Sub
   End If
            
   If Not CorrigeBaseAdmSQLServer_V361() Then     'agregada 31 ago 2020
      Exit Sub
   End If
            
   If Not CorrigeBaseAdmSQLServer_V362() Then     'agregada 2 sep 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V363() Then     'agregada 16 sep 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V364() Then     'agregada 23 sep 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V365() Then     'agregada 20 oct 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V366() Then     'agregada 2 nov 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V367() Then     'agregada 30 nov 2020
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V368() Then     'agregada 15 mar 20201
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V369() Then     'agregada 14 abr 2021
      Exit Sub
   End If
   
   If Not CorrigeBaseAdmSQLServer_V370() Then
      Exit Sub
   End If
   
   
   If Not CorrigeBaseAdmSQLServer_V371() Then
      Exit Sub
   End If
   
   '2814014 pipe
   If Not CorrigeBaseAdmSQLServer_V372() Then     'agregada 27 may 2022 ffv 2814014
      Exit Sub
   End If
   'fin 2814014
      
     
   If lDbVerAdm > 373 Then
      MsgBox1 "¡ ATENCION !" & vbCrLf & vbCrLf & "La base de datos corresponde a una versión posterior de este programa." & vbCrLf & "Debe actualizar el programa antes de continuar, de lo contrario podría dañar la información..", vbCritical
      Call CloseDb(DbMain)
      End
   End If
   
End Sub


'2814014
Public Function CorrigeBaseAdmSQLServer_V372() As Boolean 'ffv agregada 27 may 2022 2814014
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
         Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & MaxTipoDoc & ", 'Vale Pago Electronico con Exen', 'VPEE', 'ACTIVO', 1, 110, 111, 2, 2, 0)"
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

   CorrigeBaseAdmSQLServer_V372 = lUpdOK

End Function

Public Function CorrigeBaseAdmSQLServer_V371() As Boolean
Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
  CorrigeBaseAdmSQLServer_V371 = True
  
  
   If lDbVerAdm = 371 And lUpdOK Then
         lDbVerAdm = 372
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
  
End Function

Public Function CorrigeBaseAdmSQLServer_V370() As Boolean
Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String

CorrigeBaseAdmSQLServer_V370 = True

     If lDbVerAdm = 370 And lUpdOK Then
         lDbVerAdm = 371
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

End Function


Public Function CorrigeBaseAdmSQLServer_V369() As Boolean   'agregada 14 abr 2021
   Dim Q1 As String, Q2 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim QBase As String, QEnd As String

   On Error Resume Next

   '--------------------- Versión 368 -----------------------------------

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

   CorrigeBaseAdmSQLServer_V369 = lUpdOK
   
End Function



Public Function CorrigeBaseAdmSQLServer_V368() As Boolean   'agregada 15 mar 2021
   Dim Q1 As String, Q2 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim QBase As String, QEnd As String

   On Error Resume Next

   '--------------------- Versión 368 -----------------------------------

   If lDbVerAdm = 368 And lUpdOK = True Then
      
      'Agregamos campo CPS_INRPropiosPerdidas a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_INRPropiosPerdidas Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_UtilidadesPerdida a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_UtilidadesPerdida Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_IngresoDiferido a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_IngresoDiferido Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_CTDImputableIPE a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_CTDImputableIPE Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_IncentivoAhorro a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_IncentivoAhorro Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_IDPCVoluntario a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_IDPCVoluntario Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_CredActFijos a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_CredActFijos Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_CredParticipaciones a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_CredParticipaciones Float NULL;"
      Call ExecSQL(DbMain, Q1)
           
      
      If lUpdOK Then
         lDbVerAdm = 369
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdmSQLServer_V368 = lUpdOK
   
End Function


Public Function CorrigeBaseAdmSQLServer_V367() As Boolean   'agregada 30 nov 2020
   Dim Q1 As String, Q2 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim QBase As String, QEnd As String

   On Error Resume Next

   '--------------------- Versión 367 -----------------------------------

   If lDbVerAdm = 367 And lUpdOK = True Then
      
      'Agregamos campo AnoDesde a tabla PlanCuentasSII
      Q1 = "ALTER TABLE PlanCuentasSII ADD AnoDesde smallint NULL;"
      Call ExecSQL(DbMain, Q1)
          
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

   CorrigeBaseAdmSQLServer_V367 = lUpdOK
   
End Function



Public Function CorrigeBaseAdmSQLServer_V366() As Boolean   'agregada 2 nov 2020
   Dim Q1 As String, Q2 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 366 -----------------------------------

   If lDbVerAdm = 366 And lUpdOK = True Then
      
      'Agregamos campo CPS_CapPropioTrib a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_CapPropioTrib Float NULL;"
      Call ExecSQL(DbMain, Q1)
            
      'Agregamos campo CPS_CapPropioTribAnoAnt a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_CapPropioTribAnoAnt Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_RepPerdidaArrastre a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_RepPerdidaArrastre Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_CapPropioSimplVarAnual a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_CapPropioSimplVarAnual Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      
      If lUpdOK Then
         lDbVerAdm = 367
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdmSQLServer_V366 = lUpdOK
   
End Function


Public Function CorrigeBaseAdmSQLServer_V365() As Boolean   'agregada 20 oct 2020
   Dim Q1 As String, Q2 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 365 -----------------------------------

   If lDbVerAdm = 365 And lUpdOK = True Then
      
      'Agregamos campo CPS_AumentosCapital a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_AumentosCapital Float NULL;"
      Call ExecSQL(DbMain, Q1)
            
      'Agregamos campo CPS_GastosRechazadosNoPagan40 a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_GastosRechazadosNoPagan40 Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_INRPropios a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_INRPropios Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_OtrosAjustesAumentos a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_OtrosAjustesAumentos Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_OtrosAjustesDisminuciones a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_OtrosAjustesDisminuciones Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
            
      
      If lUpdOK Then
         lDbVerAdm = 366
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdmSQLServer_V365 = lUpdOK
   
End Function


Public Function CorrigeBaseAdmSQLServer_V364() As Boolean   'agregada 23 sept 2020
   Dim Q1 As String, Q2 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 364 -----------------------------------

   If lDbVerAdm = 364 And lUpdOK = True Then
      
      'Agregamos campo CPS_CapitalAportado a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_CapitalAportado Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_BaseImpPrimCat_14DN3 a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_BaseImpPrimCat_14DN3 Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_BaseImpPrimCat_14DN8 a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_BaseImpPrimCat_14DN8 Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_Participaciones a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_Participaciones Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_Disminuciones a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_Disminuciones Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_GastosRechazados a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_GastosRechazados Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_RetirosDividendos a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_RetirosDividendos Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo CPS_CapPropioSimplificado a tabla EmpresasAno
      Q1 = "ALTER TABLE EmpresasAno ADD CPS_CapPropioSimplificado Float NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos tabla CtasAjustesExContRLI
      Q1 = "CREATE TABLE CapPropioSimplAnual ("
      Q1 = Q1 & " IdCapPropioSimplAnual int IDENTITY (1,1) NOT NULL,"
      Q1 = Q1 & " IdEmpresa int,"
      Q1 = Q1 & " TipoDetCPS tinyint, "
      Q1 = Q1 & " IngresoManual bit, "
      Q1 = Q1 & " AnoValor smallint,"
      Q1 = Q1 & " Valor float,"
      Q1 = Q1 & " CONSTRAINT IdxCapPropioSimplAnual PRIMARY KEY (IdCapPropioSimplAnual) "
      Q1 = Q1 & ");"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE UNIQUE INDEX IdxEmpCapAnual ON CapPropioSimplAnual (IdEmpresa, TipoDetCPS, AnoValor )"
      Call ExecSQL(DbMain, Q1)
      
      
      If lUpdOK Then
         lDbVerAdm = 365
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdmSQLServer_V364 = lUpdOK
   
End Function

Public Function CorrigeBaseAdmSQLServer_V363() As Boolean   'agregada 16 sep 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


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

   CorrigeBaseAdmSQLServer_V363 = lUpdOK
   
End Function


Public Function CorrigeBaseAdmSQLServer_V362() As Boolean   'agregada 2 sep 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


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

   CorrigeBaseAdmSQLServer_V362 = lUpdOK
   
End Function

Public Function CorrigeBaseAdmSQLServer_V361() As Boolean   'agregada 31 ago 2020
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

   CorrigeBaseAdmSQLServer_V361 = lUpdOK
   
End Function


Public Function CorrigeBaseAdmSQLServer_V360() As Boolean   'agregada 24 ago 2020
   Dim Q1 As String, Q2 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 360 -----------------------------------

   If lDbVerAdm = 360 And lUpdOK = True Then
      
      'Agregamos campo TipoIVARetenido a tabla TipoValor
      Q1 = "ALTER TABLE TipoValor ADD TipoIVARetenido tinyint NULL;"
      Call ExecSQL(DbMain, Q1)
      
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

   CorrigeBaseAdmSQLServer_V360 = lUpdOK
   
End Function
Public Function CorrigeBaseAdmSQLServer_V359() As Boolean   'agregada 30 jul 2020
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

   CorrigeBaseAdmSQLServer_V359 = lUpdOK
   
End Function

Public Function CorrigeBaseAdmSQLServer_V358() As Boolean   'agregada 31 mar 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim QBase As String, QEnd As String


   On Error Resume Next

   '--------------------- Versión 358 -----------------------------------

   If (lDbVerAdm = 358) And lUpdOK = True Then
      
      Err.Clear
      
      'Agregamos tabla FactorActAnual
      Q1 = "CREATE TABLE FactorActAnual ("
      Q1 = Q1 & " IdFactorActAnual int IDENTITY (1,1) NOT NULL,"
      Q1 = Q1 & " Ano smallint, "
      Q1 = Q1 & " MesRow tinyint,"
      Q1 = Q1 & " MesCol tinyint,"
      Q1 = Q1 & " Factor float, "
      Q1 = Q1 & " CONSTRAINT Idx_FactorActAnual PRIMARY KEY (IdFactorActAnual) "
      Q1 = Q1 & ");"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE UNIQUE INDEX IdxAno_FactorActAnual ON FactorActAnual (Ano, MesRow, MesCol)"
      Call ExecSQL(DbMain, Q1)
         
      If lUpdOK Then
         lDbVerAdm = 359
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdmSQLServer_V358 = lUpdOK
   
End Function

Public Function CorrigeBaseAdmSQLServer_V357() As Boolean   'agregada 16 mar 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim QBase As String, QEnd As String


   On Error Resume Next

   '--------------------- Versión 357 -----------------------------------

   If (lDbVerAdm = 357) And lUpdOK = True Then
      
      Err.Clear
      
      'Agregamos campo aIPC a tabla IPC (IPC acumulado)
      Q1 = "ALTER TABLE IPC ADD aIPC FLOAT NULL;"
      Call ExecSQL(DbMain, Q1)
         
      If lUpdOK Then
         lDbVerAdm = 358
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdmSQLServer_V357 = lUpdOK
   
End Function


Public Function CorrigeBaseAdmSQLServer_V356() As Boolean   'agregada 24 dic 2019
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

   CorrigeBaseAdmSQLServer_V356 = lUpdOK
   
End Function

Public Function CorrigeBaseAdmSQLServer_V355() As Boolean   'agregada 2 sep 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset


   On Error Resume Next

   '--------------------- Versión 355 -----------------------------------

   If (lDbVerAdm = 354 Or lDbVerAdm = 355) And lUpdOK = True Then    ' se ajusta la versión de la DB a la versión de la DB de LPContab Access
      
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
      Call DeleteSQL(DbMain, "Regiones", "WHERE Comuna = 'Ranquil'")
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Trehuaco'"
      Call DeleteSQL(DbMain, "Regiones", "WHERE Comuna = 'Trehuaco'")
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Bulnes'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Chillan Viejo'"
      Call DeleteSQL(DbMain, "Regiones", "WHERE Comuna = 'Chillan Viejo'")
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Chillan'"
      Call DeleteSQL(DbMain, "Regiones", "WHERE Comuna = 'Chillan'")
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'El Carmen'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Pemuco'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Pinto'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Quillon'"
      Call DeleteSQL(DbMain, "Regiones", "WHERE Comuna = 'Quillon'")
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'San Ignacio'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Yungay'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Coihueco'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'Ñiquen'"
      Call DeleteSQL(DbMain, "Regiones", "WHERE Comuna = 'Ñiquen'")
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'San Carlos'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'San Fabian'"
      Call DeleteSQL(DbMain, "Regiones", "WHERE Comuna = 'San Fabian'")
      
      Q1 = "UPDATE Regiones SET Codigo = '16' WHERE Comuna = 'San Nicolas'"
      Call ExecSQL(DbMain, Q1)
      
      
      If lUpdOK Then
         lDbVerAdm = 356
         Q1 = "UPDATE Param SET Valor=" & lDbVerAdm & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdmSQLServer_V355 = lUpdOK
   
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
'      Set Tbl = DbMain.TableDefs("Empresas")
     
      Err.Clear
     ' Tbl.Fields.Append Tbl.CreateField("Import", dbLong)

      Q1 = ""
      Q1 = Q1 & "IF NOT EXISTS( "
      Q1 = Q1 & "SELECT * FROM INFORMATION_SCHEMA.COLUMNS "
      Q1 = Q1 & "WHERE TABLE_NAME = 'Empresas' AND COLUMN_NAME = 'Import' "
      Q1 = Q1 & ")BEGIN "
      Q1 = Q1 & "ALTER TABLE Empresas ADD Import INT NULL; "
      Q1 = Q1 & "END "
      
      Call ExecSQL(DbMain, Q1)
        
      If Err = 0 Then
         'Tbl.Fields.Refresh
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

