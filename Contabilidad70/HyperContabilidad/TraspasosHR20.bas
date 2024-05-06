Attribute VB_Name = "TraspasosHR20"
Option Explicit

'Tipos de Partidas de Resultado
Public Const MAX_TIPOPARTIDA = 19

Type Partida_t
   Codigo As Integer
   Partida As String
   IngEgr As String
   AnoDesde As Integer
   AnoHasta As Integer
   SoloArt14A As Boolean
End Type

Public gTipoPartida(MAX_TIPOPARTIDA) As Partida_t

Type TipoContribHT_t
   CodTContribHR As Integer
   TContribHR As String
   CodTContribConta As Integer
End Type

Public Const MAX_TIPOCONTROBHR = 9
Public gTipoContribHR(MAX_TIPOCONTROBHR) As TipoContribHT_t


Public Function IniTraspasos20()

   Call IniTipoPartida
   
End Function

Public Function IniTipoPartida()

   gTipoPartida(1).Partida = "01- Ingresos del Giro Percibidos o Devengados"
   gTipoPartida(1).IngEgr = "I"
   gTipoPartida(1).Codigo = 628
   
   gTipoPartida(2).Partida = "02- Rentas de Fuente Extranjera"
   gTipoPartida(2).IngEgr = "I"
   gTipoPartida(2).Codigo = 851
   
   gTipoPartida(3).Partida = "03- Intereses Percibidos o Devengados"
   gTipoPartida(3).IngEgr = "I"
   gTipoPartida(3).Codigo = 629
   
   gTipoPartida(4).Partida = "04- Otros Ingresos Percibidos o Devengados"
   gTipoPartida(4).IngEgr = "I"
   gTipoPartida(4).Codigo = 651
   
   gTipoPartida(5).Partida = "05- Costo Directo de los Bienes y Servicios"
   gTipoPartida(5).IngEgr = "E"
   gTipoPartida(5).Codigo = 630
   
   gTipoPartida(6).Partida = "06- Remuneraciones"
   gTipoPartida(6).IngEgr = "E"
   gTipoPartida(6).Codigo = 631
   
   gTipoPartida(7).Partida = "08- Depreciación Financiera del ejercicio"
   gTipoPartida(7).IngEgr = "E"
   gTipoPartida(7).Codigo = 632
   
   gTipoPartida(8).Partida = "09- Intereses Pagados o Adeudados"
   gTipoPartida(8).IngEgr = "E"
   gTipoPartida(8).Codigo = 633
   
   gTipoPartida(9).Partida = "10- Gastos por Donaciones"
   gTipoPartida(9).IngEgr = "E"
   gTipoPartida(9).Codigo = 966
   
   gTipoPartida(10).Partida = "11- Otros Gastos Financieros"
   gTipoPartida(10).IngEgr = "E"
   gTipoPartida(10).Codigo = 967
   
   gTipoPartida(11).Partida = "12- Gastos por Inversión en Investigación y Desarrollo certificados por Corfo"
   gTipoPartida(11).IngEgr = "E"
   gTipoPartida(11).Codigo = 852
   
   gTipoPartida(12).Partida = "13- Gastos por Inversión en Investigación y Desarrollo no certificados por Corfo"
   gTipoPartida(12).IngEgr = "E"
   gTipoPartida(12).Codigo = 897
   
   gTipoPartida(13).Partida = "16- Costos y Gastos necesarios para producir las Rentas de Fuente Extranjera"
   gTipoPartida(13).IngEgr = "E"
   gTipoPartida(13).Codigo = 853
   
   gTipoPartida(14).Partida = "17- Gastos por Impuesto Renta e Impuesto Diferido"
   gTipoPartida(14).IngEgr = "E"
   gTipoPartida(14).Codigo = 968
   
   gTipoPartida(15).Partida = "18- Gastos por adquisición en supermercados y negocios similares"
   gTipoPartida(15).IngEgr = "E"
   gTipoPartida(15).SoloArt14A = True
   gTipoPartida(15).Codigo = 969
   
   gTipoPartida(16).Partida = "19- Otros Gastos Deducidos de los Ingresos Brutos"
   gTipoPartida(16).IngEgr = "E"
   gTipoPartida(16).Codigo = 635
   
   gTipoPartida(17).Partida = "07- Arriendos"
   gTipoPartida(17).IngEgr = "E"
   gTipoPartida(17).AnoDesde = 2020
   gTipoPartida(17).Codigo = 1140
   
   gTipoPartida(18).Partida = "14- Gastos por exigencias medioambientales"
   gTipoPartida(18).IngEgr = "E"
   gTipoPartida(18).AnoDesde = 2020
   gTipoPartida(18).Codigo = 1141
   
   gTipoPartida(19).Partida = "15- Gasto por indemnización o compensación a clientes o usuarios"
   gTipoPartida(19).IngEgr = "E"
   gTipoPartida(19).AnoDesde = 2020
   gTipoPartida(19).Codigo = 1142


End Function


Public Function InitTipoContribHR()

   'Tipo de contribuyente

   gTipoContribHR(1).CodTContribHR = 1
   gTipoContribHR(1).TContribHR = "Soc. de personas primera categoría"
   gTipoContribHR(1).CodTContribConta = CONTRIB_PRIMCAT
   
   gTipoContribHR(2).CodTContribHR = 2
   gTipoContribHR(2).TContribHR = "Soc. anónima abierta"
   gTipoContribHR(2).CodTContribConta = CONTRIB_SAABIERTA
   
   gTipoContribHR(3).CodTContribHR = 4
   gTipoContribHR(3).TContribHR = "Soc. anónima cerrada"
   gTipoContribHR(3).CodTContribConta = CONTRIB_SACERRADA
   
   gTipoContribHR(4).CodTContribHR = 5
   gTipoContribHR(4).TContribHR = "Soc. por acción"
   gTipoContribHR(4).CodTContribConta = CONTRIB_SPORACCION
   
   gTipoContribHR(5).CodTContribHR = 6
   gTipoContribHR(5).TContribHR = "Emp. Ind. Responsabilidad ltda. (EIRL)"
   gTipoContribHR(5).CodTContribConta = CONTRIB_EMPINDIVIDUALEIRL
   
   gTipoContribHR(6).CodTContribHR = 7
   gTipoContribHR(6).TContribHR = "Emp. Individual - Persona Natural"
   gTipoContribHR(6).CodTContribConta = CONTRIB_EMPINDIVIDUAL
   
   gTipoContribHR(7).CodTContribHR = 8
   gTipoContribHR(7).TContribHR = "Soc. de segunda categoría"
   gTipoContribHR(7).CodTContribConta = 0
   
   gTipoContribHR(8).CodTContribHR = 9
   gTipoContribHR(8).TContribHR = "Soc. comandita por acción"
   gTipoContribHR(8).CodTContribConta = 0
   
   gTipoContribHR(9).CodTContribHR = 10
   gTipoContribHR(9).TContribHR = "Establecimiento Permanente"
   gTipoContribHR(9).CodTContribConta = CONTRIB_ESTABPERMANENTE
   
   gTipoContribHR(9).CodTContribHR = 11
   gTipoContribHR(9).TContribHR = "Comunidades"
   gTipoContribHR(9).CodTContribConta = CONTRIB_COMUNIDAD
   
   gTipoContribHR(9).CodTContribHR = 12
   gTipoContribHR(9).TContribHR = "Cooperativas"
   gTipoContribHR(9).CodTContribConta = CONTRIB_COOPERATIVAS
   
   gTipoContribHR(9).CodTContribHR = 13
   gTipoContribHR(9).TContribHR = "OSFL"
   gTipoContribHR(9).CodTContribConta = CONTRIB_ORGSINFINESDELUCRO
   
End Function
Public Function GetDatosEmpHR(ByVal RutEmpHR As String) As Boolean
   Dim Q1 As String
   Dim Rs As dao.Recordset
   Dim Rc As Long
   Dim i As Integer
   Dim HrDb As Database
   Dim DbPath As String
   Dim IdContrib As Long
   Dim TContribHR As String, IdTipoContrib As Integer
   Dim IdRegion As Long, IdComuna As Long, IdActividad As Long
   Dim bTblExist As Boolean
   Dim MsgHRIncompatible As Integer
   
   GetDatosEmpHR = False
   
   On Error Resume Next
   DbPath = gHRPath & "\PAR\BD_HR_admin.mdb"
   If ExistFile(DbPath) = False Then
'      MsgBox1 "No se encuentra la base de HR en" & vbCrLf & lDbPath, vbExclamation
      Exit Function
   End If
         
   Set HrDb = OpenDatabase(DbPath, False, False, ";PWD=" & "20" & "080" & "3hr" & ";")
   
   If ERR Then
      MsgBox "Error H" & Hex(ERR) & ", " & Error & NL & DbPath, vbExclamation
      Exit Function
   End If

   On Error Resume Next
   
   Error.Clear
   
   For i = 0 To HrDb.TableDefs.Count - 1
      If StrComp(HrDb.TableDefs(i).Name, "ADM_REGION_CONTRIB", vbTextCompare) = 0 Then
         bTblExist = True
         Exit For
      End If
   Next i
   
   If Not bTblExist Then
   
      MsgHRIncompatible = Val(GetIniString(gIniFile, "Msg", "HRIncompatible", "0"))
      
      If MsgHRIncompatible = 0 Then
         If MsgBox1("Versión de HR incompatible con esta versión de " & APP_FULLNAME & vbCrLf & vbCrLf & "¿Desea volver a ver este mensaje?", vbInformation + vbYesNo) = vbNo Then
            Call SetIniString(gIniFile, "Msg", "HRIncompatible", "1")
         End If
      End If
      
      Exit Function
   End If
   
   'Datos básicos
   Q1 = "SELECT c.*, Adm_Comuna_Contrib.Id_Comuna as Id_Comuna, Adm_Region_Contrib.Id_Region As Id_Region, Adm_Actividad_Contrib.Id_Activ as Id_Activ FROM ((Adm_NContrib as c "
   Q1 = Q1 & " LEFT JOIN Adm_Comuna_Contrib ON c.Id_Contrib = Adm_Comuna_Contrib.Id_Contrib)"
   Q1 = Q1 & " LEFT JOIN Adm_Region_Contrib ON c.Id_Contrib = Adm_Region_Contrib.Id_Contrib)"
   Q1 = Q1 & " LEFT JOIN Adm_Actividad_Contrib ON c.Id_Contrib = Adm_Actividad_Contrib.Id_Contrib"
   Q1 = Q1 & " WHERE c.NC_Rut='" & Right("0" & RutEmpHR, 8) & "-" & DV_Rut(RutEmpHR) & "'"
   
   Set Rs = OpenRsDao(HrDb, Q1)
   
   If ERR Then
      Call CloseRs(Rs)
      Call CloseDb(HrDb)
      Exit Function
   End If
   
   On Error GoTo 0
   
   If Rs Is Nothing Then
      Call CloseRs(Rs)
      
      Call CloseDb(HrDb)
      Exit Function
   End If
   
   If Rs.EOF = False Then
      gEmprHR.EmpConta.Rut = RutEmpHR
      gEmprHR.EmpConta.NombreCorto = Trim(vFldDao(Rs("NC_NomCorto")))
      gEmprHR.EmpConta.RazonSocial = Trim(vFldDao(Rs("NC_Paterno")) & " " & vFldDao(Rs("NC_Materno")) & " " & vFldDao(Rs("NC_Nombre")))
      gEmprHR.ApMaterno = Trim(vFldDao(Rs("NC_Materno")))
      gEmprHR.EmpConta.Nombre = Trim(vFldDao(Rs("NC_Nombre")))
      
      gEmprHR.EmpConta.Direccion = vFldDao(Rs("NC_Calle"), True)
      gEmprHR.NroCalle = vFldDao(Rs("NC_Nro"))
      gEmprHR.NroDepto = vFldDao(Rs("NC_Depto"))
      
'      If Trim(vFldDao(Rs("NC_Depto"))) <> "" Then
'         gEmprHR.EmpConta.Direccion = gEmprHR.EmpConta.Direccion & " dpto. " & vFld(Rs("NC_Depto"), True)
'      End If
      
      IdContrib = vFldDao(Rs("Id_Contrib"))
      gEmprHR.EmpConta.Telefono = vFldDao(Rs("NC_Fono"))
      gEmprHR.EmpConta.Fax = vFldDao(Rs("NC_Fax"))
      gEmprHR.EmpConta.Ciudad = vFldDao(Rs("NC_Ciudad"))
      
      gEmprHR.EmpConta.Villa = vFldDao(Rs("NC_Villa"))
      gEmprHR.EmpConta.Celular = vFldDao(Rs("NC_Celular"))
      gEmprHR.EmpConta.CodArea = vFldDao(Rs("NC_CodArea"))
      
      'region y comuna
      IdRegion = vFldDao(Rs("Id_Region"))
      IdComuna = vFldDao(Rs("Id_Comuna"))
      
      'actividad económica
      IdActividad = vFldDao(Rs("Id_Activ"))
      
      gEmprHR.EmpConta.email = vFldDao(Rs("NC_Correo"))
'      gEmprHR.Web = ""
      gEmprHR.EmpConta.Giro = vFldDao(Rs("NC_Giro"))
      
      gEmprHR.NombContador = vFldDao(Rs("NC_NomCont")) & " " & vFldDao(Rs("NC_PatCont")) & " " & vFldDao(Rs("NC_MatCont"))
      gEmprHR.RutContador = FmtCID(vFmtCID(vFldDao(Rs("NC_RutCont"))))
      gEmprHR.DirPostal = vFldDao(Rs("NC_Dir_Postal"))
      
      GetDatosEmpHR = True
   
   End If
   
   Call CloseRs(Rs)
   
   'nombres region y comuna
   If IdRegion > 0 Then
   
      Q1 = "SELECT Reg_Nombre FROM Adm_Region2019 WHERE Id_Region = " & IdRegion
         
      Set Rs = OpenRsDao(HrDb, Q1)
      gEmprHR.Region = vFldDao(Rs("Reg_Nombre"))
      
      Call CloseRs(Rs)
   End If
   
   If IdComuna > 0 Then
   
      Q1 = "SELECT Com_Nombre FROM Adm_Comuna WHERE Id_Comuna = " & IdComuna
         
      Set Rs = OpenRsDao(HrDb, Q1)
      gEmprHR.EmpConta.Comuna = vFldDao(Rs("Com_Nombre"))
      
      Call CloseRs(Rs)
   End If

   'Actividad económica
  If IdActividad > 0 Then
   
      Q1 = "SELECT Act_Codigo FROM Adm_Actividad WHERE Id_Activ = " & IdActividad
         
      Set Rs = OpenRsDao(HrDb, Q1)
      gEmprHR.EmpConta.CodActEcono = vFldDao(Rs("Act_Codigo"))
      
      Call CloseRs(Rs)
   End If

   'dirección y comuna postal
   Q1 = "SELECT c.Id_Comuna, Com_Nombre FROM ( Adm_ComunaP_Contrib as c "
   Q1 = Q1 & " LEFT JOIN Adm_Comuna ON c.Id_Comuna = Adm_Comuna.Id_Comuna )"
   Q1 = Q1 & " WHERE c.Id_Contrib=" & IdContrib

   Set Rs = OpenRsDao(HrDb, Q1)
      
   If Rs.EOF = False Then
      gEmprHR.ComunaPostal = vFldDao(Rs("Com_Nombre"))
   End If
   
   Call CloseRs(Rs)

   'Rep Legal
   Q1 = "SELECT Adm_Rep_Legal.* FROM Adm_Rep_Legal INNER JOIN Adm_Rep_Contrib ON Adm_Rep_Legal.Id_Rep = Adm_Rep_Contrib.Id_Rep"
   Q1 = Q1 & " WHERE Adm_Rep_Contrib.Id_Contrib=" & IdContrib & " AND Adm_Rep_Legal.Rep_Estado <> 0"
   Q1 = Q1 & " ORDER BY Adm_Rep_Legal.Id_Rep "

   Set Rs = OpenRsDao(HrDb, Q1)
   
   If Rs.EOF = False Then

      gEmprHR.EmpConta.RutRepLegal1 = vFldDao(Rs("Rep_Rut"))
      gEmprHR.EmpConta.RepLegal1 = Trim(vFldDao(Rs("Rep_Nombre")) & " " & vFldDao(Rs("Rep_Paterno")) & " " & vFldDao(Rs("Rep_Materno")))
   
      GetDatosEmpHR = True

   End If
   
   Call CloseRs(Rs)
   
   'Tipo Contribuyente
   If gTipoContribHR(1).CodTContribHR = 0 Then
      Call InitTipoContribHR
   End If
   
   Q1 = "SELECT Adm_NContrib.Id_TipoContrib, Adm_Tipo_Contrib.TC_Descripcion FROM Adm_NContrib INNER JOIN Adm_Tipo_Contrib ON Adm_NContrib.Id_TipoContrib = Adm_Tipo_Contrib.Id_TipoContrib"
   Q1 = Q1 & " WHERE Adm_NContrib.Id_Contrib=" & IdContrib

   Set Rs = OpenRsDao(HrDb, Q1)
   
   If Rs.EOF = False Then

      gEmprHR.TipoContrib = 0
      
      TContribHR = vFldDao(Rs("TC_Descripcion"))
      IdTipoContrib = vFldDao(Rs("Id_TipoContrib"))
      For i = 1 To UBound(gTipoContribHR)
         If gTipoContribHR(i).CodTContribHR = IdTipoContrib Then
            gEmprHR.TipoContrib = gTipoContribHR(i).CodTContribConta
         End If
      Next i
   
      GetDatosEmpHR = True

   End If
   
   Call CloseRs(Rs)
   
   
   'Franquicias Tributarias
   Q1 = "SELECT Const_CP_Nombre, Const_CP_Valor   FROM Adm_Constante_Contrib_Prod"
   Q1 = Q1 & " WHERE Adm_Constante_Contrib_Prod.Id_Contrib=" & IdContrib
   Q1 = Q1 & " AND Const_Cp_Nombre IN ('TRANSA_BOLSA', 'FORMA_RENTA', '14BIS', 'LEY', 'DL600', 'DL701', 'DS341')"

   Set Rs = OpenRsDao(HrDb, Q1)
   
   For i = 0 To UBound(gEmprHR.Franquicias)
      gEmprHR.Franquicias(i) = False
   Next i
   gEmprHR.TransaBolsa = False

   Do While Not Rs.EOF

      Select Case vFldDao(Rs("Const_CP_Nombre"))
      
         Case "TRANSA_BOLSA"
            If Val(vFldDao(Rs("Const_CP_Valor"))) <> 0 Then
               gEmprHR.TransaBolsa = True
            End If
         
'         Case "FORMA_RENTA"
            
         Case "14BIS"
            If Val(vFldDao(Rs("Const_CP_Valor"))) <> 0 Then
               gEmprHR.Franquicias(FRANQ_14BIS) = True
            End If
         
         Case "LEY"
            If Val(vFldDao(Rs("Const_CP_Valor"))) <> 0 Then
               gEmprHR.Franquicias(FRANQ_LEY18392) = True
            End If
         
         Case "DL600"
            If Val(vFldDao(Rs("Const_CP_Valor"))) <> 0 Then
               gEmprHR.Franquicias(FRANQ_DL600) = True
            End If
         
         Case "DL701"
            If Val(vFldDao(Rs("Const_CP_Valor"))) <> 0 Then
               gEmprHR.Franquicias(FRANQ_DL701) = True
            End If
         
         Case "DS341"
            If Val(vFldDao(Rs("Const_CP_Valor"))) <> 0 Then
               gEmprHR.Franquicias(FRANQ_DS341) = True
            End If
         
      End Select
         
      GetDatosEmpHR = True
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   
   'Tipo Regimen
   Q1 = "SELECT Adm_Reg_NContrib.Id_TipoRegim, Tr_Descripcion FROM Adm_Reg_NContrib INNER JOIN Adm_Tipo_Regimen ON Adm_Reg_NContrib.Id_TipoRegim = Adm_Tipo_Regimen.Id_TipoRegim "
   Q1 = Q1 & " WHERE Adm_Reg_NContrib.Id_Contrib=" & IdContrib
   Q1 = Q1 & " AND Rn_Anio = " & gEmpresa.Ano + 1

   Set Rs = OpenRsDao(HrDb, Q1)
   
   gEmprHR.EmpConta.RentaAtribuida = False
   gEmprHR.EmpConta.SemiIntegrado = False
   gEmprHR.EmpConta.Franq14Ter = False
   gEmprHR.Franquicias(FRANQ_14TER) = False

   
   If Not Rs.EOF Then
   
      Select Case vFldDao(Rs("Id_TipoRegim"))
         Case 1
            gEmprHR.EmpConta.RentaAtribuida = True
         Case 2
            gEmprHR.EmpConta.SemiIntegrado = True
         Case 3
            gEmprHR.EmpConta.Franq14Ter = True
            gEmprHR.Franquicias(FRANQ_14TER) = True
      End Select
               
   End If
   
   Call CloseRs(Rs)
   
   gEmprHR.EmpConta.R14ASemiIntegrado = False
   gEmprHR.EmpConta.ProPymeGeneral = False
   gEmprHR.EmpConta.ProPymeTransp = False
   gEmprHR.EmpConta.RentasPresuntas = False
   gEmprHR.EmpConta.RentaEfectiva = False
   gEmprHR.EmpConta.RegimenOtro = False
   gEmprHR.EmpConta.NoSujetoArt14 = False

   
   Q1 = "SELECT Adm_Reg_NContrib.Id_TipoRegim, Tr_Descripcion FROM Adm_Reg_NContrib INNER JOIN Adm_Tipo_Regimen2021 ON Adm_Reg_NContrib.Id_TipoRegim = Adm_Tipo_Regimen2021.Id_TipoRegim "
   Q1 = Q1 & " WHERE Adm_Reg_NContrib.Id_Contrib=" & IdContrib
   Q1 = Q1 & " AND Rn_Anio = " & gEmpresa.Ano + 1

   Set Rs = OpenRsDao(HrDb, Q1, False)
   
   If Not Rs Is Nothing Then
   
      If Not Rs.EOF Then
       
         Select Case vFldDao(Rs("Id_TipoRegim"))
            Case 1
               gEmprHR.EmpConta.R14ASemiIntegrado = True
            Case 2
               gEmprHR.EmpConta.ProPymeGeneral = True
            Case 3
               gEmprHR.EmpConta.ProPymeTransp = True
            Case 4
               gEmprHR.EmpConta.RentasPresuntas = True
            Case 5
               gEmprHR.EmpConta.RentaEfectiva = True
            Case 6
               gEmprHR.EmpConta.RegimenOtro = True
            Case 7
               gEmprHR.EmpConta.NoSujetoArt14 = True
         End Select
                  
      End If
      
      Call CloseRs(Rs)
      
   End If
   
   Call CloseDb(HrDb)
   

End Function

