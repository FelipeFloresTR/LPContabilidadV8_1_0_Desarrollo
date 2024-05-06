Attribute VB_Name = "ExpHRDJ"
Option Explicit

Public Const DJ1847_ES_SVS = 1
Public Const DJ1847_ES_SBIF = 2
Public Const DJ1847_ES_OTRA = 3
Public Const DJ1847_ES_NOAPLICA = 4
Public Const DJ1847_ES_CMF = 5

Public Const DJ1847_RLI_POSEEAJUSTES = 1
Public Const DJ1847_RLI_NOPOSEEAJUSTES = 2

Public gDJ1847_EntSupervisora(DJ1847_ES_CMF) As String
Public gDJ1847_AjusteRLI(DJ1847_RLI_NOPOSEEAJUSTES) As String

Public Function InitDJ1847()

   gDJ1847_EntSupervisora(DJ1847_ES_SVS) = " SVS"
   gDJ1847_EntSupervisora(DJ1847_ES_SBIF) = " SBIF"
   gDJ1847_EntSupervisora(DJ1847_ES_OTRA) = "OTRA"
   gDJ1847_EntSupervisora(DJ1847_ES_NOAPLICA) = "NO APLICA"
   gDJ1847_EntSupervisora(DJ1847_ES_CMF) = "CMF"
   
   gDJ1847_AjusteRLI(DJ1847_RLI_POSEEAJUSTES) = "Posee Ajustes en la RLI"
   gDJ1847_AjusteRLI(DJ1847_RLI_NOPOSEEAJUSTES) = "No Posee Ajustes en la RLI"

End Function

'Exporta para DJ1923. Genera 1 archivo de acuerdo a formato especificado
Public Function Export_DJ1923(fname As String, Optional ByVal HRRAB As Boolean = False) As Long
   Dim FPath As String
   Dim LogPath As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String
   Dim i As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Valor As Double
   Dim r As Integer
   Dim ExpDir As String
   Dim Msg As String
   Dim Codigo As Integer
   
   
   On Error Resume Next
      
   Sep = ";"

'   ExpDir = "\HRDJ\"
'   If HRRAB Then
'      ExpDir = "\HRRAB\"
'   End If
'
'   FPath = gExportPath & ExpDir
'
'   MkDir FPath
'   FPath = gExportPath & ExpDir & gEmpresa.Rut
'   MkDir FPath
   
   
   ExpDir = gHRPath & "\RUTS"
   MkDir ExpDir
      
   ExpDir = ExpDir & "\" & Right("00000000" & gEmpresa.Rut, 8)
   MkDir ExpDir
      
   ExpDir = ExpDir & "\ImpConta"
   MkDir ExpDir

   If HRRAB Then
      If gEmpresa.Ano >= 2020 Then
         fname = "HRRAD_" & Right(gEmpresa.Ano, 2) & ".csv"
      Else
         fname = "HRRAB_" & Right(gEmpresa.Ano, 2) & ".csv"
      End If
   Else
'      fname = "DJ1923_" & gEmpresa.Ano & ".txt"
      fname = "DJ1923_" & Right(gEmpresa.Ano, 2) & ".csv"
   End If
   
   FPath = ExpDir & "\" & fname
   
   Fd = FreeFile
   ERR.Clear
   
   Open FPath For Output As #Fd
   If ERR Then
      MsgErr FPath
      Export_DJ1923 = -ERR
      Exit Function
   End If

   On Error GoTo 0

   'seleccionamos los registros
   Q1 = "SELECT TipoPartida, Sum(MovComprobante.Debe) as SumDebe, Sum(MovComprobante.Haber) As SumHaber FROM (MovComprobante "
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovComprobante.IdCuenta = Cuentas.IdCuenta"
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante")
   Q1 = Q1 & " WHERE Comprobante.Fecha BETWEEN " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " AND " & CLng(DateSerial(gEmpresa.Ano, 12, 31))
   Q1 = Q1 & " AND ( Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
   Q1 = Q1 & " AND Comprobante.Estado = " & EC_APROBADO
   Q1 = Q1 & " AND (Cuentas.TipoPartida IS NOT NULL AND Cuentas.TipoPartida <> 0)"
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY TipoPartida"
   Q1 = Q1 & " ORDER BY TipoPartida"
   Set Rs = OpenRs(DbMain, Q1)
   
   Buf = "Concepto o Partida" & Sep & "N° Concepto" & Sep & "Monto"

'   Print #Fd, Buf

   Buf = ""
   r = 0
   
   'imprimimos el archivo
   Do While Rs.EOF = False
            
      If gTipoPartida(vFld(Rs("TipoPartida"))).IngEgr = "E" Then    'Egreso
         Valor = vFld(Rs("SumDebe")) - vFld(Rs("SumHaber"))
      Else                                                           'Ingreso
         Valor = vFld(Rs("SumHaber")) - vFld(Rs("SumDebe"))
      End If
         
      If Valor < 0 Then
         Valor = 0
      End If

      If Valor > 0 Then
         Codigo = IIf(gEmpresa.Ano <= 2019, vFld(Rs("TipoPartida")), HomologaCod14D(gTipoPartida(vFld(Rs("TipoPartida"))).Codigo)) 'Esto es por incosistencia de HR que de un año para otro cambió el tipo de código que recibe
         
         Buf = Left(gTipoPartida(vFld(Rs("TipoPartida"))).Partida, 10) & Sep & Codigo & Sep & Valor
         Print #Fd, Buf
         r = r + 1
      End If
            
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   If HRRAB Then
   
      'agregamos Capital Propio
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'CAPPROPIO' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
   
      If Rs.EOF = False Then

         Valor = Val(vFld(Rs("Valor")))
      
         If Valor >= 0 Then
            Buf = Left("Capital Propio Tributario", 10) & Sep & "98" & Sep & Valor
            Print #Fd, Buf
            r = r + 1
         End If
         
      End If
      
      Call CloseRs(Rs)
      
   End If
   
      
   Close Fd

   If HRRAB Then
      If gEmpresa.Ano >= 2020 Then
         Msg = "HR-RAD"
      Else
         Msg = "HR-RAB"
      End If
   Else
      Msg = "DJ1923"
   End If

   If r = 0 Then
      MsgBox1 "No existen datos para generar el archivo de " & Msg & "." & vbCrLf & vbCrLf & "Verifique si existen movimientos en sus cuentas de resultado y si la configuración de las cuentas ha sido realizada.", vbInformation
      fname = ""
   Else
      FPath = ReplaceStr(FPath, "C:\HR\LPContab\..\", "C:\HR\")
      MsgBox1 "Proceso de exportación a " & Msg & " finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & FPath, vbInformation + vbOKOnly
   End If
      
   Export_DJ1923 = 0
   
End Function


'Exporta para DJ1924. Genera dos archivos de acuerdo a formato especificado
Public Function Export_DJ1924(Tipo As String, fname As String, FName2 As String) As Long
   Dim FPath As String, FPathB As String, FPathC As String
   Dim LogPath As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String
   Dim i As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim BaseImponible14ter As Double, IDPCBaseImp As Double
   Dim TotRentaAtrib As Double, TotCredIDPC As Double
   Dim RentaAtribuida As Double, CreditoIDPC As Double
   Dim DblPjePart As Double
   Dim rB As Integer, Rc As Integer
   
   Sep = ";"
   
   On Error Resume Next
      
'   FPath = gExportPath & "\HRDJ\"
'   MkDir FPath
'   FPath = gExportPath & "\HRDJ\" & gEmpresa.Rut
'   MkDir FPath
   
   
   FPath = gHRPath & "\RUTS"
   MkDir FPath
      
   FPath = FPath & "\" & Right("00000000" & gEmpresa.Rut, 8)
   MkDir FPath
      
   FPath = FPath & "\ImpConta"
   MkDir FPath
   
   'Primero DJ1924 parte B

   If Tipo = "1924B" Then
      fname = "DJ1924B_" & Right(gEmpresa.Ano, 2) & ".csv"
      FPathB = FPath & "\" & fname
      
      Fd = FreeFile
      ERR.Clear
      
      Open FPathB For Output As #Fd
      If ERR Then
         MsgErr FPathB
         Export_DJ1924 = -ERR
         Exit Function
      End If
   
      On Error GoTo 0
   
      'seleccionamos los registros
      Q1 = "SELECT TipoBaseImp, IdItemBaseImp, Valor "
      Q1 = Q1 & " FROM BaseImponible14TER "
      Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY TipoBaseImp, IdItemBaseImp"
      Set Rs = OpenRs(DbMain, Q1)
      
      Buf = "Ingresos Percibidos" & Sep & "Ingreso Diferido Imputado en el Ejercicio" & Sep & "Ingresos Devengados" & Sep & "Participaciones e Intereses Percibidos" & Sep & "Otros Ingresos Percibidos o Devengados" & Sep & "Crédito sobre Activos Fijos Adquiridos y Pagados en el Ejercicio" & Sep & "Costo Directo de los Bienes o Servicios" & Sep & "Remuneraciones" & Sep & "Adquisición de Bienes del Activo Realizable y Fijo" & Sep & "Intereses Pagados" & Sep & "Pérdidas de Ejercicios Anteriores" & Sep & "Otros Gastos Deducidos de los Ingresos Brutos" & Sep & "Mayor Valor Enajenación Bienes del Activo Fisico No Depreciables, De Acuerdo a la LIR"
   
      Print #Fd, Buf
   
      Buf = ""
      rB = 0
      
      
      'imprimimos el archivo
      Do While Rs.EOF = False
               
         If vFld(Rs("IdItemBaseImp")) > 0 Then
   
            Buf = Buf & vFld(Rs("Valor")) & Sep
            
         End If
               
         Rs.MoveNext
         
      Loop
      
      Call CloseRs(Rs)
      
      If Buf <> "" Then
         Print #Fd, Buf
         rB = 1
      End If
         
      Close Fd
   
   'Ahora DJ1924 parte C
   
   ElseIf Tipo = "1924C" Then
      FName2 = "DJ1924C_" & Right(gEmpresa.Ano, 2) & ".csv"
      FPathC = FPath & "\" & FName2
      
      Fd = FreeFile
      ERR.Clear
      
      Open FPathC For Output As #Fd
      If ERR Then
         MsgErr FPathC
         Export_DJ1924 = -ERR
         Exit Function
      End If
   
      On Error GoTo 0
      
      BaseImponible14ter = 0
      IDPCBaseImp = 0
      
      'obtenemos el valor guardado en Base Imponible
      Q1 = "SELECT Valor FROM BaseImponible14Ter "
      Q1 = Q1 & " WHERE TipoBaseImp = " & BASEIMP_TOTALES & " AND IdItemBaseImp = 0"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         BaseImponible14ter = vFld(Rs("Valor"))
         IDPCBaseImp = vFld(Rs("Valor")) * gImpPrimCategoria
         IDPCBaseImp = vFmt(Format(IDPCBaseImp, NUMFMT))    'Para evitar problemas de redondeo
               
      End If
      Call CloseRs(Rs)
   
   
      'seleccionamos los registros
      Q1 = "SELECT RUT, PjePart "
      Q1 = Q1 & " FROM Socios "
      Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id
      Q1 = Q1 & " ORDER BY PjePart "
      Set Rs = OpenRs(DbMain, Q1)
      
      Buf = "RUT del Titular de la Renta" & Sep & "Renta Atribuida" & Sep & "Renta Efectivamente Percibida" & Sep & "Crédito IDPC con Derecho a Devolución" & Sep & "Crédito IDPC sin Derecho a Devolución" & Sep & "Crédito IPE" & Sep & "PPM Puesto a Disposición de sus Propietarios, Socios, Comuneros o Accionistas"
   
      Print #Fd, Buf
   
      Buf = ""
      TotRentaAtrib = 0
      TotCredIDPC = 0
      Rc = 0
      
      'imprimimos el archivo
      Do While Rs.EOF = False
      
         Buf = vFld(Rs("Rut")) & "-" & DV_Rut(vFld(Rs("Rut"))) & Sep
         
         DblPjePart = vFld(Rs("PjePart"))    'para evitar problemas de redondeo, trabajando con todos los decimales
         
         RentaAtribuida = vFmt(Format(DblPjePart / 100 * BaseImponible14ter, NUMFMT)) 'para evitar problemas de redondeo
         CreditoIDPC = vFmt(Format(DblPjePart / 100 * IDPCBaseImp, NUMFMT))
         
         TotRentaAtrib = TotRentaAtrib + RentaAtribuida
         TotCredIDPC = TotCredIDPC + CreditoIDPC
         
         Rs.MoveNext
         
         If Rs.EOF Then    'es el último
            If TotRentaAtrib > BaseImponible14ter Then
               RentaAtribuida = RentaAtribuida - (TotRentaAtrib - BaseImponible14ter)
            ElseIf TotRentaAtrib < BaseImponible14ter Then
               RentaAtribuida = RentaAtribuida + (BaseImponible14ter - TotRentaAtrib)
            End If
         
            If TotCredIDPC > IDPCBaseImp Then
               CreditoIDPC = CreditoIDPC - (TotCredIDPC - IDPCBaseImp)
            ElseIf TotCredIDPC < IDPCBaseImp Then
               CreditoIDPC = CreditoIDPC + (IDPCBaseImp - TotCredIDPC)
            End If
         
         
         End If
            
         If RentaAtribuida < 0 Then   'Solicitado por Katherine enero 2020
            RentaAtribuida = 0
         End If
         
         If CreditoIDPC < 0 Then       'Solicitado por Katherine enero 2020
            CreditoIDPC = 0
         End If
         
         Buf = Buf & RentaAtribuida & Sep & 0 & Sep & CreditoIDPC & Sep & 0 & Sep & 0 & Sep & 0
         Rc = Rc + 1
         
         Print #Fd, Buf
         
      Loop
      
      Call CloseRs(Rs)
       
      Close Fd

   End If
   
'   If Rc = 0 Then
'      Kill (FPathC)
'   End If
   
   
   If rB = 0 And Rc = 0 Then
      MsgBox1 "No existen datos para generar esta Declaración Jurada." & vbCrLf & vbCrLf & "Recuerde revisar la Composición Societaria de la Empresa y la Base Imponible de 14 TER A).", vbInformation
      fname = ""
      FName2 = ""
   Else
      FPathB = ReplaceStr(FPathB, "C:\HR\LPContab\..\", "C:\HR\")
      FPathC = ReplaceStr(FPathC, "C:\HR\LPContab\..\", "C:\HR\")
'      MsgBox1 "Proceso de exportación finalizado." & vbCrLf & vbCrLf & "Se han generado los archivos:" & vbCrLf & vbCrLf & FPathB & vbCrLf & vbCrLf & FPathC, vbInformation + vbOKOnly
      
      If Tipo = "1924B" Then
         MsgBox1 "Proceso de exportación finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & FPathB, vbInformation + vbOKOnly
      Else
         MsgBox1 "Proceso de exportación finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & FPathC, vbInformation + vbOKOnly
      End If
   End If

   Export_DJ1924 = 0
   
End Function

'Exporta para DJ1847. Genera dos archivos de acuerdo a formato especificado
Public Function Export_DJ1847(fname As String) As Long
   Dim FPath As String
   Dim LogPath As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String, BufSecion2 As String
   Dim i As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim DblPjePart As Double, Diff As Double
   Dim InvActivo As Double, InvPasivo As Double
   Dim CtasActivoSIIEliminadas As String, CtasPasivoSIIEliminadas As String
   Dim TipoPartida As String
   Dim r As Integer
   Dim CtasSIIenBlanco As String
   Dim Codigo As Integer
   
   '2826671
   Dim CtasTributarioNegativo As String
   Dim CtasTributarioNegativo2 As String
   'fin 2826671
   
   Sep = ";"
   
   On Error Resume Next
      
'   FPath = gExportPath & "\HRDJ\"
'   MkDir FPath
'   FPath = gExportPath & "\HRDJ\" & gEmpresa.Rut
'   MkDir FPath
'
'   FName = FPath & "\" & "DJ1847_" & gEmpresa.Ano & ".csv"
 
 
   FPath = gHRPath & "\RUTS"
   MkDir FPath
      
   FPath = FPath & "\" & Right("00000000" & gEmpresa.Rut, 8)
   MkDir FPath
      
   FPath = FPath & "\ImpConta"
   MkDir FPath

   fname = "DJ1847_" & Right(gEmpresa.Ano, 2) & ".csv"
   
   FPath = FPath & "\" & fname
 
   Fd = FreeFile
   ERR.Clear
   
   Open FPath For Output As #Fd
   If ERR Then
      MsgErr FPath
      Export_DJ1847 = -ERR
      Exit Function
   End If

   On Error GoTo 0

   Buf = "Sección" & Sep & "Entidad Supervisora" & Sep & "Año Ajuste IFRS" & Sep & "Folio Inicial" & Sep & "Folio Final" & Sep & "Ajuste RLI" & Sep & "Cuenta Contable" & Sep & "Código ID Partida" & Sep & "Nombre de la cuenta" & Sep & "Debitos " & Sep & "Créditos " & Sep & "Saldo Deudor " & Sep & "Saldo Acreedor " & Sep & "Activo" & Sep & "Pasivo" & Sep & "Pérdidas" & Sep & "Ganancias" & Sep & "Tipo de Partida" & Sep & "Valor Tributario"
   
'   Print #Fd, Buf

   Buf = ""
   
   
   'seleccionamos los registros para la Sección 1
   Q1 = "SELECT IdEntSupervisora, AnoAjusteIFRS, FolioInicial, FolioFinal, IdAjustesRLI"
   Q1 = Q1 & " FROM InfoAnualDJ1847 "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   
   'imprimimos el primer registro del archivo (Sección 1)
   Buf = "1" & Sep
   
   If Not Rs.EOF Then
            
      Buf = Buf & gDJ1847_EntSupervisora(vFld(Rs("IdEntSupervisora"))) & Sep    'Entidad Supervisora
      Buf = Buf & vFld(Rs("AnoAjusteIFRS")) & Sep
      Buf = Buf & vFld(Rs("FolioInicial")) & Sep
      Buf = Buf & vFld(Rs("FolioFinal")) & Sep
      Buf = Buf & vFld(Rs("IdAjustesRLI"))
            
   Else
      MsgBox1 "No se han ingresado los antecedentes para la Sección 1 de la DJ 1847", vbExclamation
      Export_DJ1847 = -1
      Exit Function
            
   End If

   Call CloseRs(Rs)
   
   Print #Fd, Buf
      
   r = 0
   
   'Ahora continuamos con los registros de la sección 2
   
   'BufSecion2 = "2,,,,,,"
   BufSecion2 = "2;;;;;;"
   
   CtasActivoSIIEliminadas = ",1.03.05.00,"
   CtasPasivoSIIEliminadas = ",2.03.01.00,2.03.02.00,2.03.03.00,2.03.04.00,2.03.05.00,2.03.06.00,2.03.07.00,2.03.08.00,2.03.09.00,2.03.20.00,2.03.21.00,2.03.30.00,2.03.31.00,2.03.99.00,"
   CtasSIIenBlanco = ",1.03.05.00,2.03.01.00,2.03.02.00,2.03.03.00,2.03.04.00,2.03.05.00,2.03.06.00,2.03.07.00,2.03.08.00,2.03.09.00,2.03.20.00,2.03.21.00,2.03.30.00,2.03.31.00,2.03.99.00,"
   
   '2826671
   CtasTributarioNegativo = ",2-01,2-02"
   CtasTributarioNegativo2 = ",1-02-92,1-02-90,1-02-95,"
   'fin 2826671
   
   Q1 = "SELECT Cuentas.Codigo, Cuentas.Descripcion, Cuentas.Clasificacion, Cuentas.TipoPartida, PlanCuentasSII.FmtCodigoSII, Sum(MovComprobante.Debe) as Debe, Sum(MovComprobante.Haber) as Haber"
   Q1 = Q1 & " FROM ((Cuentas INNER JOIN MovComprobante ON Cuentas.idCuenta = MovComprobante.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante") & " )"
   Q1 = Q1 & " LEFT JOIN PlanCuentasSII ON Cuentas.CodCtaPlanSII = PlanCuentasSII.CodigoSII"
   Q1 = Q1 & " WHERE Comprobante.Fecha BETWEEN " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " AND " & CLng(DateSerial(gEmpresa.Ano, 12, 31))
   Q1 = Q1 & " AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & "))"
   Q1 = Q1 & " AND  Comprobante.Estado= " & EC_APROBADO
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY Cuentas.Codigo, Cuentas.Descripcion, Cuentas.Clasificacion, Cuentas.TipoPartida, PlanCuentasSII.FmtCodigoSII ORDER BY Cuentas.Codigo"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   
   Do While Not Rs.EOF
         
      Buf = BufSecion2 & FmtCodCuenta(vFld(Rs("Codigo"))) & Sep      'Cód. Cuenta LPContab
      Buf = Buf & vFld(Rs("FmtCodigoSII")) & Sep              'Cód. Cuenta SII
      'Buf = Buf & """" & vFld(Rs("Descripcion")) & """" & Sep
      Buf = Buf & vFld(Rs("Descripcion")) & Sep ' se quitan las comillas ADO 2737952 27-01-2022 FPG
      Buf = Buf & vFld(Rs("Debe")) & Sep        'Débitos
      Buf = Buf & vFld(Rs("Haber")) & Sep       'Créditos
      
      InvActivo = 0
      InvPasivo = 0
      
      Codigo = IIf(gEmpresa.Ano <= 2019, vFld(Rs("TipoPartida")), HomologaCod14D(gTipoPartida(vFld(Rs("TipoPartida"))).Codigo)) 'Esto es por incosistencia de HR que de un año para otro cambió el tipo de código que recibe
            
'      TipoPartida = IIf(gTipoPartida(vFld(Rs("TipoPartida"))).Codigo > 0, gTipoPartida(vFld(Rs("TipoPartida"))).Codigo, " ")
      TipoPartida = IIf(Codigo > 0, Codigo, " ")
      
      Diff = vFld(Rs("Debe")) - vFld(Rs("Haber"))
      
      If Diff > 0 Then
         Buf = Buf & Abs(Diff) & Sep & " " & Sep         'Saldo Acreedor
      Else
         Buf = Buf & " " & Sep & Abs(Diff) & Sep         'Saldo Deudor
      End If
      
      Select Case vFld(Rs("Clasificacion"))
         
         Case CLASCTA_ACTIVO, CLASCTA_PASIVO
         
            If Diff > 0 Then
               Buf = Buf & Abs(Diff) & Sep & " " & Sep & " " & Sep & " " & Sep   'Inv. Activo, Inv. Pasivo, Pérdodas y Ganancias
               InvActivo = Abs(Diff)
            Else
               Buf = Buf & " " & Sep & Abs(Diff) & Sep & " " & Sep & " " & Sep   'Inv. Activo, Inv. Pasivo, Pérdodas y Ganancias
               InvPasivo = Abs(Diff)
            End If
            
            Buf = Buf & TipoPartida & Sep
            
            If vFld(Rs("Clasificacion")) = CLASCTA_ACTIVO Then
               If InStr(CtasActivoSIIEliminadas, "," & vFld(Rs("FmtCodigoSII")) & ",") = 0 And InStr(CtasSIIenBlanco, "," & vFld(Rs("FmtCodigoSII")) & ",") = 0 Then
                  
                  '2826671
                  If Diff < 0 Then
                   Buf = Buf & InvPasivo * -1
                  Else
                  
                    If InStr(CtasTributarioNegativo, "," & Left(FmtCodCuenta(vFld(Rs("Codigo"))), 4) & ",") = 0 And InStr(CtasTributarioNegativo2, "," & Left(FmtCodCuenta(vFld(Rs("Codigo"))), 7) & ",") = 0 Then
                  
                    Buf = Buf & InvActivo         'Valor Trinutario
                    
                    Else
                    
                     Buf = Buf & InvActivo * -1       'Valor Trinutario
                    End If
                  End If
                  
                  'fin 2826671
               Else
                  Buf = Buf & " "
               End If
                           
            ElseIf vFld(Rs("Clasificacion")) = CLASCTA_PASIVO Then
               If InStr(CtasPasivoSIIEliminadas, "," & vFld(Rs("FmtCodigoSII")) & ",") = 0 And InStr(CtasSIIenBlanco, "," & vFld(Rs("FmtCodigoSII")) & ",") = 0 Then
                  Buf = Buf & InvPasivo * -1       'Valor Tributario
               Else
                  Buf = Buf & " "
               End If
            End If
            
            
         Case CLASCTA_RESULTADO
      
            If Diff > 0 Then
               Buf = Buf & " " & Sep & " " & Sep & Abs(Diff) & Sep & " " & Sep    'Inv. Activo, Inv. Pasivo, Pérdodas y Ganancias
            Else
               Buf = Buf & " " & Sep & " " & Sep & " " & Sep & Abs(Diff) & Sep    'Inv. Activo, Inv. Pasivo, Pérdodas y Ganancias
            End If
     
            Buf = Buf & TipoPartida & Sep & " "              'Tipo Partida y Valor Tributario
      
      End Select

      Print #Fd, Buf
      
      r = r + 1
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
    
   Close Fd

   If r = 0 Then
      MsgBox1 "No existen datos para generar esta Declaración Jurada.", vbInformation
      fname = ""
   Else
      FPath = ReplaceStr(FPath, "C:\HR\LPContab\..\", "C:\HR\")
      MsgBox1 "Proceso de exportación finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivos:" & vbCrLf & vbCrLf & FPath, vbInformation + vbOKOnly
   End If
   
   Export_DJ1847 = 0
   
End Function

Public Sub ConectHRCertif(ByVal Tipo As String, ByVal CsvFile As String, Optional ByVal CsvFile2 As String = "")
   Dim fname As String, Rc As Long, Ano As Integer, TOut As Long, W As Integer, W2 As Integer
   Dim FName1 As String, FName2 As String
   
'   W2 = 0
   
   If CsvFile = "" Then
      Exit Sub
   End If
   
   Select Case Tipo
   
      Case "1923"
         W = 65
      
      Case "1924B"
         W = 59
      
      Case "1924C"
         W = 70
'
      Case "1847"
         W = 44
         
      Case "1879"
         W = 6
   
'      Case "HRRAB"
      
   End Select

   Ano = gEmpresa.Ano

   fname = gHRPath & "\Bin\Wizard" & ((Ano + 1) Mod 100) & "Ext.exe"
'   fname = lContPath & "\..\Bin\Wizard" & ((Ano + 1) Mod 100) & "Ext.exe"
   
   If ExistFile(fname) Then
      TOut = Val(GetIniString(gCfgFile, "WCert", "TOut", "30")) ' Segs

      FName1 = fname & " " & CsvFile & "|" & Right("00000000" & Trim(vFmtCID(FmtCID(gEmpresa.Rut))), 8) & "-" & DV_Rut(vFmtCID(FmtCID(gEmpresa.Rut))) & "|" & W & "|" & "C"
   
      Rc = ExecCmd(FName1, vbNormalFocus, TOut * 1000#) ' 30 segundos o INFINITE
      Call AddDebug("Cmd: [" & FName1 & "], TOut=" & TOut & "[s], Rc=" & 1000 + Rc)
      If Rc Then
         Rc = 1000 + Rc
         MsgBox1 "Error " & Rc & " en la comunicación con HR Certificados." & vbCrLf & fname, vbExclamation
         Exit Sub
      End If
   
'      If W2 <> 0 Then
'         FName2 = fname & " " & CsvFile2 & "|" & Right("00000000" & Trim(vFmtCID(FmtCID(gEmpresa.Rut))), 8) & "-" & DV_Rut(vFmtCID(FmtCID(gEmpresa.Rut))) & "|" & W2 & "|" & "C"
'
'         Rc = ExecCmd(FName2, vbNormalFocus, TOut * 1000#) ' 30 segundos o INFINITE
'         Call AddDebug("Cmd: [" & FName2 & "], TOut=" & TOut & "[s], Rc=" & 1000 + Rc)
'
'      End If
   Else
      MsgBox1 "No se encontró el programa importador de HR-Certificados en" & vbCrLf & fname, vbExclamation
      Exit Sub
   End If

End Sub

Public Function Export_RetirosDividendos(fname As String) As Long
   Dim FPath As String
   Dim LogPath As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String
   Dim i As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Valor As Double
   Dim r As Integer
   Dim ExpDir As String
   Dim TipoArchivo As String
   
   On Error Resume Next
      
   Sep = ";"
   
   If gEmpresa.TipoContrib = CONTRIB_SAABIERTA Or gEmpresa.TipoContrib = CONTRIB_SACERRADA Or gEmpresa.TipoContrib = CONTRIB_SPORACCION Then
      TipoArchivo = "Dividendos"
   Else
      TipoArchivo = "Retiros"
   End If

   ExpDir = gHRPath & "\RUTS"
   MkDir ExpDir
      
   ExpDir = ExpDir & "\" & Right("00000000" & gEmpresa.Rut, 8)
   MkDir ExpDir
      
   ExpDir = ExpDir & "\ImpConta"
   MkDir ExpDir

   fname = TipoArchivo & "_" & Right(gEmpresa.Ano, 2) & ".csv"
   
   FPath = ExpDir & "\" & fname
   
   Fd = FreeFile
   ERR.Clear
   
   Open FPath For Output As #Fd
   If ERR Then
      MsgErr FPath
      Export_RetirosDividendos = -ERR
      Exit Function
   End If

   On Error GoTo 0
   

   'seleccionamos los registros
   Q1 = "SELECT Comprobante.Fecha, Socios.RUT, MovComprobante.Debe, MovComprobante.Haber "
   Q1 = Q1 & " FROM (Socios INNER JOIN MovComprobante ON Socios.IdCuentaRetiros = MovComprobante.IdCuenta"
   Q1 = Q1 & JoinEmpAno(gDbType, "Socios", "MovComprobante") & " )"
   Q1 = Q1 & " INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
   Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
   Q1 = Q1 & " WHERE Comprobante.Fecha BETWEEN " & CLng(DateSerial(gEmpresa.Ano, 1, 1)) & " AND " & CLng(DateSerial(gEmpresa.Ano, 12, 31))
   Q1 = Q1 & " AND ( Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste = " & TAJUSTE_AMBOS & ")"
   Q1 = Q1 & " AND Comprobante.Estado = " & EC_APROBADO
   Q1 = Q1 & " AND Comprobante.Tipo = " & TC_EGRESO
   Q1 = Q1 & " AND Comprobante.IdEmpresa = " & gEmpresa.id & " AND Comprobante.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Fecha, RUT "
   Set Rs = OpenRs(DbMain, Q1)
   
   If TipoArchivo = "Retiros" Then
      Buf = "Fecha" & Sep & "Rut Socio" & Sep & "Concepto" & Sep & "Monto" & Sep & "ISFUT"
   Else
      Buf = "Rut Socio" & Sep & "Monto"
   End If

   Print #Fd, Buf

   Buf = ""
   r = 0
   
   'imprimimos el archivo
   Do While Rs.EOF = False
            
      Valor = Abs(vFld(Rs("Debe")) - vFld(Rs("Haber")))          'Egresos
         
'      If Valor < 0 Then
'         Valor = 0
'      End If

      If Valor > 0 Then
         If TipoArchivo = "Retiros" Then
            Buf = Format(vFld(Rs("Fecha")), "dd/mm/yyyy") & Sep & vFld(Rs("RUT")) & DV_Rut(vFld(Rs("RUT"))) & Sep & "R" & Sep & Valor & Sep & "N"
         Else
            Buf = vFld(Rs("RUT")) & DV_Rut(vFld(Rs("RUT"))) & Sep & Valor
         End If
         
         Print #Fd, Buf
         r = r + 1
      End If
            
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
      
   Close Fd

   If r = 0 Then
      MsgBox1 "No existen datos para generar archivo de " & TipoArchivo & "." & vbCrLf & vbCrLf & "Verifique si existen movimientos en sus cuentas de " & TipoArchivo & " y si la configuración de las cuentas ha sido realizada.", vbInformation
      fname = ""
   Else
      FPath = ReplaceStr(FPath, "C:\HR\LPContab\..\", "C:\HR\")
      MsgBox1 "Proceso de exportación de Retiros/Dividendos finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & FPath, vbInformation + vbOKOnly
   End If
      
   Export_RetirosDividendos = 0

End Function

Public Function Export_HRRAD_BaseImp14D(fname As String) As Long
   Dim FPath As String
   Dim LogPath As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String
   Dim i As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Valor As Double
   Dim r As Integer
   Dim ExpDir As String
   Dim TipoArchivo As String
   Dim Fecha As Long
   Dim Descrip As String
   
   On Error Resume Next
      
   Sep = ";"
   
   'Exportación HR-RAD BAseImponible 14D
   TipoArchivo = "HR-RAD-BIMP14D"

   ExpDir = gHRPath & "\RUTS"
   MkDir ExpDir
      
   ExpDir = ExpDir & "\" & Right("00000000" & gEmpresa.Rut, 8)
   MkDir ExpDir
      
   ExpDir = ExpDir & "\ImpConta"
   MkDir ExpDir

   fname = TipoArchivo & "_" & Right(gEmpresa.Ano, 2) & ".csv"
   
   FPath = ExpDir & "\" & fname
   
   Fd = FreeFile
   ERR.Clear
   
   Open FPath For Output As #Fd
   If ERR Then
      MsgErr FPath
      Export_HRRAD_BaseImp14D = -ERR
      Exit Function
   End If

   On Error GoTo 0
   

   'seleccionamos los registros
  ' Q1 = "SELECT Codigo, Valor FROM BaseImponible14D "
  'Q1 = Q1 & " WHERE Valor <> 0 AND Nivel = " & BIMP14D_MAXNIV
  ' Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
  ' Q1 = Q1 & " ORDER BY Codigo "
   
   
   'pipe tema 2 2738156 se realiza cambio de select para sumar "Otros gasto 11700" a  "Otras Deducciones a la RLI 10400"
    Q1 = "SELECT 10400 as Codigo,sum(valor) as valo from  BaseImponible14D  "
    Q1 = Q1 & " WHERE Valor <> 0 AND Nivel = " & BIMP14D_MAXNIV
    Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
    Q1 = Q1 & " AND CODIGO in (10410,10400) "
    Q1 = Q1 & " union "
    Q1 = Q1 & " SELECT Codigo, Valor as valo FROM BaseImponible14D "
    Q1 = Q1 & " WHERE Valor <> 0 AND Nivel = " & BIMP14D_MAXNIV
    Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
    If gEmpresa.Ano >= 2021 Then
'        '*** ADO 2747807 Tema 1 El codigo 5300 (Servicios Pagados) no debe aparecer en la exportacion a partir del 2021
        If gEmpresa.Ano > 2021 And gEmpresa.ProPymeTransp Then
            Q1 = Q1 & " AND CODIGO NOT in (10410,10400,5300,1700,2800,4700) "
        Else
            Q1 = Q1 & " AND CODIGO NOT in (10410,10400,5300) "
        End If
    Else
        Q1 = Q1 & " AND CODIGO NOT in (10410,10400,401) "
    End If
    Q1 = Q1 & " ORDER BY Codigo "
      
   Set Rs = OpenRs(DbMain, Q1)
   
   Buf = "Codigo" & Sep & "Fecha" & Sep & "Descripción" & Sep & "Valor Nominal"

   Print #Fd, Buf

   Buf = ""
   r = 0
   Fecha = DateSerial(gEmpresa.Ano, 12, 31)
   Descrip = ""
   
   i = 1
   
   'imprimimos el archivo
   Do While Rs.EOF = False
            
      Do While i <= UBound(gBaseImponible14D)
      
         If gBaseImponible14D(i).Codigo = vFld(Rs("Codigo")) Then
            If gBaseImponible14D(i).Codigo = 401 Then
                If FProPymeGeneral Then
                    Descrip = "Ingresos Devengados Ejerc Anter"
                Else
                    Descrip = "Ingresos del Giro Devengados.."
                End If
                Exit Do
            End If
            If gBaseImponible14D(i).Codigo = 5201 Then
                Descrip = "Existencias, Insumos del Ejercicio Anterior"
                Exit Do
            End If
            Descrip = "LPConta " & gBaseImponible14D(i).Nombre
            Exit Do
         Else
            i = i + 1
         End If
         
      Loop
            
      Valor = Abs(vFld(Rs("Valo")))
         
      '2975665
      If Valor <> 0 Then
      Buf = vFld(Rs("Codigo")) & Sep & Format(Fecha, "dd/mm/yyyy") & Sep & Descrip & Sep & Valor
      Print #Fd, Buf
      End If
     
      'Buf = vFld(Rs("Codigo")) & Sep & Format(Fecha, "dd/mm/yyyy") & Sep & Descrip & Sep & Valor
       '2975665
       
      'Print #Fd, Buf
      r = r + 1
            
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
      
   Close Fd

   If r = 0 Then
      MsgBox1 "No existen datos en la Base Imponible 14D para generar archivo para HR-RAD." & vbCrLf & vbCrLf & "Ingrese a la Base Imponible 14 D y verifique la información que se presenta.", vbInformation
      fname = ""
   Else
      FPath = ReplaceStr(FPath, "C:\HR\LPContab\..\", "C:\HR\")
      MsgBox1 "Proceso de exportación Base Imponible 14D para HR-RAD finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & FPath, vbInformation + vbOKOnly
   End If
      
   Export_HRRAD_BaseImp14D = 0

End Function

Public Function Export_Percepciones(fname As String) As Long
   Dim FPath As String
   Dim LogPath As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String
   Dim BufAux As String
   Dim i As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Valor As Double
   Dim IdPerc As Long
   Dim r As Integer
   Dim ExpDir As String
   Dim TipoArchivo As String
   Dim Fecha As Long
   Dim Descrip As String
   
   On Error Resume Next
      
   Sep = ";"
   
   'Exportación HR-RAD BAseImponible 14D
   TipoArchivo = "PERCEPCIONES"

   ExpDir = gHRPath & "\RUTS"
   MkDir ExpDir
      
   ExpDir = ExpDir & "\" & Right("00000000" & gEmpresa.Rut, 8)
   MkDir ExpDir
      
   ExpDir = ExpDir & "\ImpConta"
   MkDir ExpDir

   fname = TipoArchivo & "_" & Right(gEmpresa.Ano, 2) & ".csv"
   
   FPath = ExpDir & "\" & fname
   
   Fd = FreeFile
   ERR.Clear
   
   Open FPath For Output As #Fd
   If ERR Then
      MsgErr FPath
      Export_Percepciones = -ERR
      Exit Function
   End If

   On Error GoTo 0
   

   'seleccionamos los registros
   Q1 = "SELECT fecha, numcertificado, rutempresa, regimen, contabilizacion, tasatef, tasatex, d.idperc, coddet, valor "
   Q1 = Q1 & " FROM percepciones as p,  detpercepciones as d "
   Q1 = Q1 & " Where p.idperc = d.idperc "
   Q1 = Q1 & " and ano = " & gEmpresa.Ano
   Q1 = Q1 & " and coddet > 0 "
   Q1 = Q1 & " and coddet NOT IN (300,1800,3100,4400) "
   Q1 = Q1 & " ORDER BY d.idperc, coddet"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Buf = "Fecha" & Sep & "N° Certificado" & Sep & "Rut Empresa" & Sep & "Regimen" & Sep & "Contabilización" & Sep & "Tasa Tef" & Sep & "Tasa Tex" & Sep
   For i = 1 To UBound(Percepciones)
'        If Percepciones(i).Codigo <> 300 And Percepciones(i).Codigo <> 1800 And Percepciones(i).Codigo <> 3100 And Percepciones(i).Codigo <> 4400 And Percepciones(i).Nivel = 5 Then
'          Buf = Buf & Percepciones(i).Nombre & Sep '& "Fecha" & Sep & "Descripción" & Sep & "Valor Nominal"
'        End If
        If Percepciones(i).Nivel = 5 Then
          Buf = Buf & Percepciones(i).Nombre & Sep '& "Fecha" & Sep & "Descripción" & Sep & "Valor Nominal"
        End If
   Next i

   
   Print #Fd, Buf

   Buf = ""
   r = 0
'   Fecha = DateSerial(gEmpresa.Ano, 12, 31)
'   Descrip = ""
'
'   i = 1
'
   'imprimimos el archivo
   IdPerc = 0
   Do While Rs.EOF = False

      If IdPerc = 0 Then
       IdPerc = vFld(Rs("idperc"))
      End If

      'Valor = Abs(vFld(Rs("Valo")))
      If IdPerc = vFld(Rs("idperc")) Then
          BufAux = Format(vFld(Rs("fecha")), "dd/mm/yyyy") & Sep & vFld(Rs("numcertificado")) & Sep & Replace(FmtCID(vFld(Rs("rutempresa"))), ".", "") & Sep & vFld(Rs("regimen")) & Sep & vFld(Rs("contabilizacion")) & Sep & vFld(Rs("tasatef")) & Sep & vFld(Rs("tasatex")) & Sep
          If NivelPercepciones(vFld(Rs("coddet"))) = 5 Then
               Buf = Buf & vFld(Rs("valor")) & Sep '& Format(Fecha, "dd/mm/yyyy") & Sep & Descrip & Sep & Valor
          End If
      Else
         Buf = BufAux & Buf
         Print #Fd, Buf
         Buf = ""
         If NivelPercepciones(vFld(Rs("coddet"))) = 5 Then
            Buf = Buf & vFld(Rs("valor")) & Sep
         End If
         IdPerc = vFld(Rs("idperc"))
    
      End If
      
      r = r + 1

      Rs.MoveNext

   Loop
   Buf = BufAux & Buf
   Print #Fd, Buf
   
   Call CloseRs(Rs)
      
   Close Fd

   If r = 0 Then
      MsgBox1 "No existen datos en Percepciones para generar archivo para HR-RAD." & vbCrLf & vbCrLf & "Ingrese a la Base Imponible 14 D y verifique la información que se presenta.", vbInformation
      fname = ""
   Else
      'FPath = ReplaceStr(FPath, "C:\HR\LPContab\..\", "C:\HR\")
      MsgBox1 "Proceso de exportación Percepciones para HR-RAD finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & FPath, vbInformation + vbOKOnly
   End If
      
   Export_Percepciones = 0

End Function
Private Function NivelPercepciones(Codigo As Integer) As Integer
Dim i As Integer
NivelPercepciones = 0
For i = 1 To UBound(Percepciones)
   If Percepciones(i).Codigo = Codigo Then
      NivelPercepciones = Percepciones(i).Nivel
   End If
Next i
End Function

Public Function Export_HRRAD_CPS_Totales(fname As String) As Long
   Dim FPath As String
   Dim LogPath As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String
   Dim i As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Valor As Double
   Dim r As Integer
   Dim ExpDir As String
   Dim TipoArchivo As String
   Dim Fecha As String
   Dim Metodo As String
   Dim TipoReg As String
   Dim Detalle As String
   
   On Error Resume Next
      
   Sep = ";"
   
   'Exportación HR-RAD Capital Propio Simplificado
   TipoArchivo = "HR-RAD-CPS"

   ExpDir = gHRPath & "\RUTS"
   MkDir ExpDir
      
   ExpDir = ExpDir & "\" & Right("00000000" & gEmpresa.Rut, 8)
   MkDir ExpDir
      
   ExpDir = ExpDir & "\ImpConta"
   MkDir ExpDir

   fname = TipoArchivo & "_" & Right(gEmpresa.Ano, 2) & ".csv"
   
   FPath = ExpDir & "\" & fname
   
   Fd = FreeFile
   ERR.Clear
   
   Open FPath For Output As #Fd
   If ERR Then
      MsgErr FPath
      Export_HRRAD_CPS_Totales = -ERR
      Exit Function
   End If

   On Error GoTo 0
     

   Buf = "Método" & Sep & "Ítem" & Sep & "Fecha" & Sep & "Detalle" & Sep & "Monto"

   Print #Fd, Buf
   
   Buf = ""
   r = 0
   Fecha = Format(DateSerial(gEmpresa.Ano, 12, 31), "dd/mm/yyyy")
   Metodo = ""
   Detalle = "LP Conta Traspaso"
   
   'seleccionamos los registros de Variación Anual
   Q1 = "SELECT CPS_CapitalAportado, CPS_AumentosCapital, CPS_Disminuciones, CPS_OtrosAjustesAumentos, CPS_OtrosAjustesDisminuciones "
   Q1 = Q1 & " FROM EmpresasAno "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      
   Set Rs = OpenRs(DbMain, Q1)
   TipoReg = "VA"
   
   'imprimimos el archivo
   If Not Rs.EOF Then
                     
      Metodo = "CAP"
      Valor = vFld(Rs("CPS_CapitalAportado"))
      If Valor > 0 Then
         Buf = TipoReg & Sep & Metodo & Sep & Fecha & Sep & Detalle & Sep & Valor
         
         Print #Fd, Buf
         r = r + 1
      End If
      
      Metodo = "AC"
      Valor = GetCPSAnual(CPS_AUMENTOSCAP)
      If Valor > 0 Then
         Buf = TipoReg & Sep & Metodo & Sep & Fecha & Sep & Detalle & Sep & Valor
         
         Print #Fd, Buf
         r = r + 1
      End If
      
      Metodo = "DC"
      Valor = GetCPSAnual(CPS_DISMINUCIONES)
      If Valor > 0 Then
         Buf = TipoReg & Sep & Metodo & Sep & Fecha & Sep & Detalle & Sep & Valor
            
         Print #Fd, Buf
         r = r + 1
      End If
      
      Metodo = "OAA"
      Valor = GetCPSAnual(CPS_OTROSAJUSTAUMENTOS)
      If Valor > 0 Then
         Buf = TipoReg & Sep & Metodo & Sep & Fecha & Sep & Detalle & Sep & Valor
         
         Print #Fd, Buf
         r = r + 1
      End If
      
      Metodo = "OAD"
      Valor = GetCPSAnual(CPS_OTROSAJUSTDISMIN)
      If Valor > 0 Then
         Buf = TipoReg & Sep & Metodo & Sep & Fecha & Sep & Detalle & Sep & Valor
         
         Print #Fd, Buf
         r = r + 1
      End If
      
      r = r + 1
   End If
   
   Call CloseRs(Rs)
      
      
'*************** Comienzo cambio Ado 2699586 Tema 4.4 31-01-2022 FPG **************
      
'   'seleccionamos los registros de Acumolado Total
'   Q1 = "SELECT CPS_CapitalAportado, CPS_AumentosCapital, CPS_Disminuciones, CPS_OtrosAjustesAumentos, CPS_OtrosAjustesDisminuciones "
'   Q1 = Q1 & " FROM EmpresasAno "
'   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'
'   Set Rs = OpenRs(DbMain, Q1)
'
'   TipoReg = "AT"
'
'   'imprimimos el archivo
'   If Not Rs.EOF Then
'
'      Metodo = "CAP"
'      Valor = vFld(Rs("CPS_CapitalAportado"))
'      If Valor > 0 Then
'         Buf = TipoReg & Sep & Metodo & Sep & Fecha & Sep & Detalle & Sep & Valor
'
'         Print #Fd, Buf
'         r = r + 1
'      End If
'
'      Metodo = "AC"
'      Valor = vFld(Rs("CPS_AumentosCapital"))
'      If Valor > 0 Then
'         Buf = TipoReg & Sep & Metodo & Sep & Fecha & Sep & Detalle & Sep & Valor
'
'         Print #Fd, Buf
'         r = r + 1
'      End If
'
'      Metodo = "DC"
'      Valor = vFld(Rs("CPS_Disminuciones"))
'      If Valor > 0 Then
'         Buf = TipoReg & Sep & Metodo & Sep & Fecha & Sep & Detalle & Sep & Valor
'
'         Print #Fd, Buf
'         r = r + 1
'      End If
'
'      Metodo = "OAA"
'      Valor = vFld(Rs("CPS_OtrosAjustesAumentos"))
'      If Valor > 0 Then
'         Buf = TipoReg & Sep & Metodo & Sep & Fecha & Sep & Detalle & Sep & Valor
'
'         Print #Fd, Buf
'         r = r + 1
'      End If
'
'      Metodo = "OAD"
'      Valor = vFld(Rs("CPS_OtrosAjustesDisminuciones"))
'      If Valor > 0 Then
'         Buf = TipoReg & Sep & Metodo & Sep & Fecha & Sep & Detalle & Sep & Valor
'
'         Print #Fd, Buf
'         r = r + 1
'      End If
'
'   End If
'
'   Call CloseRs(Rs)
'*************** Fin cambio Ado 2699586 Tema 4.4 31-01-2022 FPG **************
   Close Fd

   If r = 0 Then
      MsgBox1 "No existen datos para generar archivo de Capital Propio Simplificado para HR-RAD." & vbCrLf & vbCrLf & "Ingrese al Capital Propio Simplificado General y verifique la información que se presenta.", vbInformation
      fname = ""
   Else
      FPath = ReplaceStr(FPath, "C:\HR\LPContab\..\", "C:\HR\")
      MsgBox1 "Proceso de exportación de Capital Propio Simplificado para HR-RAD finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & FPath, vbInformation + vbOKOnly
   End If
      
   Export_HRRAD_CPS_Totales = 0

End Function


Public Function Export_HRRAD_CPS_Detalle(fname As String) As Long
   Dim FPath As String
   Dim LogPath As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String
   Dim i As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Valor As Double
   Dim r As Integer
   Dim ExpDir As String
   Dim TipoArchivo As String
   Dim Fecha As String
   Dim Metodo As String
   Dim TipoReg As String
   Dim Detalle As String
   
   
   On Error Resume Next
      
   Sep = ";"
   
   'Exportación HR-RAD Capital Propio Simplificado
   TipoArchivo = "HR-RAD-CPS"

   ExpDir = gHRPath & "\RUTS"
   MkDir ExpDir
      
   ExpDir = ExpDir & "\" & Right("00000000" & gEmpresa.Rut, 8)
   MkDir ExpDir
      
   ExpDir = ExpDir & "\ImpConta"
   MkDir ExpDir

   fname = TipoArchivo & "_" & Right(gEmpresa.Ano, 2) & ".csv"
   
   FPath = ExpDir & "\" & fname
   
   Fd = FreeFile
   ERR.Clear
   
   Open FPath For Output As #Fd
   If ERR Then
      MsgErr FPath
      Export_HRRAD_CPS_Detalle = -ERR
      Exit Function
   End If

   On Error GoTo 0
     

   Buf = "Método" & Sep & "Ítem" & Sep & "Fecha" & Sep & "Detalle" & Sep & "Monto"

   Print #Fd, Buf
   
   Buf = ""
   r = 0
   Fecha = Format(DateSerial(gEmpresa.Ano, 12, 31), "dd/mm/yyyy")
   Metodo = ""
   
   'seleccionamos los registros de Variación Anual
   Q1 = "SELECT TipoDetCPS, Fecha, Cuentas.Descripcion as DescCuenta, Valor, Descrip "
   Q1 = Q1 & " FROM (DetCapPropioSimpl LEFT JOIN Cuentas ON DetCapPropioSimpl.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "DetCapPropioSimpl", "Cuentas") & ")"
   Q1 = Q1 & " WHERE DetCapPropioSimpl.IdEmpresa = " & gEmpresa.id & " AND DetCapPropioSimpl.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " AND TipoDetCPS IN (" & CPS_AUMENTOSCAP & ", " & CPS_DISMINUCIONES & ", " & CPS_OTROSAJUSTAUMENTOS & ", " & CPS_OTROSAJUSTDISMIN & ")"
   Q1 = Q1 & " ORDER BY Fecha, TipoDetCPS "
      
   Set Rs = OpenRs(DbMain, Q1)
   TipoReg = "VA"
   
   'imprimimos el archivo
   Do While Not Rs.EOF
                           
      Valor = vFld(Rs("Valor"))
      
      If Valor > 0 Then
      
         Select Case vFld(Rs("TipoDetCPS"))
            Case CPS_AUMENTOSCAP
               Metodo = "AC"
               
            Case CPS_DISMINUCIONES
               Metodo = "DC"
               
            Case CPS_OTROSAJUSTAUMENTOS
               Metodo = "OAA"
               
            Case CPS_OTROSAJUSTDISMIN
               Metodo = "OAD"
               
         End Select
               
         If vFld(Rs("Fecha")) > 0 Then
            Fecha = Format(vFld(Rs("Fecha")), "dd/mm/yyyy")
         Else
            Fecha = Format(DateSerial(gEmpresa.Ano, 12, 31), "dd/mm/yyyy")
         End If
         
         If vFld(Rs("DescCuenta")) <> "" Then
            Detalle = vFld(Rs("DescCuenta"))
         ElseIf vFld(Rs("Descrip")) <> "" Then
            Detalle = vFld(Rs("Descrip"))
         Else
            Detalle = "LP Conta Traspaso"
         End If

         
         Buf = TipoReg & Sep & Metodo & Sep & Fecha & Sep & Detalle & Sep & Valor
         
         Print #Fd, Buf
         r = r + 1
         
      End If
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
            
   Close Fd

   If r = 0 Then
      MsgBox1 "No existen datos para generar archivo de Capital Propio Simplificado para HR-RAD." & vbCrLf & vbCrLf & "Ingrese al Capital Propio Simplificado General y verifique la información que se presenta.", vbInformation
      fname = ""
   Else
      FPath = ReplaceStr(FPath, "C:\HR\LPContab\..\", "C:\HR\")
      MsgBox1 "Proceso de exportación de Capital Propio Simplificado para HR-RAD finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & FPath, vbInformation + vbOKOnly
   End If
      
   Export_HRRAD_CPS_Detalle = 0

End Function

'2699582
Public Function Export_PPM_BaseImp14D(fname As String) As Long
   Dim FPath As String
   Dim LogPath As String
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String
   Dim Buf2 As String
   Dim i As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Valor As Double
   Dim r As Integer
   Dim ExpDir As String
   Dim TipoArchivo As String
   Dim Total As Integer
   Dim FechaPPM As Long
   Dim TipoPPM As Boolean

   
   On Error Resume Next
      
   Sep = ";"
   
   'Exportación HR-RAD BAseImponible 14D
   TipoArchivo = "HR-RAD-REAJPPM"

   ExpDir = gHRPath & "\RUTS"
   MkDir ExpDir
      
   ExpDir = ExpDir & "\" & Right("00000000" & gEmpresa.Rut, 8)
   MkDir ExpDir
      
   ExpDir = ExpDir & "\ImpConta"
   MkDir ExpDir

   fname = TipoArchivo & "_" & Right(gEmpresa.Ano, 2) & ".csv"
   
   FPath = ExpDir & "\" & fname
   
   Fd = FreeFile
   ERR.Clear
   
   Open FPath For Output As #Fd
   If ERR Then
      MsgErr FPath
      Export_PPM_BaseImp14D = -ERR
      Exit Function
   End If

   On Error GoTo 0
   
    Buf = "Fecha PPMO" & Sep & "Monto PPMO" & Sep & "Fecha PPMV" & Sep & "Monto PPMV"

   Print #Fd, Buf
   
  
   Q1 = "SELECT Codigo, Valor FROM ParamEmpresa WHERE Tipo='PPM'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      TipoPPM = IIf(vFld(Rs("Valor")) <> 0, True, False)
      FechaPPM = DateSerial(gEmpresa.Ano, 1, 20)
   End If
   
   Call CloseRs(Rs)
  

  If gDbType = SQL_ACCESS Then
   Q1 = "SELECT distinct Comprobante.FECHA ,iif( (MovComprobante.DEBE * (select distinct factor from FactorActAnual where ano "
   Q1 = Q1 & " =year(Comprobante.FECHA)  and MesCol = 12 and MesRow =month(Comprobante.FECHA) )) is null "
   Q1 = Q1 & " ,MovComprobante.DEBE*1,(MovComprobante.DEBE * (select distinct factor from FactorActAnual where ano =year(Comprobante.FECHA)  and "
   Q1 = Q1 & " MesCol = 12 and MesRow =month(Comprobante.FECHA) )))  as Monto, iif(ParamEmpresa.tipo = 'CTAPPMOBLI','O','V' ) AS TIPO "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp) "
   Q1 = Q1 & " INNER JOIN ParamEmpresa ON MovComprobante.IdEmpresa = ParamEmpresa.IdEmpresa,FactorActAnual "
   Q1 = Q1 & " WHERE ParamEmpresa.tipo in ('CTAPPMOBLI','CTAPPMVOLU') AND  MovComprobante.idCuenta = int(ParamEmpresa.valor) "
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND COMPROBANTE.TIPO = " & TC_EGRESO
   Q1 = Q1 & " AND COMPROBANTE.ESTADO = " & EC_APROBADO
   Q1 = Q1 & " ORDER BY 3 ASC "
  ElseIf gDbType = SQL_SERVER Then
   Q1 = "SELECT distinct Comprobante.FECHA ,iif( (MovComprobante.DEBE * (select distinct factor from FactorActAnual where ano "
   Q1 = Q1 & " =year(Comprobante.FECHA)  and MesCol = 12 and MesRow =month(Comprobante.FECHA) )) is null "
   Q1 = Q1 & " ,MovComprobante.DEBE*1,(MovComprobante.DEBE * (select distinct factor from FactorActAnual where ano =year(Comprobante.FECHA)  and "
   Q1 = Q1 & " MesCol = 12 and MesRow =month(Comprobante.FECHA) )))  as Monto, iif(ParamEmpresa.tipo = 'CTAPPMOBLI','O','V' ) AS TIPO "
   Q1 = Q1 & " FROM (MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp) "
   Q1 = Q1 & " INNER JOIN ParamEmpresa ON MovComprobante.IdEmpresa = ParamEmpresa.IdEmpresa,FactorActAnual "
   Q1 = Q1 & " WHERE ParamEmpresa.tipo in ('CTAPPMOBLI','CTAPPMVOLU') AND  MovComprobante.idCuenta = Convert(int,ParamEmpresa.valor) "
   Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " AND COMPROBANTE.TIPO = " & TC_EGRESO
   Q1 = Q1 & " AND COMPROBANTE.ESTADO = " & EC_APROBADO
   Q1 = Q1 & " ORDER BY 3 ASC "
  
  End If

   Set Rs = OpenRs(DbMain, Q1)
     

   Buf = ""
   r = 0
      
   i = 0
   
   
   Dim arrO(1000) As Variant
   Dim arrV(1000) As Variant
   
  
   'imprimimos el archivo
   Do While Rs.EOF = False
            
           
                     
      If vFld(Rs("TIPO")) = "O" Then
        If TipoPPM And vFld(Rs("Fecha")) <= FechaPPM Then
        Else
          Buf = Format(vFld(Rs("Fecha")), "dd/mm/yyyy") & Sep & vFld(Rs("Monto"))
          arrO(i) = Buf
          i = i + 1
        End If
      ElseIf vFld(Rs("TIPO")) = "V" Then
      Buf2 = Format(vFld(Rs("Fecha")), "dd/mm/yyyy") & Sep & vFld(Rs("Monto"))
      arrV(r) = Buf2
      r = r + 1
      End If

      Rs.MoveNext
      
   Loop
   
   If i >= r Then
   Total = i
   ElseIf r >= i Then
   Total = r
   End If
   
   
   For i = 0 To Total
     
     If arrO(i) = Empty And Len(arrV(i)) > 0 Then
     Print #Fd, arrO(i); ";"; ";"; arrV(i)
     ElseIf arrV(i) = Empty And Len(arrO(i)) > 0 Then
     
     Print #Fd, arrO(i); ";"; ";"; arrV(i)
     
     Else
     Print #Fd, arrO(i); ";"; arrV(i)
     End If
     
    'Print #Fd, arrO(i); ";"; arrV(i)
    
   Next i
   Call CloseRs(Rs)
      
   Close Fd

   If Total = 0 Then
      MsgBox1 "No existen datos en la Base Imponible 14D para generar archivo para HR-RAD." & vbCrLf & vbCrLf & "Ingrese a la Base Imponible 14 D y verifique la información que se presenta.", vbInformation
      fname = ""
   Else
      FPath = ReplaceStr(FPath, "C:\HR\LPContab\..\", "C:\HR\")
      MsgBox1 "Proceso de exportación Base Imponible 14D para HR-RAD PPM finalizado." & vbCrLf & vbCrLf & "Se ha generado el archivo:" & vbCrLf & vbCrLf & FPath, vbInformation + vbOKOnly
   End If
      
   Export_PPM_BaseImp14D = 0

End Function





