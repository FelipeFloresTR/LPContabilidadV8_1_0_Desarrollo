Attribute VB_Name = "HyperContFca"
Option Explicit

'clases de impresión
Public gPrtLibros As ClsPrtFlxGrid
Public gPrtReportes As ClsPrtFlxGrid

Public gPrtCheques As ClsPrtCheque

'indentación de reportes con niveles
Global Const REP_INDENT = 3

'colores para reportes por nivel
Public gRepNivColor(MAX_NIVELES) As Long

'estructura para reportes por nivel
Type RepNiv_t
   Debe As Double
   Haber As Double
   InvDebe As Double
   InvHaber As Double
   ResDebe As Double
   ResHaber As Double
   Linea As Integer
End Type


Public Sub IniHyperContFca()

   gTipoComp(TC_INGRESO) = "Ingreso"
   gTipoComp(TC_EGRESO) = "Egreso"
   gTipoComp(TC_TRASPASO) = "Traspaso"
   gTipoComp(TC_APERTURA) = "Apertura"
   
   gEstadoComp(EC_PENDIENTE) = "Pendiente"
   gEstadoComp(EC_APROBADO) = "Aprobado"
   gEstadoComp(EC_ANULADO) = "Anulado"
   gEstadoComp(EC_ERRONEO) = "Erróneo"
   
   gRepNivColor(1) = &H400000
   gRepNivColor(2) = &H800000
   gRepNivColor(3) = &HC00000
   gRepNivColor(4) = &HFF0000
   gRepNivColor(5) = &HFF8080
   
End Sub
Public Sub CreatePrtFormats()
   Dim Nombres(7) As String
   Dim FntNombres(7) As FontDef_t
   Dim FntTitulos(0) As FontDef_t
   Dim FntEncabezados(0) As FontDef_t
   Dim Frm As FrmPrintPreview
   Dim i As Integer
   
   Set gPrtLibros = New ClsPrtFlxGrid
   
   gPrtLibros.PrtDemo = gAppCode.Demo
         
   FntNombres(0).FontName = "Arial"
   FntNombres(0).FontSize = 9
   FntNombres(0).FontBold = False
   
   Call gPrtLibros.FntNombres(FntNombres)
   
   FntTitulos(0).FontName = "Arial"
   FntTitulos(0).FontSize = 14
      
   Call gPrtLibros.FntTitulos(FntTitulos)
      
   FntEncabezados(0).FontName = "Arial"
   FntEncabezados(0).FontSize = 10
   
   Call gPrtLibros.FntEncabezados(FntEncabezados)
   
   
   Set gPrtReportes = New ClsPrtFlxGrid
      
   gPrtReportes.PrtDemo = gAppCode.Demo
      
   FntNombres(0).FontName = "Arial"
   FntNombres(0).FontSize = 10
   FntNombres(0).FontBold = True
   
   For i = 1 To UBound(FntNombres)
      FntNombres(i).FontName = "Arial"
      FntNombres(i).FontSize = 9
      FntNombres(i).FontBold = False
   Next i
   
   Call gPrtReportes.FntNombres(FntNombres)
   
   FntTitulos(0).FontName = "Arial"
   FntTitulos(0).FontSize = 14
      
   Call gPrtReportes.FntTitulos(FntTitulos)
      
   FntEncabezados(0).FontName = "Arial"
   FntEncabezados(0).FontSize = 10
   
   Call gPrtReportes.FntEncabezados(FntEncabezados)
   
   Call SetPrtNotas

   Set gPrtCheques = New ClsPrtCheque
   
   gPrtCheques.PrtDemo = gAppCode.Demo
   Call gPrtCheques.FntNombres(FntNombres)
   Call gPrtCheques.FntTitulos(FntTitulos)
   Call gPrtCheques.FntEncabezados(FntEncabezados)
   
   gPrtCheques.AlturaCheque = Val(GetIniString(gIniFile, "Cheques", "Altura", gPrtCheques.AlturaCheque))
   gPrtCheques.BordeIzqCheque = Val(GetIniString(gIniFile, "Cheques", "BordeIzq", gPrtCheques.BordeIzqCheque))
   gPrtCheques.BajarValDig = Val(GetIniString(gIniFile, "Cheques", "BajarValDig", ""))
   gPrtCheques.MoverValDig = Val(GetIniString(gIniFile, "Cheques", "MoverValDig", ""))
   gPrtCheques.BajarFecha = Val(GetIniString(gIniFile, "Cheques", "BajarFecha", ""))
   gPrtCheques.MoverFecha = Val(GetIniString(gIniFile, "Cheques", "MoverFecha", ""))
   gPrtCheques.BajarOrdenDe = Val(GetIniString(gIniFile, "Cheques", "BajarOrdenDe", ""))
   gPrtCheques.MoverOrdenDe = Val(GetIniString(gIniFile, "Cheques", "MoverOrdenDe", ""))
        
   gPrtCheques.BorrarALaOrden = Val(GetIniString(gIniFile, "Cheques", "BorrarOrden", ""))
   gPrtCheques.BorrarAlPortador = Val(GetIniString(gIniFile, "Cheques", "BorrarPortador", ""))
         
   gPrtCheques.BordeSuperiorPCont = Val(GetIniString(gIniFile, "Cheques", "PCont-BordeSup", gPrtCheques.BordeSuperiorPCont))
   gPrtCheques.BordeIzqChequePCont = Val(GetIniString(gIniFile, "Cheques", "PCont-BordeIzq", gPrtCheques.BordeIzqChequePCont))
   gPrtCheques.BajarValDigPCont = Val(GetIniString(gIniFile, "Cheques", "PCont-BajarValDig", ""))
   gPrtCheques.MoverValDigPCont = Val(GetIniString(gIniFile, "Cheques", "PCont-MoverValDig", ""))
   gPrtCheques.BajarFechaPCont = Val(GetIniString(gIniFile, "Cheques", "PCont-BajarFecha", ""))
   gPrtCheques.MoverFechaPCont = Val(GetIniString(gIniFile, "Cheques", "PCont-MoverFecha", ""))
   gPrtCheques.BajarOrdenDePCont = Val(GetIniString(gIniFile, "Cheques", "PCont-BajarOrdenDe", ""))
   gPrtCheques.MoverOrdenDePCont = Val(GetIniString(gIniFile, "Cheques", "PCont-MoverOrdenDe", ""))
   
   gPrtCheques.BorrarALaOrden = Val(GetIniString(gIniFile, "Cheques", "PCont-BorrarOrden", ""))
   gPrtCheques.BorrarAlPortador = Val(GetIniString(gIniFile, "Cheques", "PCont-BorrarPortador", ""))

End Sub
Public Sub SetPrtData()
   Dim Nombres(7) As String
   Dim i As Integer

   For i = 0 To UBound(Nombres)
      Nombres(i) = ""
   Next i
   
   gPrtLibros.TabNombres = FillMembreteEmp(Nombres)
         
   gPrtLibros.Nombres = Nombres

   For i = 0 To UBound(Nombres)
      Nombres(i) = ""
   Next i
   
   Nombres(0) = gEmpresa.RazonSocial
   
   If gEmpresa.RutDisp = "" Then  'lo típico
      Nombres(1) = "RUT:" & vbTab & FmtCID(gEmpresa.Rut)
   Else
      Nombres(1) = "RUT:" & vbTab & FmtCID(gEmpresa.RutDisp)   'sólo para la Asoc. de AFP que tiene varias empresas con el mismo RUT
   End If
   
   If gEmpresa.Direccion <> "" And gEmpresa.Comuna <> "" Then
      Nombres(2) = "Dirección:" & vbTab & gEmpresa.Direccion & ", " & gEmpresa.Comuna
   ElseIf gEmpresa.Direccion <> "" Then
      Nombres(2) = "Dirección:" & vbTab & gEmpresa.Direccion
   Else
      Nombres(2) = "Dirección:" & vbTab & gEmpresa.Comuna
   End If
   
   Nombres(3) = "Teléfono:" & vbTab & gEmpresa.Telefono
   If gEmpresa.Fax <> "" Then
      Nombres(4) = "Fax:" & vbTab & gEmpresa.Fax
   End If
   
   gPrtReportes.Nombres = Nombres
   gPrtReportes.TabNombres = GetPrtTextWidth("Dirección:ww", False)
   
   gPrtCheques.Nombres = Nombres
   gPrtCheques.TabNombres = GetPrtTextWidth("Dirección:ww", False)
   
   gPrtLibros.PrintFecha = ChkNoPrtFecha()
   gPrtReportes.PrintFecha = ChkNoPrtFecha()

End Sub
Public Function FillMembreteEmp(Nombres() As String) As Integer
   
   Nombres(0) = "Razón Social:  " & vbTab & gEmpresa.RazonSocial
   If gEmpresa.RutDisp = "" Then  'lo típico
      Nombres(1) = "RUT:" & vbTab & FmtCID(gEmpresa.Rut)
   Else
      Nombres(1) = "RUT:" & vbTab & FmtCID(gEmpresa.RutDisp)   'sólo para la Asoc. de AFP que tiene varias empresas con el mismo RUT
   End If
   
   Nombres(2) = "Dirección:" & vbTab & gEmpresa.Direccion & ", " & IIf(gEmpresa.Ciudad <> "", FCase(gEmpresa.Ciudad), FCase(gEmpresa.Comuna))
   Nombres(3) = "Giro:" & vbTab & gEmpresa.Giro
   If gEmpresa.RepLegal1 <> "" Then
      Nombres(4) = "Rep. Legal:" & vbTab & gEmpresa.RepLegal1
      Nombres(5) = "RUT Rep. Legal:" & vbTab & FmtCID(gEmpresa.RutRepLegal1)
   End If
   
   If gEmpresa.RepConjunta = True And gEmpresa.RepLegal2 <> "" Then
      Nombres(6) = "Rep. Legal:" & vbTab & gEmpresa.RepLegal2
      Nombres(7) = "RUT Rep. Legal:" & vbTab & FmtCID(gEmpresa.RutRepLegal2)
   End If

   FillMembreteEmp = GetPrtTextWidth("RUT Rep. Legal:ww", False)
   
End Function

Public Sub SetPrtNotas(Optional ByVal EsLibro As Boolean = False)
   
   gPrtLibros.Obs = ""
   
   If EsLibro Then
      If gNotaArt100.IncluirLib Then
         gPrtLibros.Obs = gNotaArt100.TxtNota
      End If
   Else
      If gNotaArt100.IncluirBal Then
         gPrtLibros.Obs = gNotaArt100.TxtNota
      End If
   End If
   
   If EsLibro Then
      If gNotaEspecial.IncluirLib Then
         If gPrtLibros.Obs <> "" Then
            gPrtLibros.Obs = gPrtLibros.Obs & vbNewLine & vbNewLine
         End If
         gPrtLibros.Obs = gPrtLibros.Obs & gNotaEspecial.TxtNota
      End If
   Else
      If gNotaEspecial.IncluirBal Then
         If gPrtLibros.Obs <> "" Then
            gPrtLibros.Obs = gPrtLibros.Obs & vbNewLine & vbNewLine
         End If
         gPrtLibros.Obs = gPrtLibros.Obs & gNotaEspecial.TxtNota
      End If
   End If
   gPrtReportes.Obs = ""

   If gNotaArt100.IncluirInfo Then
      gPrtReportes.Obs = gNotaArt100.TxtNota
   End If
   
   If gNotaEspecial.IncluirInfo Then
      If gPrtReportes.Obs <> "" Then
         gPrtReportes.Obs = gPrtReportes.Obs & vbNewLine & vbNewLine
      End If
      gPrtReportes.Obs = gPrtReportes.Obs & gNotaEspecial.TxtNota
   End If

End Sub

Public Sub ResetPrtBas(PrtCls As ClsPrtFlxGrid)
   Dim ColWi(0) As Integer
   Dim Total(0) As String
   Dim Titulos(0) As String
   Dim FntTitulos(0) As FontDef_t
   Dim FntEncabezados(0) As FontDef_t
   Dim Encabezados(0) As String
   Dim EncabezadosCont(0) As String

   PrtCls.CallEndDoc = True
   PrtCls.ColObligatoria = 1
   PrtCls.PrintHeader = True
   PrtCls.EsContinuacion = False
   
   PrtCls.GrFontName = ""
   PrtCls.GrFontSize = -1
   PrtCls.TotFntBold = True
   
   PrtCls.InitPag = -1
   PrtCls.CellHeight = 0
   
   PrtCls.FmtCol = -1
   PrtCls.PrintFecha = ChkNoPrtFecha()
   
   PrtCls.ColWi = ColWi
   PrtCls.Titulos = Titulos
   PrtCls.Encabezados = Encabezados
   PrtCls.EncabezadosCont = EncabezadosCont
   Call PrtCls.FntTitulos(FntTitulos)
   Call PrtCls.FntEncabezados(FntEncabezados)

End Sub

Public Function ChkNoPrtFecha() As Boolean
   ChkNoPrtFecha = Not ((gEmpresa.Opciones And OPT_NOPRTFECHA) <> 0)
End Function

Public Function GetIdCuenta(NombreCuenta As String, CodigoCuenta As String, DescCuenta As String, UltimoNivel As Boolean) As Long
   Dim Rs As Recordset, Rs2 As Recordset
   Dim Q1 As String
   Dim Nivel As Integer
   
   Q1 = "SELECT IdCuenta, Codigo, Nombre, Descripcion, Nivel, Estado FROM Cuentas "
   
   If NombreCuenta <> "" Then
      Q1 = Q1 & " WHERE Nombre = '" & ParaSQL(NombreCuenta) & "'"
   Else
      Q1 = Q1 & " WHERE Codigo = '" & ReplaceStr(CodigoCuenta, "-", "") & "'"
   End If
   
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = True Then    'no existe
      DescCuenta = ""
      GetIdCuenta = 0
      UltimoNivel = False
   Else
      GetIdCuenta = vFld(Rs("IdCuenta"))
      CodigoCuenta = vFld(Rs("Codigo"))
      NombreCuenta = vFld(Rs("Nombre"))
      DescCuenta = FCase(vFld(Rs("Descripcion"), True))
      Nivel = vFld(Rs("Nivel"))
      UltimoNivel = (Nivel = gLastNivel)
      
'      Q1 = "SELECT Codigo "
'      Q1 = Q1 & " FROM Cuentas "
'      Q1 = Q1 & " WHERE Nivel = " & Nivel + 1 & " AND IdPadre = " & vFld(Rs("IdCuenta"))
'      Set Rs2 = OpenRs(DbMain, Q1)
'
'      UltimoNivel = Rs2.EOF
'
'      Call CloseRs(Rs2)
      
   End If
   
   Call CloseRs(Rs)

End Function

Public Function GetCodCuenta(ByVal IdCuenta As Long) As String
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Codigo FROM Cuentas WHERE IdCuenta = " & IdCuenta
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = True Then    'no existe
      GetCodCuenta = ""
   Else
      GetCodCuenta = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)

End Function

Public Function GetDescCuenta(ByVal IdCuenta As Long) As String
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Descripcion FROM Cuentas WHERE IdCuenta = " & IdCuenta
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = True Then    'no existe
      GetDescCuenta = ""
   Else
      GetDescCuenta = vFld(Rs(0), True)
   End If
   
   Call CloseRs(Rs)

End Function
'retorna código cuenta
'parámtros de salida:descripción IdPadre y Nivel
Public Function GetDatosCuenta(ByVal IdCuenta As Long, DescCta As String, IdPadreCta As Long, NivelCta As Integer) As String
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Codigo, Descripcion, IdPadre, Nivel FROM Cuentas WHERE IdCuenta = " & IdCuenta
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = True Then    'no existe
      GetDatosCuenta = ""
   Else
      GetDatosCuenta = vFld(Rs("Codigo"))
      DescCta = FCase(vFld(Rs("Descripcion")))
      IdPadreCta = vFld(Rs("IdPadre"))
      NivelCta = vFld(Rs("Nivel"))
   End If
   
   Call CloseRs(Rs)

End Function
Public Function GetInfoDoc(ByVal IdDoc) As String
   Dim Q1 As String
   Dim Rs As Recordset

   GetInfoDoc = ""
   
   If IdDoc <= 0 Then
      Exit Function
   End If

   Q1 = "SELECT TipoLib, TipoDoc, NumDoc, DTE FROM Documento WHERE IdDoc = " & IdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetInfoDoc = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc"))) & " " & vFld(Rs("NumDoc"))
   End If
   
   Call CloseRs(Rs)

End Function
Public Function GetTipoRazFin(ByVal IdTipoRazon As Integer) As String
   Dim i As Integer
   
   For i = 0 To UBound(gTipoRazFin)
   
      If gTipoRazFin(i).id = IdTipoRazon Then
         GetTipoRazFin = gTipoRazFin(i).Nombre
         Exit For
      End If
      
   Next i

End Function
Public Function GetCodSucursal(ByVal IdSuc As Long) As String
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT Codigo FROM Sucursales WHERE IdSucursal=" & IdSuc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetCodSucursal = vFld(Rs(0))
   Else
      GetCodSucursal = ""
   End If
   
   Call CloseRs(Rs)
   
End Function

Public Function GetIdSucursal(ByVal CodSuc As Long) As Long
   Dim Rs As Recordset
   Dim Q1 As String
   
   Q1 = "SELECT IdSucursal FROM Sucursales WHERE Codigo='" & CodSuc & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetIdSucursal = vFld(Rs(0))
   Else
      GetIdSucursal = 0
   End If
   
   Call CloseRs(Rs)
   
End Function

'función que exporta plan de cuentas a Excel en el mismo formato en que se importa
Public Function ExportarCuentas(ByVal FName As String, ByVal nNiv As Integer, Dig() As Integer, Optional nColCta As Byte = 1) As Long
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Buf As String
   Dim i As Integer
   Dim Fd As Long
   
   On Error Resume Next
   
   Fd = FreeFile
   Open FName For Output As #Fd
   If ERR Then
      MsgErr FName
      ExportarCuentas = -ERR
      Exit Function
   End If
   
   'seleccionamos los registros
   Q1 = "SELECT Codigo, Nombre, Descripcion, Atrib" & ATRIB_RUT & ", CodFECU, Atrib" & ATRIB_CAPITALPROPIO & ", TipoCapPropio, Debe, Haber, CodF22 "
   Q1 = Q1 & " FROM Cuentas "
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY Codigo"
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 0
   
   Buf = "Codigo" & vbTab & "Nombre Corto" & vbTab & "Descripcion" & vbTab & "Doc. (RUT) asociado" & vbTab & "Código FECU" & vbTab & "Capital Propio" & vbTab & "Debe" & vbTab & "Haber"
   Print #Fd, Buf

   'imprimimos el archivo
   Do While Rs.EOF = False
   
      Buf = FmtCodCuenta(vFld(Rs("Codigo"))) & vbTab & vFld(Rs("Nombre")) & vbTab & vFld(Rs("Descripcion")) & vbTab
      Buf = Buf & vFld(Rs("Atrib" & ATRIB_RUT)) & vbTab & vFld(Rs("CodFECU")) & vbTab
      
      If vFld(Rs("Atrib" & ATRIB_CAPITALPROPIO)) <> 0 Then
                  
         Select Case vFld(Rs("TipoCapPropio"))
            
            Case CAPPROPIO_ACTIVO_NORMAL
               Buf = Buf & "Normal"
               
            Case CAPPROPIO_ACTIVO_VALINTO
               Buf = Buf & "INTO"
               
            Case CAPPROPIO_ACTIVO_COMPACTIVO
               Buf = Buf & "Comp"
               
            Case CAPPROPIO_PASIVO_EXIGIBLE
               Buf = Buf & "Exig"
               
            Case CAPPROPIO_PASIVO_EXIGIBLE
               Buf = Buf & "NoExig"
               
            Case Else
               Buf = Buf & " "
               
         End Select
            
      Else
         Buf = Buf & " "
      End If
      
      Buf = Buf & vbTab
      Buf = Buf & Format(vFld(Rs("Debe")), NUMFMT) & vbTab & Format(vFld(Rs("Haber")), NUMFMT)
      Buf = Buf & vbTab & vFld(Rs("CodF22"))
   
      Print #Fd, Buf
      i = i + 1
   
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   'agregamos las cuentas básicas
   Print #Fd, "CUENTASBASICAS"
   
   Q1 = "SELECT ParamEmpresa.Tipo, ParamEmpresa.Codigo As CodParam, Cuentas.Codigo As CodCta "
   Q1 = Q1 & " FROM ParamEmpresa INNER JOIN Cuentas ON " & SqlVal("ParamEmpresa.Valor") & " = Cuentas.IdCuenta"
   Q1 = Q1 & JoinEmpAno(gDbType, "ParamEmpresa", "Cuentas")
   Q1 = Q1 & " WHERE Left(Tipo,3)='CTA'"
   Q1 = Q1 & " AND ParamEmpresa.IdEmpresa = " & gEmpresa.id & " AND ParamEmpresa.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY ParamEmpresa.Tipo, ParamEmpresa.Codigo, Cuentas.Codigo "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      Print #Fd, "CTABAS" & vbTab & vFld(Rs("Tipo")) & vbTab & vFld(Rs("CodParam")) & vbTab & vFld(Rs("CodCta"))
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Q1 = "SELECT CuentasBasicas.TipoLib, CuentasBasicas.TipoValor, Cuentas.Codigo "
   Q1 = Q1 & " FROM CuentasBasicas INNER JOIN Cuentas ON CuentasBasicas.IdCuenta = Cuentas.IdCuenta"
   Q1 = Q1 & " AND CuentasBasicas.IdEmpresa = Cuentas.IdEmpresa AND CuentasBasicas.Ano = Cuentas.Ano"
   Q1 = Q1 & JoinEmpAno(gDbType, "CuentasBasicas", "Cuentas")
   Q1 = Q1 & " WHERE CuentasBasicas.IdEmpresa = " & gEmpresa.id & " AND CuentasBasicas.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY CuentasBasicas.TipoLib, CuentasBasicas.TipoValor, Cuentas.Codigo "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Rs.EOF = False
   
      Print #Fd, "LIBRO" & vbTab & vFld(Rs("TipoLib")) & vbTab & vFld(Rs("TipoValor")) & vbTab & vFld(Rs("Codigo"))
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Close Fd

   MsgBox1 "Proceso finalizado. Se exportaron " & i & " cuentas.", vbInformation + vbOKOnly

End Function


Public Function GetAreaNegocio(ByVal CodANeg As String, Optional ByVal DescANeg As String = "") As Long
   Dim Q1 As String
   Dim Rs As Recordset

   GetAreaNegocio = 0
   
   If Trim(CodANeg) = "" And Trim(DescANeg) = "" Then
      Exit Function
   End If

   If Trim(CodANeg) <> "" Then
      Q1 = "SELECT IdAreaNegocio FROM AreaNegocio WHERE Codigo = '" & UCase(Trim(CodANeg)) & "'"
   Else
      Q1 = "SELECT IdAreaNegocio FROM AreaNegocio WHERE Descripcion = '" & Trim(DescANeg) & "'"
   End If
   
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetAreaNegocio = vFld(Rs("IdAreaNegocio"))
   End If
   
   Call CloseRs(Rs)
   
End Function


Public Function GetCentroCosto(ByVal CodCCosto As String, Optional ByVal DescCCosto As String = "") As Long
   Dim Q1 As String
   Dim Rs As Recordset

   GetCentroCosto = 0
   
   If Trim(CodCCosto) = "" And Trim(DescCCosto) = "" Then
      Exit Function
   End If

   If Trim(CodCCosto) <> "" Then
      Q1 = "SELECT IdCCosto FROM CentroCosto WHERE Codigo = '" & UCase(Trim(CodCCosto)) & "'"
   Else
      Q1 = "SELECT IdCCosto FROM CentroCosto WHERE Descripcion = '" & Trim(DescCCosto) & "'"
   End If
   
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetCentroCosto = vFld(Rs("IdCCosto"))
   End If
   
   Call CloseRs(Rs)
   
End Function

