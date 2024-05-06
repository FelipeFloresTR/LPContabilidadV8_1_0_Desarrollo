Attribute VB_Name = "PamContab"
Option Explicit

Public Const W_MONTO = 1200

Public gFCorrMon(-2 To 12) As Double


' nColCta : Numero de columnas en que se encuentra la cuenta y la descripción
'           desde esa columa vienen las columnas fijas
Public Function ImportarCuentas(ByVal fname As String, ByVal nNiv As Integer, Dig() As Integer, Optional nColCta As Byte = 1) As Long
   Dim Fd As Long, Rc As Long, Rs As Recordset, Rs2 As Recordset '2988812 variable Rs2
   Dim Q1 As String, Buf As String, Cta As String, SCta As String, Desc As String
   Dim Fld As String, LCta As String, NCta As String, Dg As String
   Dim p As Long, l As Long, i As Integer, j As Integer, c As Integer, n As Integer
   Dim Auxiliar As String, CodFECU As String, Nivel As Integer, LNivel As Long
   Dim CapProp As String, CodF22 As Integer, Clasif As String, TipoCapPropio As String, bCapProp As Boolean
   Dim bUsaRut As Boolean, AtribRut As String, Clas(20) As String, CGestion As String, ANegocios As String
   Dim Debe As Double, Haber As Double, NombreCorto As String
   Dim TipoReg As String, TipoLib As Integer, TipoValor As Integer, CtaBas As String, CodCta As String, CodTipo As String
   Dim Tbl As String, sSet As String, sFrom As String, sWhere As String
   'Const NCOLCTA = 6 ' en cuantas columnas puede venir la cuenta y la descripcion
   
   On Error Resume Next
   
   Fd = FreeFile
   Open fname For Input As #Fd
   If ERR Then
      MsgErr fname
      ImportarCuentas = -ERR
      Exit Function
   End If
   
   l = 0
   n = 0
   
   Q1 = " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Rc = DeleteSQL(DbMain, "Cuentas", Q1)
   
   Do Until EOF(Fd)
      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf
         
      p = 1
      Buf = Trim(Buf)
      
      If Buf = "" Then
         GoTo NextRec
      End If
      
      If Left(Buf, Len("CUENTASBASICAS")) = "CUENTASBASICAS" Then
         Exit Do
      End If
      
      Desc = ""
      Cta = ""
      For c = 1 To nColCta
         If Cta = "" Then  ' todas las col en que puede venir la cuenta
            Auxiliar = Trim(NextField2(Buf, p))
            ' Si no viene una cuenta pasamos al siguiente registro
            If (Auxiliar <> "" And IsNumeric(Left(Auxiliar, Dig(1))) = False) Then
               GoTo NextRec
            End If
            Cta = Auxiliar
         End If
      Next c
      
      If Cta = "" Then
         GoTo NextRec
      End If
          
      NombreCorto = Left(Trim(NextField2(Buf, p)), 10)
      Desc = Left(Trim(NextField2(Buf, p)), 100)
      AtribRut = Trim(NextField2(Buf, p))
      bUsaRut = (AtribRut <> "" And AtribRut <> "0")
      CodFECU = Trim(NextField2(Buf, p))
      CapProp = Trim(NextField2(Buf, p))
      'CGestion = NextField2(Buf, p)
      'ANegocios = NextField2(Buf, p)
      Debe = vFmt(NextField2(Buf, p))
      Haber = vFmt(NextField2(Buf, p))
      CodF22 = Val(NextField2(Buf, p))
   
      Debug.Print Cta & " ; " & Desc
'      If Cta >= "3" Then
'         Beep
'      End If
      
      NCta = GetCodCta(Cta, Nivel, nNiv, Dig())
'      j = 1
'      Nivel = 0
'      For i = 1 To nNiv
'         Dg = Mid(Cta, j, Dig(i))
'         NCta = NCta & Dg
'         j = j + Dig(i) + 1
'         If Val(Dg) Then
'            Nivel = i
'         End If
'      Next i
   
      Debug.Print NCta & " ; " & Desc & " ; " & CodFECU
      
      '2988812
      'If NCta < LCta Then
      If IIf(IsNumeric(Replace(Cta, "-", "")), CLng(Replace(Cta, "-", "")), Cta) < CLng(IIf(LCta <> "", LCta, "0")) Then
         MsgBox1 "Linea " & l & ": La cuenta " & NCta & " no debería estar después de la cuenta " & LCta, vbExclamation
      End If
      'Fin '2988812
   
      If Nivel = 1 Then
         If InStr(1, Desc, "activ", vbTextCompare) Then
            Clasif = CLASCTA_ACTIVO
            Clas(CLASCTA_ACTIVO) = NCta
         ElseIf InStr(1, Desc, "pasiv", vbTextCompare) Then
            Clasif = CLASCTA_PASIVO
            Clas(CLASCTA_PASIVO) = NCta
         ElseIf InStr(1, Desc, "orden", vbTextCompare) Then
            Clasif = CLASCTA_ORDEN
            Clas(CLASCTA_ORDEN) = NCta
         Else
            Clasif = CLASCTA_RESULTADO
            Clas(CLASCTA_RESULTADO) = NCta
         End If
      Else
         Clasif = "NULL"
      End If
            
      If Nivel = nNiv And CapProp <> "" Then
         If InStr(1, CapProp, "into", vbTextCompare) Then
            TipoCapPropio = CAPPROPIO_ACTIVO_VALINTO
         ElseIf InStr(1, CapProp, "comp", vbTextCompare) Then
            TipoCapPropio = CAPPROPIO_ACTIVO_COMPACTIVO
         ElseIf InStr(1, CapProp, "noexig", vbTextCompare) Then
            TipoCapPropio = CAPPROPIO_PASIVO_NOEXIGIBLE
         ElseIf InStr(1, CapProp, "exig", vbTextCompare) Then
            TipoCapPropio = CAPPROPIO_PASIVO_EXIGIBLE
         ElseIf InStr(1, CapProp, "normal", vbTextCompare) Then
            TipoCapPropio = CAPPROPIO_ACTIVO_NORMAL
         End If
      ElseIf Debe <> 0 Or Haber <> 0 Then
         MsgBox1 "Linea " & l & ": La cuenta " & NCta & " no debería tener saldo inicial. Se dejará en cero.", vbExclamation
         Debe = 0
         Haber = 0
         TipoCapPropio = "NULL"
      Else
         TipoCapPropio = "NULL"
      End If
      bCapProp = (TipoCapPropio <> "NULL")

      '2988812
      Q1 = "SELECT IdCuenta FROM Cuentas WHERE Codigo = '" & Trim(NCta) & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs2 = OpenRs(DbMain, Q1)
         
      If Rs2.EOF Then
      
        Q1 = "INSERT INTO Cuentas (IdEmpresa, Ano, Codigo, Nombre, Descripcion, CodFecu, Nivel, Estado, Clasificacion, TipoCapPropio, CodF22, Atrib" & ATRIB_CAPITALPROPIO & ", Atrib" & ATRIB_RUT & ", Debe, Haber )"
        Q1 = Q1 & " VALUES (" & gEmpresa.id & "," & gEmpresa.Ano & ",'" & Trim(NCta) & " ','" & ParaSQL(NombreCorto) & "','" & ParaSQL(Desc) & "','" & Trim(CodFECU) & "'," & Nivel
        Q1 = Q1 & ", 1," & Clasif & "," & TipoCapPropio & "," & CodF22 & "," & Abs(bCapProp) & "," & Abs(bUsaRut)
        Q1 = Q1 & ", " & Str0(Debe) & "," & Str0(Haber)
        Q1 = Q1 & " )"
        Rc = ExecSQL(DbMain, Q1)
        n = n + 1
      
      End If
      Call CloseRs(Rs2)
      LCta = Replace(NCta, "-", "")
      LNivel = Nivel
      'Fin 2988812
   
NextRec:
   Loop

   'Close Fd        'no se cierra el archivo poque más abajo se importan las cuentas básicas, si hay

   MsgBox1 "Se importaron " & n & " cuentas", vbInformation

   ' Ponemos los padres
   p = Len(LCta)
   
   j = 0
   For i = 1 To nNiv - 1
   
      j = j + Dig(i)
            
'      Q1 = "UPDATE Cuentas INNER JOIN Cuentas AS CtasPadre ON (Cuentas.Nivel - 1 = CtasPadre.Nivel)"
'      Q1 = Q1 & " AND Left(Cuentas.Codigo, " & j & ") & '" & String(p - j, "0") & "' = CtasPadre.Codigo"
'      Q1 = Q1 & " AND Cuentas.IdEmpresa = CtasPadre.IdEmpresa AND Cuentas.Ano = CtasPadre.Ano"
'      Q1 = Q1 & " SET Cuentas.idPadre = CtasPadre.idCuenta, Cuentas.Clasificacion = CtasPadre.Clasificacion"
'      Q1 = Q1 & " WHERE Cuentas.Nivel=" & i + 1 & " AND CtasPadre.Nivel=" & i
'      Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
'      Rc = ExecSQL(DbMain, Q1)
      
      Tbl = "Cuentas"
      sFrom = " Cuentas INNER JOIN Cuentas AS CtasPadre ON (Cuentas.Nivel - 1 = CtasPadre.Nivel)"
      sFrom = sFrom & " AND " & SqlConcat(gDbType, "Left(Cuentas.Codigo, " & j & ")", "'" & String(p - j, "0") & "'") & " = CtasPadre.Codigo"
      sFrom = sFrom & " AND Cuentas.IdEmpresa = CtasPadre.IdEmpresa AND Cuentas.Ano = CtasPadre.Ano"
      sSet = " Cuentas.idPadre = CtasPadre.idCuenta, Cuentas.Clasificacion = CtasPadre.Clasificacion"
      sWhere = " WHERE Cuentas.Nivel=" & i + 1 & " AND CtasPadre.Nivel=" & i
      sWhere = sWhere & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
      Rc = UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   Next i

   Q1 = "SELECT Count(*) as N FROM Cuentas WHERE Nivel <> 1 AND (idPadre Is NULL or idPadre=0)"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   Rc = Rs("N")
   Call CloseRs(Rs)
   
   If Rc Then
      Q1 = "SELECT Codigo FROM Cuentas WHERE Nivel <> 1 AND (idPadre Is NULL or idPadre=0)"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      Auxiliar = ""
      i = 0
      Do Until Rs.EOF Or i > 20
         Auxiliar = Auxiliar & ", " & vFld(Rs("Codigo"))
         i = i + 1
         
         Rs.MoveNext
      Loop
      
      Call CloseRs(Rs)

      MsgBox1 "Quedaron " & Rc & " cuentas sin padre directo." & vbLf & "Algunas son:" & Mid(Auxiliar, 2), vbExclamation
   End If
   
   ImportarCuentas = Rc

   ' ponemos los nietos si no tienen padres
'   j = 0
'   For i = 1 To nNiv - 2
'
'      j = j + Dig(i)
'
'      Q1 = "UPDATE Cuentas INNER JOIN Cuentas AS CtasPadre ON (Cuentas.Nivel - 2= CtasPadre.Nivel)"
'      Q1 = Q1 & " AND Left(Cuentas.Codigo, " & j & ") = Left(CtasPadre.Codigo, " & j & ")"
'      Q1 = Q1 & " SET Cuentas.idPadre = CtasPadre.idCuenta, Cuentas.Clasificacion = CtasPadre.Clasificacion"
'      Q1 = Q1 & " WHERE Cuentas.idPadre = 0 and Cuentas.Nivel=" & i + 2 & " AND CtasPadre.Nivel=" & i
'      Rc = ExecSQL(DbMain, Q1)
'
'   Next i
   
   ' Propagamos la clasificación a los hijos
   For i = 1 To nNiv - 1

'      Q1 = "UPDATE Cuentas INNER JOIN Cuentas AS CtasPadre ON Cuentas.idPadre = CtasPadre.idCuenta"
'      Q1 = Q1 & " AND Cuentas.IdEmpresa = CtasPadre.IdEmpresa AND Cuentas.Ano = CtasPadre.Ano"
'      Q1 = Q1 & " SET Cuentas.Clasificacion = CtasPadre.Clasificacion"
'      Q1 = Q1 & " WHERE Cuentas.Clasificacion IS NULL AND CtasPadre.Nivel=" & i
'      Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
'      Rc = ExecSQL(DbMain, Q1)
      
      Tbl = "Cuentas"
      sFrom = " Cuentas INNER JOIN Cuentas AS CtasPadre ON Cuentas.idPadre = CtasPadre.idCuenta "
      sFrom = sFrom & " AND Cuentas.IdEmpresa = CtasPadre.IdEmpresa AND Cuentas.Ano = CtasPadre.Ano"
      sSet = " Cuentas.Clasificacion = CtasPadre.Clasificacion"
      sWhere = " WHERE Cuentas.Clasificacion IS NULL AND CtasPadre.Nivel=" & i
      sWhere = sWhere & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
      Rc = UpdateSQL(DbMain, Tbl, sSet, sFrom, sWhere)
   Next i
   
   'importamos Cuentas Básicas, si hay
   If EOF(Fd) Then
      Close Fd
      Exit Function
   End If
   
   If Left(Buf, Len("CUENTASBASICAS")) <> "CUENTASBASICAS" Then
      Close Fd
      Exit Function
   End If
      
   Do Until EOF(Fd)
   
      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf
         
      p = 1
      Buf = Trim(Buf)

      TipoReg = Trim(NextField2(Buf, p))
      If TipoReg = "CTABAS" Then
         CtaBas = Trim(NextField2(Buf, p))
         CodTipo = Trim(NextField2(Buf, p))
         CodCta = Trim(NextField2(Buf, p))
         
         Q1 = "SELECT IdCuenta FROM Cuentas WHERE Codigo = '" & CodCta & "'"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = False Then
            Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor, IdEmpresa, Ano) VALUES('" & CtaBas & "'," & SqlVal(CodTipo) & ", '" & vFld(Rs("IdCuenta")) & "'," & gEmpresa.id & "," & gEmpresa.Ano & ")"
            Call ExecSQL(DbMain, Q1)
         End If
         
         Call CloseRs(Rs)
         
      ElseIf TipoReg = "LIBRO" Then
         TipoLib = Val(Trim(NextField2(Buf, p)))
         TipoValor = Val(Trim(NextField2(Buf, p)))
         CodCta = Trim(NextField2(Buf, p))
         
         If TipoLib >= LIB_COMPRAS And TipoLib <= LIB_RETEN Then
         
            Q1 = "SELECT IdCuenta FROM Cuentas WHERE Codigo = '" & CodCta & "'"
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Set Rs = OpenRs(DbMain, Q1)
            
            If Rs.EOF = False Then
               Q1 = "INSERT INTO CuentasBasicas (Tipo, TipoLib, TipoValor, IdCuenta, IdEmpresa, Ano) VALUES( 0, " & TipoLib & "," & TipoValor & "," & vFld(Rs("IdCuenta")) & "," & gEmpresa.id & "," & gEmpresa.Ano & ")"
               Call ExecSQL(DbMain, Q1)
            End If
            
            Call CloseRs(Rs)
         End If
         
      End If
      
   Loop
   
   Close Fd
   
   MsgBox1 "Las Cuentas Básicas fueron asignadas.", vbInformation + vbOKOnly
   
End Function

'Importar comprobante tipo

Public Function ImportarComprobante(ByVal fname As String, ByVal nNiv As Integer, Dig() As Integer) As Long
   Dim Fd As Long, Rc As Long, Q1 As String, Rs As Recordset
   Dim TipoComp As Byte, GlosaComp As String, c As Integer, Ctas(100, 2) As String
   Dim i As Integer, Buf As String, Aux As String, l As Long, p As Long, Aux2 As String
   Dim Nivel As Integer, idcomp As Long, NomComp As String, TComp As String

   Fd = FreeFile
   Open fname For Input As #Fd
   If ERR Then
      MsgErr fname
      ImportarComprobante = -ERR
      Exit Function
   End If

   i = rInStr(fname, "\")
   If i Then
      NomComp = Mid(fname, i + 1)
   Else
      NomComp = fname
   End If
   i = InStr(NomComp, ".")
   NomComp = Left(NomComp, i - 1)
   NomComp = Left(NomComp, 30)
   
'   NomComp = ReplaceStr(NomComp, " ", "")
'   NomComp = Left(ReplaceStr(NomComp, " ", ""), 15)
   
   Q1 = "SELECT idComp FROM CT_Comprobante WHERE Nombre='" & NomComp & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   i = Rs.EOF
   Call CloseRs(Rs)
   If i = 0 Then
      ImportarComprobante = -4
      Close #Fd
      MsgBox1 "El nombre de comprobante '" & NomComp & "' ya existe, cambie el nombre del archivo e impórtelo nuevamente.", vbExclamation
      Exit Function
   End If
      
   c = 0
   
   Do Until EOF(Fd)
   
      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf
         
      p = 1
      Buf = Trim(Buf)
      
      If TipoComp = 0 Then
         Aux = vbTab & "Tipo:" & vbTab
         i = InStr(1, Buf, Aux, vbTextCompare)
         If i Then
            p = i + Len(Aux)
            Aux = UCase(Trim(NextField2(Buf, p)))
            Select Case Aux
               Case "INGRESO"
                  TipoComp = TC_INGRESO
                  TComp = "I"
            
               Case "EGRESO"
                  TipoComp = TC_EGRESO
                  TComp = "E"
            
               Case "TRASPASO"
                  TipoComp = TC_TRASPASO
                  TComp = "T"
            End Select
            
         End If
      ElseIf GlosaComp = "" Then
         Aux = vbTab & "Glosa:" & vbTab
         i = InStr(1, Buf, Aux, vbTextCompare)
         If i Then
            p = i + Len(Aux)
            Aux = UCase(Trim(NextField2(Buf, p)))
         
            Aux2 = ReplaceStr(Mid(Buf, p), vbTab, " ")
         ElseIf Aux2 <> "" Then
            Aux = ReplaceStr(Buf, vbTab, " ")
            i = InStr(1, Buf, "Nombre de la Cuenta", vbTextCompare)
            If i <> 0 Or Trim(Aux) = "" Then  ' se acabó la glosa
               GlosaComp = ReplaceStr(Aux2, "  ", " ")
               GlosaComp = ReplaceStr(GlosaComp, "  ", " ")
            Else
               Aux2 = Aux2 & " " & Aux
            End If
         End If
      Else
              
         p = 1
         Aux = NextField2(Buf, p)   ' col vacía
         Aux = NextField2(Buf, p)   ' orden
         Aux = NextField2(Buf, p)   ' cta ?

         If IsNumeric(Left(Aux, Dig(1))) = False Then
            GoTo NextRec
         End If
      
         Ctas(c, 0) = GetCodCta(Aux, Nivel, nNiv, Dig())
         If Ctas(c, 0) = "" Then
            GoTo NextRec
         End If
      
         Aux = NextField2(Buf, p)         ' CCosto
         Aux = NextField2(Buf, p)         ' nombre cta
         Ctas(c, 1) = NextField2(Buf, p)  ' glosa trans
         
         Q1 = "SELECT idCuenta FROM Cuentas WHERE Codigo='" & Ctas(c, 0) & "'"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         If Rs.EOF = False Then
            Ctas(c, 2) = vFld(Rs("idCuenta"))
            c = c + 1
         Else
            MsgBox1 "Línea " & l & ": La cuenta '" & Ctas(c, 0) & " - " & Aux & "', no existe en el plan de cuentas.", vbExclamation
         End If
         Call CloseRs(Rs)
         
   
      End If

NextRec:
   Loop
   Close #Fd

   If TipoComp = 0 Then
      ImportarComprobante = -1
   End If

   If Trim(GlosaComp) = "" Then
      ImportarComprobante = -2
   End If

   If c <= 0 Then
      ImportarComprobante = -3
   End If

   idcomp = AdvTbAddNew(DbMain, "CT_Comprobante", "IdComp", "IdEmpresa", gEmpresa.id)
   Q1 = "UPDATE CT_Comprobante SET Tipo=" & TipoComp & ",Glosa='" & ParaSQL(Left(FCase(GlosaComp), 100)) & "'"
   Q1 = Q1 & ", Correlativo=" & idcomp & ", Nombre='" & TComp & idcomp & "', Descrip='" & ParaSQL(Left(FCase(NomComp), 40)) & "'"
   Q1 = Q1 & ", Estado=" & EC_PENDIENTE
   Q1 = Q1 & " WHERE idComp=" & idcomp
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id

   Rc = ExecSQL(DbMain, Q1)

   Aux = "INSERT INTO CT_MovComprobante (IdEmpresa, idComp, Orden, idCuenta, CodCuenta, Glosa) VALUES (" & gEmpresa.id & ","

   For i = 0 To c - 1
   
      Q1 = Aux & idcomp & "," & i + 1 & "," & Ctas(i, 2) & ",'" & ParaSQL(Ctas(i, 0)) & "','" & ParaSQL(Left(Ctas(i, 1), 50)) & "' )"
      Rc = ExecSQL(DbMain, Q1)
   
   Next i



   ImportarComprobante = idcomp

End Function
' Se asume que Cta viene con separadores tipo "-"
Public Function GetCodCta(ByVal Cta As String, Nivel As Integer, ByVal nNiv As Integer, Dig() As Integer) As String
   Dim CodCta As String, i As Integer, j As Integer, Dg As String
   
   CodCta = ""
   j = 1
   Nivel = 0
   For i = 1 To nNiv
      Dg = Mid(Cta, j, Dig(i))
      If IsNumeric(Dg) = False Then
         GetCodCta = ""
         Exit Function
      End If
      CodCta = CodCta & Dg
      j = j + Dig(i) + 1
      If Val(Dg) Then
         Nivel = i
      End If
   Next i

   GetCodCta = CodCta

End Function

Public Function ImportarEntidades(ByVal fname As String)
   Dim Fd As Long, Rc As Long, Q1 As String, Buf As String, l As Long, p As Long
   Dim QI As String, Aux As String, Rs As Recordset, i As Integer, r As Integer
   Dim AuxRut As String, NotValidRut As Boolean
   
   Dim RutEnt As String, CodEnt As String, NomEnt As String, DirEnt As String
   Dim RegEnt As Integer, ComuEnt As Integer, CiuEnt As String, TelEnt As String, FaxEnt As String
   Dim CodActEconEnt As String, DirPostEnt As String, ComuPostEnt As String, emailEnt As String
   Dim UrlEnt As String, ObsEnt As String, TipoEnt(MAX_ENTCLASIF) As Byte
   Dim Giro As String
   Dim SinClasif As Boolean
   Dim EsSupermercado As Boolean
   Dim IsValidRUT As Boolean, NoEsRUT As Integer
   Dim NSinClasif As Integer, NError As Integer
   Dim Msg As String
   Dim FranqTribEnt As Integer, EntRelacionada As Boolean, Ret3Porc As Boolean
   
   Fd = FreeFile
   Open fname For Input As #Fd
   If ERR Then
      MsgErr fname
      ImportarEntidades = -ERR
      Exit Function
   End If

   QI = "INSERT INTO Entidades (IdEmpresa, Rut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, Giro, DomPostal, ComPostal, Email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, EsSupermercado, NotValidRut, EntRelacionada, FranqTribEnt, Ret3Porc ) VALUES (" & gEmpresa.id & ","
   
   Do Until EOF(Fd)
      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf
         
      p = 1
      Buf = Trim(Buf)
      
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 And InStr(1, Buf, "Nombre", vbTextCompare) Then
         GoTo NextRec
      End If
      
      NotValidRut = False
      AuxRut = Trim(NextField2(Buf, p))
      If Len(AuxRut) < 2 Then ' 26 feb 2020: se agrega verificación
            If MsgBox1("Línea " & l & ": RUT vacío en la columna 1 [" & AuxRut & "].", vbExclamation + vbOKCancel) = vbCancel Then
               Exit Do
            End If
            GoTo NextRec
      End If
      
      NoEsRUT = (Val(NextField2(Buf, p)) <> 0) '
      
      IsValidRUT = False
      If NoEsRUT = 0 Then
         IsValidRUT = ValidRut(AuxRut)
         If Not IsValidRUT Then
            NError = NError + 1
            If MsgBox1("Línea " & l & ": RUT inválido.", vbExclamation + vbOKCancel) = vbCancel Then
               Exit Do
            End If
            GoTo NextRec
         End If
      End If
      
      If IsValidRUT Then
         RutEnt = vFmtCID(AuxRut)
      Else
         RutEnt = AuxRut
         NotValidRut = True
      End If
      
      If RutEnt = "0" Or RutEnt = "" Then
         NError = NError + 1
         If MsgBox1("Línea " & l & ": Falta el RUT de la entidad", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      CodEnt = Trim(NextField2(Buf, p))
      NomEnt = Trim(NextField2(Buf, p))
            
      If RutEnt = "" Then
         NError = NError + 1
         If MsgBox1("Línea " & l & ": Falta el RUT de la entidad", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      If CodEnt = "" Then
         NError = NError + 1
         If MsgBox1("Línea " & l & ": Falta el código de la entidad", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      If NomEnt = "" Then
         NError = NError + 1
         If MsgBox1("Línea " & l & ": Falta el nombre de la entidad", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      Q1 = "SELECT idEntidad FROM Entidades WHERE (RUT='" & RutEnt & "' OR Codigo='" & CodEnt & "')"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      Set Rs = OpenRs(DbMain, Q1)
      i = Rs.EOF
      Call CloseRs(Rs)
      If i = False Then
         NError = NError + 1
         If MsgBox1("Línea " & l & ": La entidad '" & NomEnt & "' (RUT=" & RutEnt & ", Código=" & CodEnt & ") ya existe.", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      DirEnt = Trim(NextField2(Buf, p))
      
      Aux = Trim(NextField2(Buf, p))
      If Aux <> "" Then
         Q1 = "SELECT Id, Codigo FROM Regiones WHERE Comuna='" & UCase(Aux) & "'"
         Set Rs = OpenRs(DbMain, Q1)
         If Rs.EOF = False Then
            RegEnt = vFld(Rs("Codigo"))
            ComuEnt = vFld(Rs("id"))
         Else
            NError = NError + 1
            If MsgBox1("Línea " & l & ": No se encontró la comuna '" & Aux & "' en la tabla de comunas.", vbExclamation + vbOKCancel) = vbCancel Then
               Exit Do
            End If
            RegEnt = -1
            ComuEnt = -1
         End If
         Call CloseRs(Rs)
      Else
         RegEnt = -1
         ComuEnt = -1
      End If

      CiuEnt = Trim(NextField2(Buf, p))
      TelEnt = Trim(NextField2(Buf, p))
      FaxEnt = Trim(NextField2(Buf, p))
'      CodActEconEnt = Trim(NextField2(Buf, p))
'      If CodActEconEnt <> "" Then
'         Q1 = "SELECT Codigo FROM CodActiv WHERE Codigo='" & CodActEconEnt & "'"
'         Set Rs = OpenRs(DbMain, Q1)
'
'         If Rs.EOF Then
'            MsgBox1 "Línea " & l & ": No se encontró la actividad económica '" & CodActEconEnt & "' en la tabla de actividades.", vbExclamation
'            CodActEconEnt = ""
'         End If
'         Call CloseRs(Rs)
'      End If
      Giro = Trim(NextField2(Buf, p))
      DirPostEnt = Trim(NextField2(Buf, p))
      ComuPostEnt = Trim(NextField2(Buf, p))
      emailEnt = Trim(NextField2(Buf, p))
      UrlEnt = Trim(NextField2(Buf, p))
      ObsEnt = Trim(NextField2(Buf, p))
      SinClasif = True

      'clasificación de la entidad
      For i = 0 To MAX_ENTCLASIF
         
         Aux = LCase(Trim(NextField2(Buf, p)))
         TipoEnt(i) = Abs(Aux = "x" Or Val(Aux) <> 0)
         If TipoEnt(i) <> 0 Then
            SinClasif = False
         End If
      Next i
      
      If SinClasif Then
         TipoEnt(0) = 1
         NSinClasif = NSinClasif + 1
         NError = NError + 1
         If MsgBox1("Línea " & l & ": Falta la clasificación de la entidad '" & NomEnt & "' (RUT=" & RutEnt & ", Código=" & CodEnt & ").", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If

     'Es supermercado?
      Aux = LCase(Trim(NextField2(Buf, p)))
      EsSupermercado = Abs(Aux = "x" Or Val(Aux) <> 0)
      
     'Se acoge a normas 14 TER / 14 D
      Aux = LCase(Trim(NextField2(Buf, p)))
      EntRelacionada = Abs(Aux = "x" Or Val(Aux) <> 0)

     'Franquicia Tributaria
      FranqTribEnt = Val(NextField2(Buf, p))
      If FranqTribEnt < 0 Or FranqTribEnt > UBound(gFranqTribEnt) Then
         If MsgBox1("Línea " & l & ": Valor franquicia tributaria inválido '" & FranqTribEnt & "'.", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If

     'Aplica a Retención 3% Préstamo Solidario
      Aux = LCase(Trim(NextField2(Buf, p)))
      Ret3Porc = Abs(Aux = "x" Or Val(Aux) <> 0)

      Q1 = "'" & RutEnt & "'"
      Q1 = Q1 & ",'" & CodEnt & "'"
      Q1 = Q1 & ",'" & NomEnt & "'"
      Q1 = Q1 & ",'" & DirEnt & "'"
      Q1 = Q1 & "," & RegEnt
      Q1 = Q1 & "," & ComuEnt
      Q1 = Q1 & ",'" & CiuEnt & "'"
      Q1 = Q1 & ",'" & TelEnt & "'"
      Q1 = Q1 & ",'" & FaxEnt & "'"
      Q1 = Q1 & ",'" & Giro & "'"
      Q1 = Q1 & ",'" & DirPostEnt & "'"
      Q1 = Q1 & ",'" & ComuPostEnt & "'"
      Q1 = Q1 & ",'" & emailEnt & "'"
      Q1 = Q1 & ",'" & UrlEnt & "'"
      Q1 = Q1 & "," & EE_ACTIVO
      Q1 = Q1 & ",'" & ObsEnt & "'"
      For i = 0 To MAX_ENTCLASIF
         Q1 = Q1 & "," & TipoEnt(i)
      Next i
      Q1 = Q1 & "," & Abs(EsSupermercado)
      Q1 = Q1 & "," & IIf(NotValidRut <> 0, 1, 0)
      Q1 = Q1 & "," & Abs(EntRelacionada)
      Q1 = Q1 & "," & FranqTribEnt
      Q1 = Q1 & "," & IIf(Ret3Porc, 1, 0)       'FCA - 12/10/2021
      Q1 = Q1 & " )"
      
      Debug.Print Q1

      Rc = ExecSQL(DbMain, QI & Q1)
      r = r + 1

NextRec:
   Loop

   Close #Fd
   
   If NSinClasif > 0 Then
      MsgBox1 "Se encontraron " & NSinClasif & " Entidades que no tienen clasificación. Estas entidades no han sido importadas al sistema. " & vbCrLf & vbCrLf & "Se recomienda corregir el archivo y volver a importarlo.", vbExclamation
   End If
   
   If r > 0 Then
      Msg = "Proceso de importación finalizado. Se importaron " & r & " entidades exitosamente."
   Else
      Msg = "Proceso de importación finalizado sin entidades nuevas ingresadas al sistema."
   End If
   
   If NError > 0 Then
      Msg = Msg & vbCrLf & vbCrLf & "Si hubo algún error, se recomienda corregir el archivo y volver a importarlo."
   End If

   MsgBox1 Msg, vbInformation
   
   ImportarEntidades = r

End Function

Public Sub Grid2Html(Gr As Control, ByVal Title As String, Optional ByVal bIncludeCero As Boolean = 0)
   
   On Error Resume Next

   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.DialogTitle = "Guardar como..."
   FrmMain.Cm_ComDlg.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
   FrmMain.Cm_ComDlg.Filter = "Archivos HTML (htm, html)|*.htm;*.html|Todos (*.*)|*.*"
   FrmMain.Cm_ComDlg.DefaultExt = "htm"

   FrmMain.Cm_ComDlg.ShowSave
   
   If ERR = 0 Then
      Call FGr2Html(Gr, FrmMain.Cm_ComDlg.Filename, Title, bIncludeCero)
   End If

End Sub

Public Sub FillCtasConcil(CbCuentas As ComboBox, Optional ByVal Def As Long = 1, Optional ByVal TipoLib As Long = 0)
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   Dim DBwhere As String
   
   CbCuentas.Clear
   
   If gDbType = SQL_ACCESS Then
        DBwhere = " Where cint(idcuenta) in (Select cint(Valor) from Paramempresa Where Tipo in ('CTAODFACTI','CTAODFPASI')) "
   Else
        DBwhere = " Where idcuenta in (Select Valor from Paramempresa Where Tipo in ('CTAODFACTI','CTAODFPASI')) "
   End If
   
   Q1 = "SELECT Codigo, idCuenta, Descripcion FROM Cuentas"
   If TipoLib <> 0 Then
    Q1 = Q1 & " WHERE Atrib" & ATRIB_RUT & "=" & ATRIB_CONCILIACION
   Else
    Q1 = Q1 & " WHERE Atrib" & ATRIB_CONCILIACION & "=" & ATRIB_CONCILIACION
   End If
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   'se comenta ya en sql se cae ffv 2861591
  ' Q1 = Q1 & " ORDER BY Codigo "
   ' ffv 2861591
   If TipoLib <> 0 Then
    Q1 = Q1 & " Union All "
    Q1 = Q1 & " Select Codigo, idCuenta, Descripcion "
    Q1 = Q1 & " From Cuentas "
    Q1 = Q1 & DBwhere
    Q1 = Q1 & " AND Atrib" & ATRIB_RUT & "<>" & ATRIB_CONCILIACION
    Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
     'se comenta ya en sql se cae ffv 2861591
  ' Q1 = Q1 & " ORDER BY Codigo "
   ' ffv 2861591
   End If
   Q1 = Q1 & " ORDER BY Codigo "
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Def = 1 Then
      CbCuentas.AddItem "(Todas las cuentas que se concilian)"
      CbCuentas.ItemData(CbCuentas.NewIndex) = 0
   Else
      CbCuentas.AddItem ""
      CbCuentas.ItemData(CbCuentas.NewIndex) = 0
   End If
   
   If CbCuentas.ListCount > 0 Then
      CbCuentas.ListIndex = CbCuentas.NewIndex
   End If
   
   Do While Rs.EOF = False
      CbCuentas.AddItem Format(vFld(Rs("Codigo")), gFmtCodigoCta) & " - " & FCase(vFld(Rs("Descripcion"), True))
      CbCuentas.ItemData(CbCuentas.NewIndex) = vFld(Rs("idCuenta"))
      
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   
     
End Sub

Public Sub FillParam(CbParam As ComboBox, Tipo As String, Optional ByVal Def As Long = 1)
   Dim Q1 As String
   Dim Rs As Recordset
   
   CbParam.Clear
   
   Q1 = "SELECT CODIGO, VALOR"
   Q1 = Q1 & " From Param "
   Q1 = Q1 & " WHERE TIPO = '" & Tipo & "' "
   Q1 = Q1 & " AND TIPO IS NOT NULL"
   Set Rs = OpenRs(DbMain, Q1)

   
   Do While Rs.EOF = False
      CbParam.AddItem vFld(Rs("VALOR"))
      CbParam.ItemData(CbParam.NewIndex) = vFld(Rs("CODIGO"))
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
   
     
End Sub



Public Sub FillCorrMon(ByVal Ano As Integer)
   Dim Q1 As String, Dt1 As Long, Dt2 As Long, m As Integer, pIPC0 As Double
   Dim Rs As Recordset
   
   Dt1 = DateSerial(Ano - 1, 11, 1)
   Dt2 = DateSerial(Ano, 12, 1)
   
   pIPC0 = -64
   
   For m = -2 To UBound(gFCorrMon)
      gFCorrMon(m) = 1
   Next m
   
   Q1 = "SELECT AnoMes, pIPC FROM IPC WHERE AnoMes BETWEEN " & Dt1 & " AND " & Dt2 & " ORDER BY AnoMes DESC"
   Set Rs = OpenRs(DbMain, Q1)
   Do Until Rs.EOF
   
      Dt1 = vFld(Rs("AnoMes"))
      m = (Year(Dt1) - Ano) * 12 + month(Dt1)
      
      If pIPC0 = -64 Then
         pIPC0 = vFld(Rs("pIPC"))
         gFCorrMon(m) = 1
      ElseIf vFld(Rs("pIPC")) Then
         gFCorrMon(m) = pIPC0 / vFld(Rs("pIPC"))
      Else
         gFCorrMon(m) = 1
      End If
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)

End Sub

Public Function GetFactCorrMon(ByVal Fecha As Double) As Double
   Dim m As Integer

   m = (Year(Fecha) - gEmpresa.Ano) * 12 + month(Fecha)

   GetFactCorrMon = gFCorrMon(m)
   
End Function

Public Function LinkLau() As Integer

#If DATACON = 1 Then       'Access
   
   Dim Buf As String, Rc As Long

   LinkLau = 0
   
   Call UnLinkLau
   
   If gLinkF22 = False Then
      LinkLau = -1
      Exit Function
   End If

   If Left(gHRPath, 3) <> Left(gDbPath, 3) Then
      MsgBox1 "¡ ATENCIÓN !" & vbCrLf & "La ubicación de los productos HyperRenta no permite interactuar con ellos si hay varios usuarios trabajando con Contabilidad.", vbCritical
   End If

   ' Linkeamos las tablas del Lau

   Buf = gHRPath & "\RUTS\" & Right("000000000" & gEmpresa.Rut, 8) & "\Lau" & gEmpresa.Ano & ".MDB"
   
   Rc = True
   
   Rc = Rc And LinkMdbTable(DbMain, Buf, "FacturaCompras", "LAU_FacturaCompras", , False)
   Rc = Rc And LinkMdbTable(DbMain, Buf, "mPersonas", "LAU_mPersonas", , False)
   Rc = Rc And LinkMdbTable(DbMain, Buf, "ReferenciaExp", "LAU_ReferenciaExp", , False)
   Rc = Rc And LinkMdbTable(DbMain, Buf, "VentasFacturaNac", "LAU_VentasFacturaNac", , False)
   Rc = Rc And LinkMdbTable(DbMain, Buf, "VentasFacturaExp", "LAU_VentasFacturaExp", , False)
   Rc = Rc And LinkMdbTable(DbMain, Buf, "VentasBoletas", "LAU_VentasBoletas", , False)
   Rc = Rc And LinkMdbTable(DbMain, Buf, "VentasDevolucion", "LAU_VentasDevolucion", , False)
   Rc = Rc And LinkMdbTable(DbMain, Buf, "Retenciones", "LAU_Retenciones", , False)
   
   'tabla Sucursales de LAU
   Buf = gHRPath & "\RUTS\" & Right("000000000" & gEmpresa.Rut, 8) & "\DefL" & gEmpresa.Ano & ".MDB"
   Rc = Rc And LinkMdbTable(DbMain, Buf, "Sucursales", "LAU_Sucursales", , False)
   
   If Rc = False Then
      LinkLau = -2
      If ERR = 3024 Or ERR = 3044 Then ' archivo o path no existe
         MsgBox1 "Aún no se ha abierto el año de este contribuyente con el producto HR-IVA.", vbExclamation
      End If
      Exit Function
   End If
   
#End If

End Function

Public Function UnLinkLau() As Integer
#If DATACON = 1 Then       'Access
   Dim Rc As Long

   UnLinkLau = 0
   
   ' Des-Linkeamos las tablas del Lau
   
   Rc = True
   
   Rc = Rc And UnLinkTable(DbMain, "LAU_FacturaCompras")
   Rc = Rc And UnLinkTable(DbMain, "LAU_mPersonas")
   Rc = Rc And UnLinkTable(DbMain, "LAU_ReferenciaExp")
   Rc = Rc And UnLinkTable(DbMain, "LAU_VentasFacturaNac")
   
   Rc = Rc And UnLinkTable(DbMain, "LAU_Retenciones")
   Rc = Rc And UnLinkTable(DbMain, "LAU_VentasBoletas")
   Rc = Rc And UnLinkTable(DbMain, "LAU_VentasDevolucion")
   Rc = Rc And UnLinkTable(DbMain, "LAU_VentasFacturaExp")
   
   Rc = Rc And UnLinkTable(DbMain, "LAU_Sucursales")
   
   Rc = Rc And UnLinkTable(DbMain, "HR_FUTGrItems")    'FUT
   'Rc = Rc And UnLinkTable(DbMain, "HR_NContrib")      'Form 22
   
   If Rc = False Then
      UnLinkLau = -2
      Exit Function
   End If
#End If

End Function
