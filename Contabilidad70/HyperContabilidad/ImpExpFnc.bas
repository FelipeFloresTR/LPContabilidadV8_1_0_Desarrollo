Attribute VB_Name = "ImpExpFnc"
Option Explicit

Type RegComp_t
   IdCuenta As Long
   Debe As Double
   Haber As Double
   Descrip As String
   IdAreaNeg As Long
   IdCCosto As Long
   IdDoc As Long
End Type

Dim lRegComp() As RegComp_t

Dim lFNameLogImp As String

Public Function ImportComprobantes(Frm As FrmImpComp, ByVal fname As String, ByVal Ano As Integer, ByVal Mes As Integer, ByVal SoloRevisar As Boolean) As Integer
   Dim Buf As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TipoComp As String
   Dim IdTipoComp As Integer
   Dim ImpEnable As Boolean
   Dim i As Integer, j As Integer, l As Integer, k As Integer
   Dim p As Long, mov As Integer
   Dim NUpd As Long
   Dim NIns As Long
   Dim Rc As Integer
   Dim Fd As Long
   Dim Glosa As String
   Dim DtComp As Long
   Dim CampoInvalido As String
   Dim Row As Integer, r As Integer
   Dim ImpTipoComp(N_TIPOCOMP) As Integer
   Dim Descrip As String
   Dim Dt1 As Long, Dt2 As Long
   Dim Estado As String, IdEstado As Integer
   Dim CodCuenta As String, IdCuenta As Long
   Dim NombCta As String, DescCuenta As String, UltimoNivel As Boolean
   Dim Debe As Double, Haber As Double
   Dim CompErr As Boolean
   Dim TotComp As Double, TotDebe As Double, TotHaber As Double, AuxTotComp As Double
   Dim PrimerCampo As String
   Dim Aux As String
   Dim NewComp As Boolean
   Dim NCompsConError As Integer
   Dim LargoArchivo As Single, TotLeido As Single, Avance As Single
   Dim MsgTipoComp As String
   Dim TipoAjuste As Integer
   Dim CodANeg As String, IdAreaNeg As Long
   Dim CodCCosto As String, IdCCosto As Long
   Dim TipoLib As Integer, TipoDoc As Integer, NumDoc As String, DTE As Boolean
   Dim RutEnt As String, IdEnt As Long, IdDoc As Long
   Dim NombEnt As String, NotValidRut As Boolean
   Dim DiminutivoDoc As String
   
   ImportComprobantes = False
   
   TipoAjuste = TAJUSTE_AMBOS   'por omisión. Por ahora no se importa el tipo de ajuste en los comprobanes, sino que se asume como ambos. (21 ene 2014)
   
   
   If SoloRevisar Then
      Frm.Lb_Proceso = "Revisando el archivo...·"
   Else
      Frm.Lb_Proceso = "Importando comprobantes...·"
   End If
         
   Call FirstLastMonthDay(DateSerial(Ano, Mes, 1), Dt1, Dt2)
   
   lFNameLogImp = gImportPath & "\Log\ImpComprob-" & Format(Now, "yyyymmdd") & ".log"
   
   LargoArchivo = FileSize(fname)

         
   'abrimos el archivo
   Fd = FreeFile
   Open fname For Input As #Fd
   If ERR Then
      MsgErr fname
      ImportComprobantes = -ERR
      Exit Function
   End If
      
   r = 0
   
   For i = 0 To N_TIPOCOMP
      ImpTipoComp(i) = 0
   Next i
   
   ReDim lRegComp(1000)
      
   Do Until EOF(Fd)
         
      Line Input #Fd, Buf
      l = l + 1
      
      If gDbType = SQL_ACCESS Then
        If l > 30000 Then
           MsgBox1 "El archivo tiene demasiadas líneas." & vbCrLf & vbCrLf & "Proceso finalizado.", vbExclamation
           Exit Do
        End If
      End If
         
         
      'Debug.Print l & ")" & Buf
         
      TotLeido = TotLeido + Len(Buf)
      Avance = (TotLeido / LargoArchivo) * 100
      Frm.Pb_Proceso.Value = Int(Avance)
      
      Buf = Trim(Buf)
      
      '1er registro con nombres de campos
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 Then
         GoTo NextRec
      End If
      
      p = 1
      
      'Tipo Comprobante
      TipoComp = Trim(NextField2(Buf, p))
      
      If TipoComp <> "" Then    'encabezado => nuevo comprobante
      
         If r > 0 Then   'no es el primer comprobante => debemos cerrar el anterior y grabar
         
            'vemos si hay error en los totales y si el comprobante cuadra
            If TotDebe <> TotHaber Then
               Call AddLogImp(lFNameLogImp, fname, l, "Comprobante no cuadra.")
               CompErr = True
            ElseIf TotDebe <> TotComp Then
               Call AddLogImp(lFNameLogImp, fname, l, "Total comprobante no cuadra con totales Debe y Haber.")
               CompErr = True
            End If
            
            If CampoInvalido <> "" Or CompErr Then
               NCompsConError = NCompsConError + 1
            End If
            
            'si no hay errores y estamos grabando
            If NCompsConError = 0 And Not SoloRevisar Then
            
               Call SaveComp(IdTipoComp, DtComp, Glosa, IdEstado, TipoAjuste, TotComp)
               
            End If

            IdTipoComp = 0
           
         End If
      
         r = r + 1
      
         CampoInvalido = ""   'error en algún campo
         CompErr = False      'error en los totales
         
         NewComp = True
         mov = 0
         TotDebe = 0
         TotHaber = 0
         
         For k = 0 To UBound(lRegComp)
            lRegComp(k).IdCuenta = 0
            lRegComp(k).Debe = 0
            lRegComp(k).Haber = 0
            lRegComp(k).Descrip = ""
         Next k
            
         IdTipoComp = FindTipoComp(TipoComp)
         If IdTipoComp < 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Tipo de comprobante inválido. Valores permitidos. Egreso, Ingreso o Traspaso")
         Else
            ImpTipoComp(IdTipoComp) = ImpTipoComp(IdTipoComp) + 1
         End If
         
         
      Else
         NewComp = False
         
      End If
               
      'Fecha ingreso
      Aux = Trim(NextField2(Buf, p))
      If NewComp Then
         DtComp = ValFmtDate(Aux, False)
         If DtComp = 0 Or DtComp < Dt1 Or DtComp > Dt2 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Fecha ingreso inválida o fuera del último mes abierto. Formato permitido dd/mm/aaaa")
         End If
      End If
      
      'Total Comp
      AuxTotComp = Int(vFmt(Trim(NextField2(Buf, p))))
      If NewComp Then
         TotComp = AuxTotComp
         If TotComp < 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Total comprobante inválido.")
         End If
      End If
      
      
      'Estado Comprobante
      Estado = Trim(NextField2(Buf, p))
      If NewComp Then
         IdEstado = FindEstadoComp(Estado)
         If IdEstado <= 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Estado inválido. Valores permitidos: Aprobado, Pendiente o Anulado")
         End If
      End If
      
      'Glosa
      Aux = Trim(NextField2(Buf, p))
      If NewComp Then
         Glosa = Aux
         If Glosa = "" Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Falta glosa del comprobante.")
         End If
      End If
      
      'Cod. Cuenta
      CodCuenta = Trim(NextField2(Buf, p))
      CodCuenta = VFmtCodigoCta(CodCuenta)
      If CodCuenta = "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Falta código de cuenta.")
      Else
         NombCta = ""
         IdCuenta = GetIdCuenta(NombCta, CodCuenta, DescCuenta, UltimoNivel)
         If IdCuenta = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta inválido.")
         ElseIf Not UltimoNivel Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta no es de {ultimo nivel.")
         End If
      End If
      
      'Valores
      Debe = Int(vFmt(Trim(NextField2(Buf, p))))
      Haber = Int(vFmt(Trim(NextField2(Buf, p))))
      
      If Debe < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Debe inválido.")
      Else
         TotDebe = TotDebe + Debe
      End If
                     
      If Haber < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Haber inválido.")
      Else
         TotHaber = TotHaber + Haber
      End If
      
                                       
      'Descrip movimiento
      Descrip = Trim(NextField2(Buf, p))
      
      'Area de Negocio
      CodANeg = Trim(NextField2(Buf, p))
      'Centro de Costo
      CodCCosto = Trim(NextField2(Buf, p))
      
      IdAreaNeg = GetAreaNegocio(CodANeg)
      IdCCosto = GetCentroCosto(CodCCosto)
      
      If CodANeg <> "" And IdAreaNeg = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Código Área de Negocio inválido.")
      End If
      
      If CodCCosto <> "" And IdCCosto = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Código Centro de Gestión inválido.")
      End If
      
      TipoLib = 0
      TipoDoc = 0
      DTE = 0
      NumDoc = ""
      IdEnt = 0
      IdDoc = 0
      
      'Tipo Lib
      Aux = Trim(NextField2(Buf, p))
      
      Select Case Left(UCase(Aux), 1)
         '
         Case "C"
            TipoLib = LIB_COMPRAS
            
         Case "V"
            TipoLib = LIB_VENTAS
            
         Case "R"
            TipoLib = LIB_RETEN
            
         Case "O"
            TipoLib = LIB_OTROS
            
         Case "S"
            TipoLib = LIB_REMU
            
         Case "F"
            TipoLib = LIB_OTROFULL
            
         Case ""
            TipoLib = 0
            
         Case Else
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Tipo de libro inválido.")
            
      End Select
         
      Aux = Trim(NextField2(Buf, p))
      If TipoLib <> 0 Then
         TipoDoc = FindTipoDoc(TipoLib, Aux)
         If TipoDoc = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Tipo de documento inválido o no corresponde al libro de " & gTipoLib(TipoLib) & ".")
         End If
'      ElseIf Aux <> "" Then
'         CampoInvalido = CampoInvalido & "," & p
'         Call AddLogImp(lFNameLogImp, FName, l, "Falta ingresar tipo de libro.")
      End If
                         
      DiminutivoDoc = ""
      If TipoDoc <> 0 Then
         DiminutivoDoc = GetDiminutivoDoc(TipoLib, TipoDoc)
      End If
      
      If TipoLib = LIB_COMPRAS Or TipoLib = LIB_RETEN Then
         If DiminutivoDoc <> "NCC" And DiminutivoDoc <> "NCF" Then
            If IdTipoComp <> TC_EGRESO And IdTipoComp <> TC_TRASPASO Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, fname, l, "Tipo de libro-documento inválido.  No corresponde al tipo de comprobante ingresado.")
            End If
            
         ElseIf IdTipoComp <> TC_INGRESO And IdTipoComp <> TC_TRASPASO Then
         
            '2879479 se agrega condicion que solo de error de libro cuando monto este en debe
            If Not Debe = 0 And Haber > 0 Then
                CampoInvalido = CampoInvalido & "," & p
                Call AddLogImp(lFNameLogImp, fname, l, "Tipo de libro-documento inválido.  No corresponde al tipo de comprobante ingresado.")
            End If
            'fin 2879479
         End If
         
      ElseIf TipoLib = LIB_VENTAS Then
         If DiminutivoDoc <> "NCV" And DiminutivoDoc <> "NCE" And DiminutivoDoc <> "DVB" Then
            If IdTipoComp <> TC_INGRESO And IdTipoComp <> TC_TRASPASO Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, fname, l, "Tipo de libro-documento inválido.  No corresponde al tipo de comprobante ingresado.")
            End If
            
         ElseIf IdTipoComp <> TC_EGRESO And IdTipoComp <> TC_TRASPASO Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Tipo de libro-documento inválido.  No corresponde al tipo de comprobante ingresado.")
         
         End If
      End If
                           
      'DTE
      Aux = Trim(NextField2(Buf, p))
          
      DTE = IIf(Val(Aux) = 0 Or Trim(Aux) = "", 0, 1)
      
   
      'NumDoc
      NumDoc = Trim(NextField2(Buf, p))
      If TipoLib <> 0 And TipoDoc <> 0 And NumDoc = "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "N° de documento inválido.")
      End If
           
      'Entidad
      IdEnt = 0
      NotValidRut = False
      Aux = Trim(NextField2(Buf, p))
      If Aux = "0-0" Or Aux = "" Then
         RutEnt = ""
      Else
         RutEnt = vFmtCID(Aux)
         If RutEnt = "0" Or RutEnt = "" Then    'es inválido
            NotValidRut = True
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "RUT inválido")
         Else
            RutEnt = FmtCID(RutEnt)
            IdEnt = GetIdEntidad(RutEnt, NombEnt, NotValidRut)
            If IdEnt = 0 Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, fname, l, "RUT entidad no ha sido ingresado al sistema.")
            End If
         End If
      End If
              
      'validamos que exista el documento
      If TipoLib <> 0 And TipoDoc <> 0 And NumDoc <> "" And (TipoLib = LIB_VENTAS Or ((TipoLib = LIB_COMPRAS Or TipoLib = LIB_RETEN) And IdEnt <> 0) Or TipoLib = LIB_OTROS Or TipoLib = LIB_REMU Or TipoLib = LIB_OTROFULL) Then
      
         Q1 = "SELECT IdDoc FROM Documento "
         Q1 = Q1 & " WHERE TipoLib = " & TipoLib & " AND TipoDoc = " & TipoDoc
         
         'de acuerdo a lo solicitado por Claudio Villegas, no se valida la entidad en el caso de las ventas
         If TipoLib = LIB_COMPRAS Or TipoLib = LIB_RETEN Then
            Q1 = Q1 & " AND IdEntidad = " & IdEnt    'en rigor si la entidad no existe, tampoco existe el documento
         End If
         
         Q1 = Q1 & " AND NumDoc = '" & NumDoc & "'"
         
'      If gDbType = SQL_ACCESS Then
'        Q1 = Q1 & " AND iif(DTE <> 0, -1, 0) = " & Abs(DTE)
'      Else
        Q1 = Q1 & " AND iif(DTE <> 0, 1, 0) = " & Abs(DTE)
'      End If
         
       '  Q1 = Q1 & " AND iif(DTE <> 0, 1, 0) = " & Abs(DTE)
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         Set Rs = OpenRs(DbMain, Q1)
          
         If Rs.EOF Then       'documento no existe
            IdDoc = 0
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Documento no ha sido ingresado al sistema.")
         Else
            IdDoc = vFld(Rs("IdDoc"))
         End If
         
         Call CloseRs(Rs)
      ElseIf TipoLib <> 0 Or TipoDoc <> 0 Or NumDoc <> "" Or IdEnt <> 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Información del documento incompleta.")
         
      End If
      
      If CampoInvalido = "" Then    'registro válido
         lRegComp(mov).IdCuenta = IdCuenta
         lRegComp(mov).Debe = Debe
         lRegComp(mov).Haber = Haber
         lRegComp(mov).Descrip = Descrip
         lRegComp(mov).IdAreaNeg = IdAreaNeg
         lRegComp(mov).IdCCosto = IdCCosto
         lRegComp(mov).IdDoc = IdDoc
         mov = mov + 1
         
         If mov >= UBound(lRegComp) Then
            ReDim Preserve lRegComp(UBound(lRegComp) + 100)
         End If
            
      Else
         CompErr = True
      End If
            
NextRec:
   Loop

   Close #Fd
   
   'guardamos el último comprobante
   If IdTipoComp <> 0 Then     'hay uno por guardar
   
      If TotDebe <> TotHaber Then
         Call AddLogImp(lFNameLogImp, fname, l, "Comprobante no cuadra.")
         CompErr = True
      ElseIf TotDebe <> TotComp Then
         Call AddLogImp(lFNameLogImp, fname, l, "Total comprobante no cuadra con totales Debe y Haber.")
         CompErr = True
      End If
      
      If CampoInvalido <> "" Or CompErr Then
         NCompsConError = NCompsConError + 1
      End If
      
      'si no hay errores y estamos grabando
      If NCompsConError = 0 And Not SoloRevisar Then
      
         Call SaveComp(IdTipoComp, DtComp, Glosa, IdEstado, TipoAjuste, TotComp)
         
      End If
      
   End If
   
   Frm.Pb_Proceso.Value = 100
   
   MsgTipoComp = ""
         
   If NCompsConError = 0 Then
   
      MsgTipoComp = MsgTipoComp & vbCrLf & vbCrLf & "   - " & gTipoComp(TC_APERTURA) & ": " & ImpTipoComp(TC_APERTURA)
      
      For i = 1 To N_TIPOCOMP
         If i <> TC_APERTURA Then
            MsgTipoComp = MsgTipoComp & vbCrLf & vbCrLf & "   - " & gTipoComp(i) & ": " & ImpTipoComp(i)
         End If
      Next i
         
   
      If SoloRevisar Then
         If r > 1 Then
            MsgBox1 "Proceso de Revisión finalizado exitosamente." & vbNewLine & vbNewLine & "Resultado:" & vbNewLine & vbNewLine & "Se encontraron " & r & " comprobantes en el archivo.", vbInformation + vbOKOnly
            ImportComprobantes = True
         ElseIf r = 1 Then
            MsgBox1 "Proceso de Revisión finalizado exitosamente." & vbNewLine & vbNewLine & "Resultado:" & vbNewLine & vbNewLine & "Se encontró 1 comprobante en el archivo.", vbInformation + vbOKOnly
            ImportComprobantes = True
         Else   'r=0
            MsgBox1 "Proceso de Revisión finalizado." & vbNewLine & vbNewLine & "Resultado:" & vbNewLine & vbNewLine & "No se encontraron comprobantes en el archivo.", vbInformation + vbOKOnly
         End If
      Else
         If r > 1 Then
            MsgBox1 "Proceso de Importación finalizado exitosamente." & vbNewLine & vbNewLine & "Se importaron " & r & " comprobantes:" & MsgTipoComp, vbInformation + vbOKOnly
            ImportComprobantes = True
         ElseIf r = 1 Then
            MsgBox1 "Proceso de Importación finalizado exitosamente." & vbNewLine & vbNewLine & "Se importó 1 comprobante:" & MsgTipoComp, vbInformation + vbOKOnly
            ImportComprobantes = True
         Else
            MsgBox1 "Proceso de Importación finalizado." & vbNewLine & vbNewLine & "No se importaron comprobantes.", vbInformation + vbOKOnly
         End If
      End If
   
   Else
      If NCompsConError = 1 Then
         MsgBox1 "Se encontró 1 comprobante con errores en el archivo de importación." & vbNewLine & vbNewLine & "No se realizó el proceso de importación.", vbExclamation + vbOKOnly
      Else
         MsgBox1 "Se encontraron " & NCompsConError & " comprobantes con errores en el archivo de importación." & vbNewLine & vbNewLine & "No se realizó el proceso de importación.", vbExclamation + vbOKOnly
      End If
      
      If MsgBox1("¿Desea revisar el log de importación " & lFNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Frm.hWnd, "open", lFNameLogImp, "", "", SW_SHOW)
      End If
   End If

   Frm.Pb_Proceso.Value = 0
   Frm.Lb_Proceso = ""

End Function
Private Function FindTipoComp(ByVal TipoComp As String) As Integer
   Dim i As Integer
   
   If TipoComp = "" Then
      FindTipoComp = 0
   End If
   
   For i = 1 To UBound(gTipoComp)
      If StrComp(gTipoComp(i), TipoComp, vbTextCompare) = 0 Then
         FindTipoComp = i
         Exit Function
      End If
   Next i
   
   FindTipoComp = -1
   
End Function
Private Function FindEstadoComp(ByVal Estado As String) As Integer
   Dim i As Integer
   
   If Estado = "" Then
      FindEstadoComp = 0
   End If
   
   For i = 1 To UBound(gEstadoComp)
      If StrComp(gEstadoComp(i), Estado, vbTextCompare) = 0 Then
         FindEstadoComp = i
         Exit Function
      End If
   Next i
   
   FindEstadoComp = -1
   
End Function

Private Function SaveComp(ByVal IdTipoComp As Integer, ByVal FechaComp As Long, ByVal Glosa As String, ByVal Estado As Integer, ByVal TipoAjuste As Integer, Total As Double) As Boolean
   Dim Rs As Recordset
   Dim Rc As Long, RNum As Long
   Dim Q1 As String
   Dim sWhere As String, WhConWhere As String, WhConAnd As String
   Dim AddUniqueRecord As Boolean
   Dim MesActual As Integer
   Dim Correlativo As Long, idcomp As Long
   Dim WhTAjuste As String, WhEmp As String
   Dim FldArray(7) As AdvTbAddNew_t
   
   SaveComp = True
   
'   IdComp = AdvTbAddNew(DbMain, "Comprobante", "IdComp", "Tipo", IdTipoComp)
'
'   If IdComp = 0 Then
'      SaveComp = False
'      Exit Function
'   End If
'
'   Q1 = "UPDATE Comprobante SET "
'   Q1 = Q1 & "  Fecha = " & FechaComp
'   Q1 = Q1 & ", Tipo = " & IdTipoComp
'   Q1 = Q1 & ", IdUsuario = " & gUsuario.IdUsuario
'   Q1 = Q1 & ", FechaCreacion = " & CLng(Int(Now))
'   Q1 = Q1 & ", TipoAjuste = " & TipoAjuste
'   Q1 = Q1 & ", Correlativo = -1"
'   Q1 = Q1 & ", IdEmpresa = " & gEmpresa.id
'   Q1 = Q1 & ", Ano = " & gEmpresa.Ano
'   Q1 = Q1 & " WHERE IdComp = " & IdComp
'
'   Call ExecSQL(DbMain, Q1)
    
   FldArray(0).FldName = "IdUsuario"
   FldArray(0).FldValue = gUsuario.IdUsuario
   FldArray(0).FldIsNum = True
   
   FldArray(1).FldName = "FechaCreacion"
   FldArray(1).FldValue = CLng(Int(Now))
   FldArray(1).FldIsNum = True
         
   FldArray(2).FldName = "IdEmpresa"
   FldArray(2).FldValue = gEmpresa.id
   FldArray(2).FldIsNum = True
               
   FldArray(3).FldName = "Ano"
   FldArray(3).FldValue = gEmpresa.Ano
   FldArray(3).FldIsNum = True
   
   FldArray(4).FldName = "Fecha"
   FldArray(4).FldValue = FechaComp
   FldArray(4).FldIsNum = True
         
   FldArray(5).FldName = "TipoAjuste"
   FldArray(5).FldValue = TipoAjuste
   FldArray(5).FldIsNum = True
         
   FldArray(6).FldName = "Tipo"
   FldArray(6).FldValue = IdTipoComp
   FldArray(6).FldIsNum = True
         
   FldArray(7).FldName = "Correlativo"
   FldArray(7).FldValue = -1
   FldArray(7).FldIsNum = True
      
   idcomp = AdvTbAddNewMult(DbMain, "Comprobante", "IdComp", FldArray)
   
   '3376884
    Call SeguimientoComprobantes(idcomp, gEmpresa.id, gEmpresa.Ano, "ImpExpFnc.SaveComp", "", 1, "", gUsuario.IdUsuario, 1, 1)
    'fin 3376884
    
   If idcomp = 0 Then
      SaveComp = False
      Exit Function
   End If

    
   Correlativo = 0
   
   MesActual = month(FechaComp)
   
   If gTipoCorrComp = TCC_UNICO Then
               
      If gPerCorrComp = TCC_MENSUAL Then   'si es anual o continuo sWhere = ""
         sWhere = SqlMonthLng("Fecha") & " = " & MesActual
      End If
      
   ElseIf gTipoCorrComp = TCC_TIPOCOMP Then
      sWhere = " Tipo = " & IdTipoComp
      
      If gPerCorrComp = TCC_MENSUAL Then
         sWhere = sWhere & " AND " & SqlMonthLng("Fecha") & " = " & MesActual
      End If
      
   End If
   
   'agregamos el tipo de ajuste
   If TipoAjuste = TAJUSTE_TRIBUTARIO Then
      WhTAjuste = " TipoAjuste = " & TAJUSTE_TRIBUTARIO
   Else
      WhTAjuste = " TipoAjuste IN ( " & TAJUSTE_FINANCIERO & "," & TAJUSTE_AMBOS & ")"
   End If
   
   If sWhere <> "" Then
      sWhere = sWhere & " AND " & WhTAjuste
   Else
      sWhere = WhTAjuste
   End If
   
   If sWhere <> "" Then
      WhConWhere = " WHERE " & sWhere & " AND Correlativo > 0"
      WhConAnd = " AND " & sWhere  ' sin > 0
   Else
      WhConWhere = " WHERE Correlativo > 0"
      
   End If
   
   WhEmp = " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Do
      Q1 = "SELECT Max(Correlativo) as N FROM Comprobante " & WhConWhere & WhEmp
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         Correlativo = vFld(Rs("N")) + 1
      Else
         Correlativo = 1
      End If
      Call CloseRs(Rs)
               
      Q1 = "UPDATE Comprobante SET Correlativo=" & Correlativo
      Q1 = Q1 & " WHERE IdComp=" & idcomp & WhEmp
      Rc = ExecSQL(DbMain, Q1)
      
      '3376884
      Call SeguimientoComprobantes(0, gEmpresa.id, gEmpresa.Ano, "ImpExpFnc.SaveComp1", "", 1, "WHERE IdComp=" & idcomp & WhEmp, gUsuario.IdUsuario, 1, 2)
      'fin 3376884
   
      DoEvents    'produce cosas raras
      
      Q1 = "SELECT Correlativo, idComp FROM Comprobante "
      Q1 = Q1 & " WHERE Correlativo = " & Correlativo & WhConAnd & WhEmp
      
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         AddUniqueRecord = True
      
         Do Until Rs.EOF
            If vFld(Rs("idComp")) <> idcomp Then
               AddUniqueRecord = False
               Exit Do
            End If
            Rs.MoveNext
         Loop
         Call CloseRs(Rs)
         
         If AddUniqueRecord = False Then
            Q1 = "UPDATE Comprobante SET Correlativo=-1"
            Q1 = Q1 & " WHERE IdComp=" & idcomp & WhEmp
            Rc = ExecSQL(DbMain, Q1)
            
            '3376884
            Call SeguimientoComprobantes(0, gEmpresa.id, gEmpresa.Ano, "FrmConfigCorrComp.Bt_MarcarRes_Click", "", 1, " WHERE IdComp=" & idcomp & WhEmp, gUsuario.IdUsuario, 1, 2)
            'fin 3376884
            
         Else
            Exit Do ' tenemos el correlativo
         End If
         
      Else
         Call CloseRs(Rs)
      End If
                     
   Loop
               
   
   If idcomp <> 0 Then
   
      'actualizamos el encabezado
      
      Q1 = "UPDATE Comprobante SET "
      Q1 = Q1 & "  Fecha = " & FechaComp
      Q1 = Q1 & ", Tipo = " & IdTipoComp
      Q1 = Q1 & ", Estado = " & Estado
      Q1 = Q1 & ", Glosa = '" & ParaSQL(Left(Glosa, 100)) & "'"
      Q1 = Q1 & ", ImpResumido = 0"
      Q1 = Q1 & ", TotalDebe = " & Total
      Q1 = Q1 & ", TotalHaber = " & Total
      Q1 = Q1 & ", FechaImport = " & Int(DbMainDate)     'SqlNow(DbMain)
      Q1 = Q1 & "  WHERE IdComp = " & idcomp & WhEmp
      
      Call ExecSQL(DbMain, Q1, False)
      
      '3376884
      Call SeguimientoComprobantes(0, gEmpresa.id, gEmpresa.Ano, "FrmConfigCorrComp.Bt_MarcarRes_Click", "", 1, "WHERE IdComp = " & idcomp & WhEmp, gUsuario.IdUsuario, 1, 2)
      'fin 3376884
      
      Call AddLogComprobantes(idcomp, gUsuario.IdUsuario, O_IMPORT, Now, Estado, Correlativo, FechaComp, IdTipoComp, Estado, TipoAjuste)
      
   End If
      
   Call SaveMovsComp(idcomp, Total)
      
End Function
Private Sub SaveMovsComp(ByVal idcomp As Long, ByVal Total As Double)
   Dim i As Integer
   Dim Lin As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim idMov As Long
   Dim StrIdDoc As String
   Dim StrIdDocSel As String
   Dim SumDebe As Double
   Dim SumHaber As Double
   Dim FldArray(2) As AdvTbAddNew_t
   
   SumDebe = Total
   SumHaber = Total


   If idcomp <= 0 Then
      Exit Sub
   End If

   Lin = 1
   For i = 0 To UBound(lRegComp)
            
      If lRegComp(i).IdCuenta = 0 Then     'ya terminó la lista de mov.
         Exit For
      End If
            
      FldArray(0).FldName = "IdComp"
      FldArray(0).FldValue = idcomp
      FldArray(0).FldIsNum = True
            
      FldArray(1).FldName = "IdEmpresa"
      FldArray(1).FldValue = gEmpresa.id
      FldArray(1).FldIsNum = True
                  
      FldArray(2).FldName = "Ano"
      FldArray(2).FldValue = gEmpresa.Ano
      FldArray(2).FldIsNum = True
      
           
      idMov = AdvTbAddNewMult(DbMain, "MovComprobante", "IdMov", FldArray)
                  
      Q1 = "UPDATE MovComprobante SET "
      Q1 = Q1 & "  Orden = " & Lin
      Q1 = Q1 & ", IdCuenta = " & lRegComp(i).IdCuenta
      Q1 = Q1 & ", Debe = " & lRegComp(i).Debe
      Q1 = Q1 & ", Haber = " & lRegComp(i).Haber
      Q1 = Q1 & ", Glosa = '" & ParaSQL(Left(lRegComp(i).Descrip, 50)) & "'"
      Q1 = Q1 & ", IdAreaNeg = " & lRegComp(i).IdAreaNeg
      Q1 = Q1 & ", IdCCosto = " & lRegComp(i).IdCCosto
      Q1 = Q1 & ", DePago = 0"
      Q1 = Q1 & ", IdDoc = " & lRegComp(i).IdDoc
      Q1 = Q1 & ", Nota = ' '"
      
      Q1 = Q1 & " WHERE IdMov = " & idMov
      Q1 = Q1 & " AND IdComp = " & idcomp
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
      
      '3376884
      Call SeguimientoMovComprobante(0, gEmpresa.id, gEmpresa.Ano, "FrmConfigCorrComp.Bt_MarcarRes_Click", Q1, 1, " WHERE IdMov = " & idMov & " AND IdComp = " & idcomp & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano, 1, 2)
      'fin 3376884
            
      Lin = Lin + 1
      
   Next i

End Sub


Public Function ImportActFijoFile(Frm As FrmImpActFijoFile, ByVal fname As String, ByVal SoloRevisar As Boolean) As Integer
   Dim Buf As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim ImpEnable As Boolean
   Dim i As Integer, j As Integer, l As Integer, k As Integer
   Dim p As Long, mov As Integer
   Dim NUpd As Long
   Dim NIns As Long
   Dim Rc As Integer
   Dim Fd As Long
   Dim CampoInvalido As String
   Dim Row As Integer, r As Integer
   Dim Descrip As String
   Dim Dt1 As Long, Dt2 As Long
   Dim PrimerCampo As String
   Dim Aux As String, AuxVal As Double
   Dim Neto As Double, IVA As Double
   Dim NewAF As Boolean
   Dim NAFConError As Integer
   Dim LargoArchivo As Single, TotLeido As Single, Avance As Single
   Dim DtCompra As Long, DtUtilizacion As Long
   Dim AFTotDepreciado As Boolean, AFNoDepreciable As Boolean
   Dim Cred33bis As Integer
   Dim Cantidad As Long
   Dim EsDepreciable As Boolean
   Dim MesesDepNormal As Long, MesesDepAcelerada As Long, MesesDepInstantanea As Long, MesesDepDecimaParte As Long, MesesDepDecimaParte2 As Long
   Dim MesesDepNormalHist As Long, MesesDepAceleradaHist As Long, MesesDepInstantaneaHist As Long, MesesDepDecimaparteHist As Long, MesesDepDecimaparte2Hist As Long
   Dim TipoDep As Integer, TipoDepHist As Integer
   Dim DepHistorica As Double
   Dim VentaBaja As Integer, NetoVenta As Double, IVAVenta As Double, DtVentaBaja As Long
   Dim DtIncorporacion As Long, DtDisponible As Long
   Dim DerechosInternacion As Double, Transporte As Double, Adaptacion As Double, OtrosAdquisicion As Double
   Dim IVARec As Double, FormacionPersonal As Double, Reubicacion As Double, OtrosGastos As Double
   Dim PrecioAdquisicion As Double, TotalGastos As Double
   Dim IdGrupo As Long, IdCuenta As Long, NomCuenta As String, CodCuenta As String, DescCuenta As String, UltNivel As Boolean
   Dim ValCredito As Double, VidaUtil As Integer
   Dim IdActFijo As Long, IdFicha As Long
   Dim AuxDep As Integer
   Dim FldArray(3) As AdvTbAddNew_t
   Dim Ley21210DepInsteInmed As Boolean
   Dim Ley21210DepAraucania As Boolean
   Dim Ley21256DepInsteInmed As Boolean
   Dim NDepSel As Integer
   Dim PatenteRol As String, NombreProy As String, FechaProy As Long
   
   '2861733 tema 2
   Dim vCodAreaNegocio As Long, vCodCentroGestion As Long
   Dim vIdAreaNegocio As Long, vIdCentroGestion As Long
   '2861733 tema 2
   
   ImportActFijoFile = False

   If SoloRevisar Then
      Frm.Lb_Proceso = "Revisando el archivo...·"
   Else
      Frm.Lb_Proceso = "Importando Activos Fijos...·"
   End If

   lFNameLogImp = gImportPath & "\Log\ImpActFijo-" & Format(Now, "yyyymmdd") & ".log"

   LargoArchivo = FileSize(fname)


   'abrimos el archivo
   Fd = FreeFile
   Open fname For Input As #Fd
   If ERR Then
      MsgErr fname
      ImportActFijoFile = -ERR
      Exit Function
   End If

   r = 0

   Do Until EOF(Fd)

      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf

      TotLeido = TotLeido + Len(Buf)
      Avance = (TotLeido / LargoArchivo) * 100
      Frm.Pb_Proceso.Value = Int(Avance)

      Aux = ReplaceStr(Buf, vbTab, "")
      Aux = Trim(Aux)
      If Aux = "" Then
         Buf = Aux
      End If
      Buf = Trim(Buf)

      '1er registro con nombres de campos
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 Then
         GoTo NextRec
      End If

      p = 1

      CampoInvalido = ""
      
      'AF Totalmente Depreciado
      Aux = Left(Trim(NextField2(Buf, p)), 1)
      If Aux = "" Then
         AFTotDepreciado = VAL_NO
      Else
         AFTotDepreciado = ValSiNo(Aux)
      End If
      If Abs(AFTotDepreciado) <> VAL_SI And Abs(AFTotDepreciado) <> VAL_NO Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Campo AF Totalmente Depreciado inválido. Valores permitidos S/N/blanco")
      End If

      'AF no depreciable
      Aux = Left(Trim(NextField2(Buf, p)), 1)
      If Aux = "" Then
         AFNoDepreciable = VAL_NO
      Else
         AFNoDepreciable = ValSiNo(Aux)
      End If
      If Abs(AFNoDepreciable) <> VAL_SI And Abs(AFNoDepreciable) <> VAL_NO Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Campo AF no Depreciable inválido. Valores permitidos S/N/blanco")
      End If

      EsDepreciable = False
      If Not (AFTotDepreciado Or AFNoDepreciable) Then
         EsDepreciable = True
      End If
      
      EsDepreciable = Not (AFTotDepreciado Or AFNoDepreciable)

      'Fecha compra
      Aux = Trim(NextField2(Buf, p))
      DtCompra = ValFmtDate(Aux, False)
      If DtCompra = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Fecha compra inválida. Formato permitido dd/mm/aaaa")
      End If

      'Fecha utilización
      Aux = Trim(NextField2(Buf, p))
      DtUtilizacion = ValFmtDate(Aux, False)
      If EsDepreciable Then
         If DtUtilizacion = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Fecha utilización inválida. Formato permitido dd/mm/aaaa")

         ElseIf DtUtilizacion < DtCompra Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "La fecha de utilización debe ser posterior a la fecha de compra.")
         End If

      End If

      'Cantidad
      Cantidad = Int(vFmt(Trim(NextField2(Buf, p))))
      If Cantidad <= 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Falta Cantidad del Activo Fijo.")
      End If


      'Descripción
      Descrip = Trim(NextField2(Buf, p))
      If Descrip = "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Falta Descripción del Activo Fijo.")
      End If


      'Neto e IVA
      Neto = Int(vFmt(Trim(NextField2(Buf, p))))
      IVA = Int(vFmt(Trim(NextField2(Buf, p))))

      If Neto < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Neto inválido.")
      End If

      If IVA < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de IVA inválido.")
      End If

      If IVA <> 0 And Abs(Neto * gIVA - IVA) > 2 Then    'IVA puede ser cero de acuerdo a lo indicado por Victor 15/9/2016
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de IVA no calza con el valor neto.")
      End If


      'Cred. 33 bis
      Aux = Trim(NextField2(Buf, p))
      If Aux = "" Then
         Cred33bis = VAL_NO
      Else
         Cred33bis = ValSiNo(Aux)
      End If
      If Cred33bis <> VAL_SI And Cred33bis <> VAL_NO Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Campo Cred. 33 bis inválido. Valores permitidos S/N/blanco")
      End If


      'Valor Crédito
      Aux = Trim(NextField2(Buf, p))
      ValCredito = Int(vFmt(Aux))
      If Aux = "" Then
         ValCredito = -1    'para indicar que el campo viene en blanco
      ElseIf ValCredito < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Crédito inválido.")
      End If

      'Vida Util
      VidaUtil = Int(vFmt(Trim(NextField2(Buf, p))))
      If EsDepreciable And VidaUtil <= 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar la vida útil del Activo Fijo")
      End If

      'Meses Dep Normal  y Meses Dep Acelerada
      MesesDepNormal = Int(vFmt(Trim(NextField2(Buf, p))))
      MesesDepAcelerada = Int(vFmt(Trim(NextField2(Buf, p))))
      MesesDepInstantanea = Int(vFmt(Trim(NextField2(Buf, p))))
      MesesDepDecimaParte = Int(vFmt(Trim(NextField2(Buf, p))))
      MesesDepDecimaParte2 = Int(vFmt(Trim(NextField2(Buf, p))))
     
      NDepSel = 0
      If MesesDepNormal > 0 Then
         NDepSel = NDepSel + 1
         TipoDep = DEP_NORMAL
      End If
      If MesesDepAcelerada > 0 Then
         NDepSel = NDepSel + 1
         TipoDep = DEP_ACELERADA
      End If
      If MesesDepInstantanea > 0 Then
         NDepSel = NDepSel + 1
         TipoDep = DEP_INSTANTANEA
      End If
      If MesesDepDecimaParte > 0 Then
         NDepSel = NDepSel + 1
         TipoDep = DEP_DECIMAPARTE
      End If
      If MesesDepDecimaParte2 > 0 Then
         NDepSel = NDepSel + 1
         TipoDep = DEP_DECIMAPARTE2
      End If
         
         
      'Dep. Ley 21210, Inst. e Inmed o Araucanía
      If (MesesDepInstantanea > 0 Or MesesDepDecimaParte > 0) And DtUtilizacion < gFechaInicioDepInstantanea Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "La depreciación Instantánea o Décima Parte sólo se puede aplicar" & vbCrLf & "para activos fijos cuya fecha de utilización es posterior al " & FmtDate(gFechaInicioDepInstantanea, "dd mmm yyyy"))
      End If
      
      'Dep. Ley 21256, Inst. e Inmed
      If (MesesDepInstantanea > 0 Or MesesDepDecimaParte > 0) And DtUtilizacion < gFechaInicioDepInstantanea Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "La depreciación Instantánea e Inmediata sólo se puede aplicar" & vbCrLf & "para activos fijos cuya fecha de utilización es posterior al " & FmtDate(gFechaInicioDepInstantanea, "dd mmm yyyy"))
      End If
      
      If (MesesDepDecimaParte2 > 0) And DtCompra < gFechaInicioDepDecimaParte2 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "La depreciación Décima Parte MT sólo se puede aplicar" & vbCrLf & "para activos fijos cuya fecha de compra es posterior al " & FmtDate(gFechaInicioDepDecimaParte2, "dd mmm yyyy"))
      End If
      
      Aux = Left(Trim(NextField2(Buf, p)), 1)
      If Aux = "" Then
         Ley21210DepInsteInmed = VAL_NO
      Else
         Ley21210DepInsteInmed = ValSiNo(Aux)
      End If
      If Abs(Ley21210DepInsteInmed) <> VAL_SI And Abs(Ley21210DepInsteInmed) <> VAL_NO Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Campo Ley 21.210 - Dep. Inst. e Inmed. inválido. Valores permitidos S/N/blanco")
      End If

      Aux = Left(Trim(NextField2(Buf, p)), 1)
      If Aux = "" Then
         Ley21210DepAraucania = VAL_NO
      Else
         Ley21210DepAraucania = ValSiNo(Aux)
      End If
      If Abs(Ley21210DepAraucania) <> VAL_SI And Abs(Ley21210DepAraucania) <> VAL_NO Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Campo Ley 21.210 - Dep. Araucanía inválido. Valores permitidos S/N/blanco")
      End If
      
      If Ley21210DepInsteInmed And Ley21210DepAraucania Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "No es posible aplicar Depreciación Intantánea e Inmediata junto con Depreciación de la Araucanía. Debe elegir una de las dos si aplica.")
      End If

      If Ley21210DepAraucania And gEmpresa.Region <> 9 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "La Depreciación de la Araucanía sólo se puede aplicar a empresas cuya dirección corresponde a la región de la Araucanía.")
      End If

      If (Ley21210DepInsteInmed Or Ley21210DepAraucania) And (DtCompra < gFechaInicioDepLey21210 Or DtCompra > gFechaTerminoDepLey21210) Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "La depreciación Ley 21.210 sólo se puede aplicar para activos fijos cuya fecha de compra es está entre el " & vbCrLf & vbCrLf & FmtDate(gFechaInicioDepLey21210, "dd mmm yyyy") & " y el " & FmtDate(gFechaTerminoDepLey21210, "dd mmm yyyy"))
      End If
      
      Aux = Left(Trim(NextField2(Buf, p)), 1)
      If Aux = "" Then
         Ley21256DepInsteInmed = VAL_NO
      Else
         Ley21256DepInsteInmed = ValSiNo(Aux)
      End If
      If Abs(Ley21256DepInsteInmed) <> VAL_SI And Abs(Ley21256DepInsteInmed) <> VAL_NO Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Campo Ley 21.256 - Dep. Inst. e Inmed. inválido. Valores permitidos S/N/blanco")
      End If
      
      If (Ley21210DepInsteInmed Or Ley21210DepAraucania) And Ley21256DepInsteInmed Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "No es posible aplicar Depreciación Ley 21210 junto con Depreciación Ley 21.256. Debe elegir una de las dos si aplica.")
      End If
      
      If Ley21256DepInsteInmed And (DtCompra < gFechaInicioDepLey21256 Or DtCompra > gFechaTerminoDepLey21256) Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "La depreciación Ley 21.256 sólo se puede aplicar para activos fijos cuya fecha de compra es está entre el " & vbCrLf & vbCrLf & FmtDate(gFechaInicioDepLey21210, "dd mmm yyyy") & " y el " & FmtDate(gFechaTerminoDepLey21210, "dd mmm yyyy"))
      End If
      
      If EsDepreciable Then
      
         If NDepSel = 0 And Not Ley21210DepAraucania And Not Ley21256DepInsteInmed Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar los meses a depreciar este año (normal, acelerada, instantánea o décima parte)")
         ElseIf NDepSel > 1 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Debe ingresar meses a depreciar sólo en uno de los dos campos: Meses Dep. Normal, Meses Dep. Acelerada, Meses Dep. Instantánea o Meses Dep. Décima Parte")
         End If

      ElseIf NDepSel > 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Si el Activo Fijo no es depreciable o está totalmente depreciado no aplica campo Meses Dep. Normal o Meses Dep. Acelerada")

      End If
      
      'Meses Dep Normal Histórica  y Meses Dep Acelerada Histórica
      MesesDepNormalHist = Int(vFmt(Trim(NextField2(Buf, p))))
      MesesDepAceleradaHist = Int(vFmt(Trim(NextField2(Buf, p))))
      MesesDepInstantaneaHist = Int(vFmt(Trim(NextField2(Buf, p))))
      MesesDepDecimaparteHist = Int(vFmt(Trim(NextField2(Buf, p))))


      AuxDep = 0
      If MesesDepNormalHist > 0 Then
         AuxDep = AuxDep + 1
         TipoDepHist = DEP_NORMAL
      End If
      If MesesDepAceleradaHist > 0 Then
         AuxDep = AuxDep + 1
         TipoDepHist = DEP_ACELERADA
      End If
      If MesesDepInstantaneaHist > 0 Then
         AuxDep = AuxDep + 1
         TipoDepHist = DEP_INSTANTANEA
      End If
      If MesesDepDecimaparteHist > 0 Then
         AuxDep = AuxDep + 1
         TipoDepHist = DEP_DECIMAPARTE
      End If

      If MesesDepDecimaparte2Hist > 0 Then
         AuxDep = AuxDep + 1
         TipoDepHist = DEP_DECIMAPARTE2
      End If

      If EsDepreciable Then

         If DtUtilizacion < DateSerial(gEmpresa.Ano - 1, 12, 31) Then

            If AuxDep = 0 Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar los meses ya depreciados (dep. normal, acelerada, instantánea, deécima parte o décima parte MT histórica)")
            ElseIf AuxDep > 1 Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, fname, l, "Debe ingresar meses depreciados sólo en uno de los dos campos: Meses Dep. Normal Hist., Meses Dep. Acelerada Hist., Meses Dep. Instantánea Hist., Meses Dep. Décima Parte Hist. o Meses Dep. Décima Parte MT Hist.")
            End If

         ElseIf AuxDep > 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Si el Activo Fijo tiene fecha de utilización de este año, no corresponde depreciación histórica.")

         End If

      ElseIf AuxDep > 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Si el Activo Fijo no es depreciable o está totalmente depreciado no aplica Depreciación Histórica")

      End If

      'Dep. Histórica
      DepHistorica = Int(vFmt(Trim(NextField2(Buf, p))))
      If EsDepreciable Then

         If DtUtilizacion < DateSerial(gEmpresa.Ano - 1, 12, 31) Then
            If DepHistorica <= 0 Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar la depreciación acumulada histórica")
            End If
         ElseIf DepHistorica > 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "No corresponde ingresar depreciación acumulada histórica")
         End If

      End If

      'Venta o Baja
      Aux = UCase(Trim(NextField2(Buf, p)))
      Select Case Aux
         Case "V"
            VentaBaja = MOVAF_VENTA
         Case "B"
            VentaBaja = MOVAF_BAJA
         Case ""
            VentaBaja = MOVAF_COMPRA
         Case Else
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Campo Venta o Baja inválido, Valores permitidos: 'V', 'B' o blanco")
      End Select

      'Fecha venta o baja
      Aux = Trim(NextField2(Buf, p))
      DtVentaBaja = ValFmtDate(Aux, False)
      
      If Aux = "" Then
         If VentaBaja <> MOVAF_COMPRA Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Falta indicar la fecha de venta o baja. Formato permitido dd/mm/aaaa")
         End If
         
      ElseIf DtVentaBaja = 0 Then
         If VentaBaja <> MOVAF_COMPRA Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Fecha venta o baja inválida. Formato permitido dd/mm/aaaa")
         End If
         
      ElseIf DtVentaBaja < DtCompra Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Fecha venta o baja inválida. Debe ser posterior a la fecha de compra.")
      End If
      
      'Neto Venta
      NetoVenta = Int(vFmt(Trim(NextField2(Buf, p))))
      If NetoVenta < 0 Or (NetoVenta > 0 And VentaBaja <> MOVAF_VENTA) Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de venta inválido.")
      End If

      'IVA Venta
      IVAVenta = Int(vFmt(Trim(NextField2(Buf, p))))
      If IVAVenta < 0 Or (IVAVenta > 0 And VentaBaja <> MOVAF_VENTA) Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de IVA venta inválido.")
      End If

      'Cuenta contable
      CodCuenta = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      
      NomCuenta = ""
      
      If CodCuenta <> "" Then
         IdCuenta = GetIdCuenta(NomCuenta, CodCuenta, DescCuenta, UltNivel)
         If IdCuenta <= 0 Or Not UltNivel Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta inválido")
         ElseIf GetAtribCuenta(IdCuenta, ATRIB_ACTIVOFIJO) = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Cuenta no tiene atributo Activo Fijo")
         End If
      Else
         IdCuenta = 0
         DescCuenta = ""
      End If
      
      'PatenteRol, NombreProy, FechaProy
      PatenteRol = Left(Trim(NextField2(Buf, p)), 30)
      
'      If PatenteRol = "" Then
'         CampoInvalido = CampoInvalido & "," & p
'         Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar Patente, Rol o Inscripción según proceda.")
'      End If
      
      NombreProy = Left(Trim(NextField2(Buf, p)), 60)

      Aux = Trim(NextField2(Buf, p))
      FechaProy = ValFmtDate(Aux, False)
'      If FechaProy = 0 And Aux <> "" Then
'         CampoInvalido = CampoInvalido & "," & p
'         Call AddLogImp(lFNameLogImp, fname, l, "Fecha proyecto inválida. Formato permitido dd/mm/aaaa")
'      End If
      
      'Grupo
      Aux = Trim(NextField2(Buf, p))
      IdGrupo = GetIdGrupo(Aux)
      
      If Aux <> "" And IdGrupo = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Nombre Grupo inválido.")
      End If

      'Fecha incorporación
      Aux = Trim(NextField2(Buf, p))
      If Aux <> "" Then
         DtIncorporacion = ValFmtDate(Aux, False)
         If DtIncorporacion = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Fecha Incorporación inválida. Formato permitido dd/mm/aaaa")
         End If
      End If

      'Fecha disponible
      Aux = Trim(NextField2(Buf, p))
      
      If Aux <> "" Then
         DtDisponible = ValFmtDate(Aux, False)
         If DtDisponible = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Fecha Disponible inválida. Formato permitido dd/mm/aaaa")
         End If
      End If

      'Derechos Internacion
      DerechosInternacion = Int(vFmt(Trim(NextField2(Buf, p))))
      If DerechosInternacion < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Derechos de Internación inválido.")
      End If

      'Transporte
      Transporte = Int(vFmt(Trim(NextField2(Buf, p))))
      If Transporte < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Transporte inválido.")
      End If

      'Adaptación
      Adaptacion = Int(vFmt(Trim(NextField2(Buf, p))))
      If Adaptacion < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Adaptación inválido.")
      End If

      'Otros Adquisición
      OtrosAdquisicion = Int(vFmt(Trim(NextField2(Buf, p))))
      If OtrosAdquisicion < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Otros Adquisición inválido.")
      End If

      'IVA Rec
      IVARec = Int(vFmt(Trim(NextField2(Buf, p))))
      If IVARec < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de IVA Recuperable inválido.")
      End If

      'Formación Personal
      FormacionPersonal = Int(vFmt(Trim(NextField2(Buf, p))))
      If FormacionPersonal < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Formación Personal inválido.")
      End If

      'Reubicacion
      Reubicacion = Int(vFmt(Trim(NextField2(Buf, p))))
      If Reubicacion < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Reubicación inválido.")
      End If

      'Otros gastos no reconocidos
      OtrosGastos = Int(vFmt(Trim(NextField2(Buf, p))))
      If OtrosGastos < 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor de Otros Gastos no Reconocidos inválido.")
      End If
      
      '2861733 tema 2
       'Codigo Area de negocio
      vCodAreaNegocio = Int(vFmt(Trim(NextField2(Buf, p))))
      '3371014
      'If vCodAreaNegocio <> "" Then
      If vCodAreaNegocio <> 0 Then
      '3371014
         vIdAreaNegocio = GetAreaNegocio(vCodAreaNegocio)
         If vIdAreaNegocio <= 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de Area de Negocio Invalido")
         End If
      Else
         vIdAreaNegocio = 0
     End If
     
'       'Codigo Centro de Gestion
     vCodCentroGestion = Int(vFmt(Trim(NextField2(Buf, p))))
     '3371014
      'If vCodCentroGestion <> "" Then
      If vCodCentroGestion <> 0 Then
     '3371014
         vIdCentroGestion = GetCentroCosto(vCodCentroGestion)
         If vIdCentroGestion <= 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de Centro de Gestion Invalido")
         End If
      Else
         vIdCentroGestion = 0
     End If
     '2861733 tema 2

      If CampoInvalido = "" Then    'registro válido
      
         If SoloRevisar = False Then
            
            'grabarAF

            FldArray(0).FldName = "TipoMovAF"
            FldArray(0).FldValue = VentaBaja
            FldArray(0).FldIsNum = True
            
            FldArray(1).FldName = "Fecha"
            FldArray(1).FldValue = DtCompra
            FldArray(1).FldIsNum = True
                  
            FldArray(2).FldName = "IdEmpresa"
            FldArray(2).FldValue = gEmpresa.id
            FldArray(2).FldIsNum = True
                        
            FldArray(3).FldName = "Ano"
            FldArray(3).FldValue = gEmpresa.Ano
            FldArray(3).FldIsNum = True
      
            IdActFijo = AdvTbAddNewMult(DbMain, "MovActivoFijo", "IdActFijo", FldArray)

            If IdActFijo <> 0 Then

               Q1 = "UPDATE MovActivoFijo SET "
               Q1 = Q1 & "  FechaUtilizacion = " & DtUtilizacion
               Q1 = Q1 & ", FechaVentaBaja = " & DtVentaBaja
               Q1 = Q1 & ", Cantidad = " & Cantidad
               Q1 = Q1 & ", Descrip = '" & ParaSQL(Left(Descrip, 80)) & "'"
               Q1 = Q1 & ", Neto = " & Neto
               Q1 = Q1 & ", VidaUtil = " & VidaUtil
               Q1 = Q1 & ", IVA = " & IVA
               Q1 = Q1 & ", NetoVenta = " & NetoVenta
               Q1 = Q1 & ", IVAVenta = " & IVAVenta
               Q1 = Q1 & ", Cred4Porc = " & Cred33bis
               Q1 = Q1 & ", ValCred33 = " & ValCredito
               Q1 = Q1 & ", NoDepreciable = " & IIf(AFNoDepreciable, -1, 0)
               Q1 = Q1 & ", TotalmenteDepreciado = " & IIf(AFTotDepreciado, -1, 0)
               Q1 = Q1 & ", TipoDepLey21210 = " & IIf(Ley21210DepInsteInmed, DEP_LEY21210_INST, IIf(Ley21210DepAraucania, DEP_LEY21210_ARAUCANIA, 0))
               Q1 = Q1 & ", DepLey21256 = " & Abs(Ley21256DepInsteInmed)
               Q1 = Q1 & ", DepNormal = " & MesesDepNormal
               Q1 = Q1 & ", DepAcelerada = " & MesesDepAcelerada
               Q1 = Q1 & ", DepInstant = " & MesesDepInstantanea
               Q1 = Q1 & ", DepDecimaParte = " & MesesDepDecimaParte
               Q1 = Q1 & ", DepDecimaParte2 = " & MesesDepDecimaParte2
               Q1 = Q1 & ", TipoDep = " & TipoDep
               Q1 = Q1 & ", DepNormalHist = " & MesesDepNormalHist
               Q1 = Q1 & ", DepAceleradaHist = " & TipoDepHist
               Q1 = Q1 & ", DepInstantHist = " & MesesDepInstantaneaHist
               Q1 = Q1 & ", DepDecimaParteHist = " & MesesDepDecimaparteHist
               Q1 = Q1 & ", DepDecimaParte2Hist = " & MesesDepDecimaparte2Hist
               Q1 = Q1 & ", DepAcumHist = " & DepHistorica
               Q1 = Q1 & ", TipoDepHist = " & TipoDepHist
               Q1 = Q1 & ", IdCuenta = " & IdCuenta
               Q1 = Q1 & ", FechaImportFile = " & Int(DbMainDate)    'SqlNow(DbMain)
               Q1 = Q1 & ", PatenteRol = '" & ParaSQL(PatenteRol) & "'"
               Q1 = Q1 & ", NombreProy = '" & ParaSQL(NombreProy) & "'"
               Q1 = Q1 & ", FechaProy = " & FechaProy
               
               '2861733 tema 2
               Q1 = Q1 & ", idCCosto = '" & vIdCentroGestion & "'"
               Q1 = Q1 & ", IdAreaNeg = '" & vIdAreaNegocio & "'"
               '2861733 tema 2
                

               Q1 = Q1 & " WHERE IdActFijo = " & IdActFijo
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

               Call ExecSQL(DbMain, Q1)
               
               'ahora agregamos la ficha

               FldArray(0).FldName = "IdActFijo"
               FldArray(0).FldValue = IdActFijo
               FldArray(0).FldIsNum = True
               
               FldArray(1).FldName = "FechaIncorporacion"
               FldArray(1).FldValue = DtIncorporacion
               FldArray(1).FldIsNum = True
                     
               FldArray(2).FldName = "IdEmpresa"
               FldArray(2).FldValue = gEmpresa.id
               FldArray(2).FldIsNum = True
                           
               FldArray(3).FldName = "Ano"
               FldArray(3).FldValue = gEmpresa.Ano
               FldArray(3).FldIsNum = True
      
               IdFicha = AdvTbAddNewMult(DbMain, "ActFijoFicha", "IdFicha", FldArray)

               Q1 = "UPDATE ActFijoFicha SET "
           
               Q1 = Q1 & "  IdGrupo = " & IdGrupo
               Q1 = Q1 & ", FechaDisponible = " & DtDisponible
            
               Q1 = Q1 & ", PrecioFactura = " & Neto
               Q1 = Q1 & ", DerechosIntern = " & DerechosInternacion
               Q1 = Q1 & ", Transporte = " & Transporte
               Q1 = Q1 & ", ObrasAdapt = " & Adaptacion
               Q1 = Q1 & ", AdquiOtrosConceptos = " & OtrosAdquisicion
               
               PrecioAdquisicion = Neto + DerechosInternacion + Transporte + Adaptacion + OtrosAdquisicion
               
               Q1 = Q1 & ", PrecioAdquis = " & PrecioAdquisicion
            
               Q1 = Q1 & ", IVARecuperable = " & IVARec
               Q1 = Q1 & ", FormacionPers = " & FormacionPersonal
               Q1 = Q1 & ", ObrasReubic = " & Reubicacion
               Q1 = Q1 & ", GastoOtrosConceptos = " & OtrosGastos
               
               TotalGastos = IVARec + FormacionPersonal + Reubicacion + OtrosGastos
               
               Q1 = Q1 & ", TotalGastos = " & TotalGastos
               Q1 = Q1 & ", IdEmpresa = " & gEmpresa.id
               Q1 = Q1 & ", Ano = " & gEmpresa.Ano
            
               Q1 = Q1 & " WHERE IdFicha = " & IdFicha
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            
               Call ExecSQL(DbMain, Q1)

            End If
            
         End If

      Else
         NAFConError = NAFConError + 1
      End If
      
      r = r + 1

NextRec:
   Loop

   Close #Fd


   Frm.Pb_Proceso.Value = 100

   If NAFConError = 0 Then

      If SoloRevisar Then
         If r > 1 Then
            MsgBox1 "Proceso de Revisión finalizado exitosamente." & vbNewLine & vbNewLine & "Resultado:" & vbNewLine & vbNewLine & "Se encontraron " & r & " activos fijos en el archivo.", vbInformation + vbOKOnly
            ImportActFijoFile = True
         ElseIf r = 1 Then
            MsgBox1 "Proceso de Revisión finalizado exitosamente." & vbNewLine & vbNewLine & "Resultado:" & vbNewLine & vbNewLine & "Se encontró 1 activo fijo en el archivo.", vbInformation + vbOKOnly
            ImportActFijoFile = True
         Else   'r=0
            MsgBox1 "Proceso de Revisión finalizado." & vbNewLine & vbNewLine & "Resultado:" & vbNewLine & vbNewLine & "No se encontraron activos fijos en el archivo.", vbInformation + vbOKOnly
         End If
      Else
         If r > 1 Then
            MsgBox1 "Proceso de Importación finalizado exitosamente." & vbNewLine & vbNewLine & "Se importaron " & r & " activos fijos.", vbInformation + vbOKOnly
            ImportActFijoFile = True
         ElseIf r = 1 Then
            MsgBox1 "Proceso de Importación finalizado exitosamente." & vbNewLine & vbNewLine & "Se importó 1 activo fijo.", vbInformation + vbOKOnly
            ImportActFijoFile = True
         Else
            MsgBox1 "Proceso de Importación finalizado." & vbNewLine & vbNewLine & "No se importaron activos fijos.", vbInformation + vbOKOnly
         End If
      End If

   Else
      If NAFConError = 1 Then
         MsgBox1 "Se encontró 1 activo fijo con errores en el archivo de importación." & vbNewLine & vbNewLine & "No se realizó el proceso de importación.", vbExclamation + vbOKOnly
      Else
         MsgBox1 "Se encontraron " & NAFConError & " activos fijos con errores en el archivo de importación." & vbNewLine & vbNewLine & "No se realizó el proceso de importación.", vbExclamation + vbOKOnly
      End If

      If MsgBox1("¿Desea revisar el log de importación " & lFNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Frm.hWnd, "open", lFNameLogImp, "", "", SW_SHOW)
      End If
   End If

   Frm.Pb_Proceso.Value = 0
   Frm.Lb_Proceso = ""

End Function
Function GetIdGrupo(ByVal NombGrupo As String) As Long
   Dim Rs As Recordset
   Dim Q1 As String
   
   GetIdGrupo = 0
   If Trim(NombGrupo) <> "" Then
   
      Q1 = "SELECT IdGrupo FROM AFGrupos WHERE NombGrupo = '" & NombGrupo & "' AND IdEmpresa = " & gEmpresa.id
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         GetIdGrupo = vFld(Rs(0))
      End If
   
      Call CloseRs(Rs)
   
   End If
   
End Function
Public Function ImportOtrosDocs(Frm As Form, ByVal fname As String) As Boolean
   Dim Buf As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer, l As Integer
   Dim p As Long
   Dim Rc As Integer
   Dim Fd As Long
   Dim CampoInvalido As String
   Dim Row As Integer, r As Integer, ra As Integer
   Dim Descrip As String
   Dim PrimerCampo As String
   Dim Aux As String, AuxVal As Double
   Dim Valor As Double
   Dim DtEmision As Long, DtVenc As Long
   Dim RegConError As Integer
   Dim TipoDoc As String, IdTipoDoc As Integer
   Dim TipoLib As Integer
   Dim TxtDTE As String, DTE As Integer
   Dim NumDoc As String
   Dim IdEnt As Long
   Dim NotValidRut As Boolean
   Dim RutEnt As String, CodEnt As String, NombEnt As String
   Dim CodCta As String, IdCuenta As Long
   Dim NumInterno As Long
   Dim NombreCta As String, DescCta As String, UltNivelCta As Boolean
   Dim DocAnalitico As Integer
   Dim ConCta As String
   Dim FldArray(3) As AdvTbAddNew_t
   Dim IdDoc As Long
   Dim MsgAct As String

   ImportOtrosDocs = False

   lFNameLogImp = gImportPath & "\Log\ImpOtrosDocs-" & Format(Now, "yyyymmdd") & ".log"


   'abrimos el archivo
   Fd = FreeFile
   Open fname For Input As #Fd
   If ERR Then
      MsgErr fname
      ImportOtrosDocs = -ERR
      Exit Function
   End If

   'Campos
   'Fecha Emisión & vbTab & (TD) Tipo de Documento   & vbTab & DTE & vbTab & N° Doc.   & vbTab & RUT & vbTab &  Razón social & vbTab & Observaciones   & vbTab & Valor   & vbTab & Cod. Cuenta Banco & vbTab & Fecha Vencimiento   & vbTab & N° Interno & vbTab & Doc En Analitico

   r = 0
   ra = 0

   Do Until EOF(Fd)

      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf

      Aux = ReplaceStr(Buf, vbTab, "")
      Aux = Trim(Aux)
      If Aux = "" Then
         Buf = Aux
      End If
      Buf = Trim(Buf)

      '1er registro con nombres de campos
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 Then
         GoTo NextRec
      End If

      p = 1

      CampoInvalido = ""
   

      'Fecha emisión
      Aux = Trim(NextField2(Buf, p))
      DtEmision = ValFmtDate(Aux, False)
      If DtEmision = 0 Or Year(DtEmision) < gEmpresa.Ano - 1 Or Year(DtEmision) > gEmpresa.Ano Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Fecha emisión inválida. Debe corresponder al año anterior o al año actual.")
      End If
      
      'Tipo Doc
      TipoDoc = Trim(NextField2(Buf, p))
      If TipoDoc = "REM" Then
         TipoLib = LIB_REMU
      Else
         TipoLib = LIB_OTROS
      End If
      
      IdTipoDoc = FindTipoDoc(TipoLib, TipoDoc)
      If IdTipoDoc = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Tipo de documento inválido o no corresponde a Otros Documentos.")
      End If
               
      'DTE
      TxtDTE = Trim(NextField2(Buf, p))
      DTE = IIf(Val(TxtDTE) = 0 Or Trim(TxtDTE) = "", 0, 1)
      
      'NumDoc
      NumDoc = Trim(NextField2(Buf, p))
      If NumDoc = "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "N° de documento inválido.")
      End If
               
      'Entidad
      IdEnt = 0
      NotValidRut = False
      Aux = Trim(NextField2(Buf, p))
      If Aux = "0-0" Or Aux = "" Then
         RutEnt = ""
      Else
         RutEnt = vFmtCID(Aux)
         If RutEnt = "0" Or RutEnt = "" Then    'es inválido
            NotValidRut = True
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "RUT inválido")
         End If
      End If
      
      CodEnt = RutEnt
      
      'nombtre o razón social
      NombEnt = Trim(NextField2(Buf, p))
      If NombEnt = "" And RutEnt <> "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar nombre o razón social entidad.")
      End If
      
      If CampoInvalido = "" Then
         IdEnt = GetIdEntidad(FmtCID(RutEnt), NombEnt, False)
      End If
      
     
      'Descripción
      Descrip = Trim(NextField2(Buf, p))
      
      'Valor
      Valor = vFmt(Trim(NextField2(Buf, p)))
      
      If Valor <= 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor documento inválido.")
      End If
      
      'código cuenta Banco
      CodCta = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      If CodCta <> "" Then
         IdCuenta = GetIdCuenta(NombreCta, CodCta, DescCta, UltNivelCta)
         If UltNivelCta = False Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta inválido.")
         End If
      Else
         IdCuenta = 0
      End If
      
      'Fecha Vencim
      Aux = Trim(NextField2(Buf, p))
      DtVenc = 0
      If Aux <> "" Then
         DtVenc = ValFmtDate(Aux, False)
         If DtVenc = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Fecha vencimiento inválida.")
         End If
'      Else
'         CampoInvalido = CampoInvalido & "," & p
'         Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar Fecha Vencimiento.")
      End If
            
      If DtVenc > 0 And DtEmision > DtVenc Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Fecha de emisión mayor a la fecha de vencimiento.")
      End If
      
      NumInterno = vFmt(Trim(NextField2(Buf, p)))
      
      'Incluir en informe analítico
      DocAnalitico = Val(NextField2(Buf, p))
      DocAnalitico = IIf(DocAnalitico <> 0, 1, 0)
            
      IdDoc = 0
      
      'Si no hay errores
      
      If CampoInvalido = "" Then
      
         'veamos si este documento ya ha sido ingresado al sistema, sólo si no tiene entidad asociada o la entidad ya está en el sistema

         If RutEnt = "" Or IdEnt > 0 Then
         
            Q1 = "SELECT IdDoc FROM Documento "
            Q1 = Q1 & " WHERE TipoLib=" & TipoLib
            Q1 = Q1 & " AND TipoDoc=" & IdTipoDoc
            Q1 = Q1 & " AND NumDoc='" & NumDoc & "'"
            If CodCta <> "" Then
               Q1 = Q1 & " AND IdCtaBanco =" & IdCuenta
            Else
               Q1 = Q1 & " AND IdCtaBanco =0"
            End If
               
               
            Q1 = Q1 & " AND IdEntidad =" & IdEnt
            
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Set Rs = OpenRs(DbMain, Q1)
            
            If Rs.EOF = False Then   'documento ya existe
            
               IdDoc = vFld(Rs("IdDoc"))
                                 
               'Actualizamos el saldo y la cuenta del banco
               
               Q1 = "UPDATE Documento SET "
               Q1 = Q1 & "  Total =" & Valor
               Q1 = Q1 & ", SaldoDoc =" & Valor
               Q1 = Q1 & ", idCtaBanco =" & IdCuenta
               Q1 = Q1 & " WHERE IdDoc =" & IdDoc
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
               Call ExecSQL(DbMain, Q1)
              
               ra = ra + 1
               
            End If
         
            Call CloseRs(Rs)
      
         End If
         'si la entidad no existe y/o el documento no existe, los insertamos
         
         If IdDoc = 0 Then
         
            If RutEnt <> "" And IdEnt = 0 Then
                  
               'insertamos la nueva entidad
   
               FldArray(0).FldName = "RUT"
               FldArray(0).FldValue = RutEnt
               FldArray(0).FldIsNum = False
               
               FldArray(1).FldName = "Codigo"
               FldArray(1).FldValue = CodEnt
               FldArray(1).FldIsNum = False
                     
               FldArray(2).FldName = "Nombre"
               FldArray(2).FldValue = NombEnt
               FldArray(2).FldIsNum = False
                           
               FldArray(3).FldName = "IdEmpresa"
               FldArray(3).FldValue = gEmpresa.id
               FldArray(3).FldIsNum = True
                     
               IdEnt = AdvTbAddNewMult(DbMain, "Entidades", "IdEntidad", FldArray)
                                             
            End If
            
            'Ahora insertamos el documento
            FldArray(0).FldName = "IdUsuario"
            FldArray(0).FldValue = gUsuario.IdUsuario
            FldArray(0).FldIsNum = True
            
            FldArray(1).FldName = "FechaCreacion"
            FldArray(1).FldValue = CLng(Int(Now))
            FldArray(1).FldIsNum = True
                  
            FldArray(2).FldName = "IdEmpresa"
            FldArray(2).FldValue = gEmpresa.id
            FldArray(2).FldIsNum = True
                        
            FldArray(3).FldName = "Ano"
            FldArray(3).FldValue = gEmpresa.Ano
            FldArray(3).FldIsNum = True
                  
            IdDoc = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray)
                     
            Q1 = "UPDATE Documento SET "
            Q1 = Q1 & "  TipoDoc =" & IdTipoDoc
            Q1 = Q1 & ", TipoLib =" & TipoLib
            Q1 = Q1 & ", NumDoc ='" & NumDoc & "'"
            Q1 = Q1 & ", DTE =" & DTE
            Q1 = Q1 & ", CorrInterno =" & NumInterno
            Q1 = Q1 & ", idEntidad =" & IdEnt
            Q1 = Q1 & ", FEmision =" & DtEmision
            Q1 = Q1 & ", FEmisionOri =" & DtEmision
            Q1 = Q1 & ", FVenc =" & DtVenc
            Q1 = Q1 & ", Total =" & Valor
            Q1 = Q1 & ", Estado =" & ED_APROBADO
            Q1 = Q1 & ", Descrip ='" & ParaSQL(Descrip) & "'"
            Q1 = Q1 & ", SaldoDoc =" & Valor
            Q1 = Q1 & ", DocOtrosEnAnalitico =" & IIf(DocAnalitico <> 0, 1, 0)
            Q1 = Q1 & ", idCtaBanco =" & IdCuenta
            Q1 = Q1 & " WHERE IdDoc =" & IdDoc
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Call ExecSQL(DbMain, Q1)
            
            r = r + 1
            
            'Tracking 3227543
            Call SeguimientoDocumento(IdDoc, gEmpresa.id, gEmpresa.Ano, "ImpExpFnc.ImportOtrosDocs", Q1, 1, "", gUsuario.IdUsuario, 2, 1)
            ' fin 3227543
      
         End If
         
      End If
   
NextRec:
   Loop

   Close #Fd

   MsgAct = ""
   
   If RegConError = 0 Then
   
      If ra > 1 Then
         MsgAct = "Se actualizaron " & ra & " documentos."
      ElseIf ra = 1 Then
         MsgAct = "Se actualizó 1 documento."
      End If

      If r > 1 Then
         MsgBox1 "Proceso de Importación finalizado exitosamente." & vbNewLine & vbNewLine & "Se importaron " & r & " documentos." & IIf(MsgAct <> "", vbNewLine & vbNewLine & MsgAct, ""), vbInformation + vbOKOnly
         ImportOtrosDocs = True
      ElseIf r = 1 Then
         MsgBox1 "Proceso de Importación finalizado exitosamente." & vbNewLine & vbNewLine & "Se importó 1 documento." & IIf(MsgAct <> "", vbNewLine & vbNewLine & MsgAct, ""), vbInformation + vbOKOnly
         ImportOtrosDocs = True
      Else
         MsgBox1 "Proceso de Importación finalizado." & vbNewLine & vbNewLine & "No se importaron nuevos documentos." & IIf(MsgAct <> "", vbNewLine & vbNewLine & MsgAct, "") & vbNewLine & "Revisar el log de importación " & lFNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores", vbInformation + vbOKOnly
      End If

   Else
      If RegConError = 1 Then
         MsgBox1 "Se encontró 1 documento con errores en el archivo de importación." & vbNewLine & vbNewLine & "No se realizó el proceso de importación.", vbExclamation + vbOKOnly
      Else
         MsgBox1 "Se encontraron " & RegConError & " documentos con errores en el archivo de importación." & vbNewLine & vbNewLine & "No se realizó el proceso de importación.", vbExclamation + vbOKOnly
      End If

      If MsgBox1("¿Desea revisar el log de importación " & lFNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Frm.hWnd, "open", lFNameLogImp, "", "", SW_SHOW)
      End If
   End If


End Function

Public Function ImportOtrosDocFull(Frm As Form, ByVal fname As String) As Boolean
   Dim Buf As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer, l As Integer
   Dim p As Long
   Dim Rc As Integer
   Dim Fd As Long
   Dim CampoInvalido As String
   Dim Row As Integer, r As Integer, ra As Integer
   Dim Descrip As String
   Dim PrimerCampo As String
   Dim Aux As String, AuxVal As Double
   Dim Valor As Double
   Dim DtEmision As Long, DtVenc As Long
   Dim RegConError As Integer
   Dim TipoDoc As String, IdTipoDoc As Integer
   Dim TipoLib As Integer
   Dim TxtDTE As String, DTE As Integer
   Dim NumDoc As String
   Dim IdEnt As Long
   Dim NotValidRut As Boolean
   Dim RutEnt As String, CodEnt As String, NombEnt As String
   Dim CodCta As String, IdCuenta As Long
   Dim NumInterno As Long
   Dim NombreCta As String, DescCta As String, UltNivelCta As Boolean
   Dim DocAnalitico As Integer
   Dim Tratamiento As Integer
   Dim ConCta As String
   Dim FldArray(3) As AdvTbAddNew_t
   Dim IdDoc As Long
   Dim MsgAct As String

   ImportOtrosDocFull = False
   Tratamiento = 0
   lFNameLogImp = gImportPath & "\Log\ImpOtrosDocs-" & Format(Now, "yyyymmdd") & ".log"


   'abrimos el archivo
   Fd = FreeFile
   Open fname For Input As #Fd
   If ERR Then
      MsgErr fname
      ImportOtrosDocFull = -ERR
      Exit Function
   End If

   'Campos
   'Fecha Emisión & vbTab & (TD) Tipo de Documento   & vbTab & DTE & vbTab & N° Doc.   & vbTab & RUT & vbTab &  Razón social & vbTab & Observaciones   & vbTab & Valor   & vbTab & Cod. Cuenta Banco & vbTab & Fecha Vencimiento   & vbTab & N° Interno & vbTab & Doc En Analitico

   r = 0
   ra = 0

   Do Until EOF(Fd)

      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf

      Aux = ReplaceStr(Buf, vbTab, "")
      Aux = Trim(Aux)
      If Aux = "" Then
         Buf = Aux
      End If
      Buf = Trim(Buf)

      '1er registro con nombres de campos
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 Then
         GoTo NextRec
      End If

      p = 1

      CampoInvalido = ""
   

      'Fecha emisión
      Aux = Trim(NextField2(Buf, p))
      DtEmision = ValFmtDate(Aux, False)
      If DtEmision = 0 Or Year(DtEmision) < gEmpresa.Ano - 1 Or Year(DtEmision) > gEmpresa.Ano Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Fecha emisión inválida. Debe corresponder al año anterior o al año actual.")
      End If
      
      'Tipo Doc
      TipoDoc = Trim(NextField2(Buf, p))
      If TipoDoc = "REM" Then
         TipoLib = LIB_REMU
      Else
         TipoLib = LIB_OTROFULL
      End If
      
      TipoDoc = "ODF"
      IdTipoDoc = 1 'FindTipoDoc(TipoLib, TipoDoc)
      If IdTipoDoc = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Tipo de documento inválido o no corresponde a Otros Documentos.")
      End If
               
      'DTE
      TxtDTE = Trim(NextField2(Buf, p))
      DTE = IIf(Val(TxtDTE) = 0 Or Trim(TxtDTE) = "", 0, 1)
      
      'NumDoc
      NumDoc = Trim(NextField2(Buf, p))
      If NumDoc = "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "N° de documento inválido.")
      End If
               
      'Entidad
      IdEnt = 0
      NotValidRut = False
      Aux = Trim(NextField2(Buf, p))
      If Aux = "0-0" Or Aux = "" Then
         RutEnt = ""
      Else
         RutEnt = vFmtCID(Aux)
         If RutEnt = "0" Or RutEnt = "" Then    'es inválido
            NotValidRut = True
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "RUT inválido")
         End If
      End If
      
      CodEnt = RutEnt
      
      'nombtre o razón social
      NombEnt = Trim(NextField2(Buf, p))
      If NombEnt = "" And RutEnt <> "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar nombre o razón social entidad.")
      End If
      
      If CampoInvalido = "" Then
         IdEnt = GetIdEntidad(FmtCID(RutEnt), NombEnt, False)
      End If
      
     
      'Descripción
      Descrip = Trim(NextField2(Buf, p))
      
      'Valor
      Valor = vFmt(Trim(NextField2(Buf, p)))
      
      If Valor <= 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Valor documento inválido.")
      End If
      
      'código cuenta Banco
      CodCta = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      If CodCta <> "" Then
         NombreCta = ""
         IdCuenta = GetIdCuenta(NombreCta, CodCta, DescCta, UltNivelCta)
         If UltNivelCta = False Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta inválido.")
         End If
      Else
         IdCuenta = 0
      End If
      
      'Fecha Vencim
      Aux = Trim(NextField2(Buf, p))
      DtVenc = 0
      If Aux <> "" Then
         DtVenc = ValFmtDate(Aux, False)
         If DtVenc = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, fname, l, "Fecha vencimiento inválida.")
         End If
'      Else
'         CampoInvalido = CampoInvalido & "," & p
'         Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar Fecha Vencimiento.")
      End If
            
      If DtVenc > 0 And DtEmision > DtVenc Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, fname, l, "Fecha de emisión mayor a la fecha de vencimiento.")
      End If
      
      NumInterno = vFmt(Trim(NextField2(Buf, p)))
      
      'Incluir en informe analítico
      DocAnalitico = Val(NextField2(Buf, p))
      DocAnalitico = IIf(DocAnalitico <> 0, 1, 0)
      Tratamiento = Val(NextField2(Buf, p))
      If Tratamiento <> 1 And Tratamiento <> 2 Then
        Tratamiento = 1
      End If
      
      IdDoc = 0
      
      'Si no hay errores
      
      If CampoInvalido = "" Then
      
         'veamos si este documento ya ha sido ingresado al sistema, sólo si no tiene entidad asociada o la entidad ya está en el sistema

         If RutEnt = "" Or IdEnt > 0 Then
         
            Q1 = "SELECT IdDoc FROM Documento "
            Q1 = Q1 & " WHERE TipoLib=" & TipoLib
            Q1 = Q1 & " AND TipoDoc=" & IdTipoDoc
            Q1 = Q1 & " AND NumDoc='" & NumDoc & "'"
            If CodCta <> "" Then
               Q1 = Q1 & " AND IdCtaBanco =" & IdCuenta
            Else
               Q1 = Q1 & " AND IdCtaBanco =0"
            End If
               
               
            Q1 = Q1 & " AND IdEntidad =" & IdEnt
            
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Set Rs = OpenRs(DbMain, Q1)
            
            If Rs.EOF = False Then   'documento ya existe
            
               IdDoc = vFld(Rs("IdDoc"))
                                 
               'Actualizamos el saldo y la cuenta del banco
               
               Q1 = "UPDATE Documento SET "
               Q1 = Q1 & "  Total =" & Valor
               Q1 = Q1 & ", SaldoDoc =" & Valor
               Q1 = Q1 & ", idCtaBanco =" & IdCuenta
               Q1 = Q1 & " WHERE IdDoc =" & IdDoc
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
               Call ExecSQL(DbMain, Q1)
              
               ra = ra + 1
               
            End If
         
            Call CloseRs(Rs)
      
         End If
         'si la entidad no existe y/o el documento no existe, los insertamos
         
         If IdDoc = 0 Then
         
            If RutEnt <> "" And IdEnt = 0 Then
                  
               'insertamos la nueva entidad
   
               FldArray(0).FldName = "RUT"
               FldArray(0).FldValue = RutEnt
               FldArray(0).FldIsNum = False
               
               FldArray(1).FldName = "Codigo"
               FldArray(1).FldValue = CodEnt
               FldArray(1).FldIsNum = False
                     
               FldArray(2).FldName = "Nombre"
               FldArray(2).FldValue = NombEnt
               FldArray(2).FldIsNum = False
                           
               FldArray(3).FldName = "IdEmpresa"
               FldArray(3).FldValue = gEmpresa.id
               FldArray(3).FldIsNum = True
                     
               IdEnt = AdvTbAddNewMult(DbMain, "Entidades", "IdEntidad", FldArray)
                                             
            End If
            
            'Ahora insertamos el documento
            FldArray(0).FldName = "IdUsuario"
            FldArray(0).FldValue = gUsuario.IdUsuario
            FldArray(0).FldIsNum = True
            
            FldArray(1).FldName = "FechaCreacion"
            FldArray(1).FldValue = CLng(Int(Now))
            FldArray(1).FldIsNum = True
                  
            FldArray(2).FldName = "IdEmpresa"
            FldArray(2).FldValue = gEmpresa.id
            FldArray(2).FldIsNum = True
                        
            FldArray(3).FldName = "Ano"
            FldArray(3).FldValue = gEmpresa.Ano
            FldArray(3).FldIsNum = True
                  
            IdDoc = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray)
                     
            Q1 = "UPDATE Documento SET "
            Q1 = Q1 & "  TipoDoc =" & IdTipoDoc
            Q1 = Q1 & ", TipoLib =" & TipoLib
            Q1 = Q1 & ", NumDoc ='" & NumDoc & "'"
            Q1 = Q1 & ", DTE =" & DTE
            Q1 = Q1 & ", CorrInterno =" & NumInterno
            Q1 = Q1 & ", idEntidad =" & IdEnt
            Q1 = Q1 & ", FEmision =" & DtEmision
            Q1 = Q1 & ", FEmisionOri =" & DtEmision
            Q1 = Q1 & ", FVenc =" & DtVenc
            Q1 = Q1 & ", Total =" & Valor
            Q1 = Q1 & ", Estado =" & ED_APROBADO
            Q1 = Q1 & ", Descrip ='" & ParaSQL(Descrip) & "'"
            Q1 = Q1 & ", SaldoDoc =" & Valor
            Q1 = Q1 & ", DocOtrosEnAnalitico =" & IIf(DocAnalitico <> 0, 1, 0)
            Q1 = Q1 & ", idCtaBanco =" & IdCuenta
            Q1 = Q1 & ", Tratamiento =" & IIf(Tratamiento = 0, 1, Tratamiento)
            Q1 = Q1 & " WHERE IdDoc =" & IdDoc
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
            Call ExecSQL(DbMain, Q1)
            
            r = r + 1
            
            'Tracking 3227543
            Call SeguimientoDocumento(IdDoc, gEmpresa.id, gEmpresa.Ano, "ImpExpFnc.ImportOtrosDocFull", Q1, 1, "", gUsuario.IdUsuario, 2, 1)
            ' fin 3227543
      
         End If
         
      End If
   
NextRec:
   Loop

   Close #Fd

   MsgAct = ""
   
   If RegConError = 0 Then
   
      If ra > 1 Then
         MsgAct = "Se actualizaron " & ra & " documentos."
      ElseIf ra = 1 Then
         MsgAct = "Se actualizó 1 documento."
      End If

      If r > 1 Then
         MsgBox1 "Proceso de Importación finalizado exitosamente." & vbNewLine & vbNewLine & "Se importaron " & r & " documentos." & IIf(MsgAct <> "", vbNewLine & vbNewLine & MsgAct, ""), vbInformation + vbOKOnly
         ImportOtrosDocFull = True
      ElseIf r = 1 Then
         MsgBox1 "Proceso de Importación finalizado exitosamente." & vbNewLine & vbNewLine & "Se importó 1 documento." & IIf(MsgAct <> "", vbNewLine & vbNewLine & MsgAct, ""), vbInformation + vbOKOnly
         ImportOtrosDocFull = True
      Else
         MsgBox1 "Proceso de Importación finalizado." & vbNewLine & vbNewLine & "No se importaron nuevos documentos" & IIf(MsgAct <> "", vbNewLine & vbNewLine & MsgAct, ""), vbInformation + vbOKOnly
         
         If MsgBox1("¿Desea revisar el log de importación " & lFNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
           Call ShellExecute(Frm.hWnd, "open", lFNameLogImp, "", "", SW_SHOW)
         End If
      End If
   Else
      If RegConError = 1 Then
         MsgBox1 "Se encontró 1 documento con errores en el archivo de importación." & vbNewLine & vbNewLine & "No se realizó el proceso de importación.", vbExclamation + vbOKOnly
      Else
         MsgBox1 "Se encontraron " & RegConError & " documentos con errores en el archivo de importación." & vbNewLine & vbNewLine & "No se realizó el proceso de importación.", vbExclamation + vbOKOnly
      End If

      If MsgBox1("¿Desea revisar el log de importación " & lFNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Frm.hWnd, "open", lFNameLogImp, "", "", SW_SHOW)
      End If
   End If


End Function

Public Sub AddLogExp(ByVal FNameLogExp As String, ByVal IdReg As String, ByVal Msg As String)
   Dim Er As Integer, sErr As String, Fd As Integer

   Er = ERR
   sErr = Error
   On Error Resume Next

   Fd = FreeFile
   Open FNameLogExp For Append Access Write As #Fd

   Print #Fd, Format(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & IdReg & ": " & vbTab & Msg
   
   Close #Fd
   On Error GoTo 0

   ERR = Er

End Sub
'Esta función llena la configuración de las cuentas asociadas a las entidades para la importación de libro de compras o ventas del SII, en base a los documentos del año actual
Public Sub FillCuentasUtilizadas(ByVal TipoLib As Integer, ByVal IdTipoValLib As Integer, ByVal CodCtaFldName As String, ByVal CodCCostoFldName As String, ByVal CodANegFldName As String, ByVal CodCuentaDef As String)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdEntidad As Long
   
   'cuentas contables
   Q1 = "SELECT IdEntidad, Cuentas.Codigo, Count(*) as N "
   Q1 = Q1 & " FROM (Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento")
   Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib & " AND MovDocumento.IdTipoValLib = " & IdTipoValLib
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY IdEntidad, Codigo "
   Q1 = Q1 & " ORDER BY IdEntidad, Count(*) Desc"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   IdEntidad = 0
   Do While Not Rs.EOF
   
      If vFld(Rs("IdEntidad")) <> IdEntidad Then
         IdEntidad = vFld(Rs("IdEntidad"))
   
         If vFld(Rs("Codigo")) <> CodCuentaDef Then
            Q1 = "UPDATE Entidades SET " & CodCtaFldName & " = '" & vFld(Rs("Codigo")) & "'"
            Q1 = Q1 & " WHERE IdEntidad = " & vFld(Rs("IdEntidad")) & " AND (" & CodCtaFldName & " IS NULL OR " & CodCtaFldName & " = '') "
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
            Call ExecSQL(DbMain, Q1)
         End If
      End If
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)

   'áreas de negocio
   Q1 = "SELECT IdEntidad, AreaNegocio.Codigo, Count(*) as N "
   Q1 = Q1 & " FROM (Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " INNER JOIN AreaNegocio ON MovDocumento.IdAreaNeg = AreaNegocio.IdAreaNegocio "
   Q1 = Q1 & JoinEmpAno(gDbType, "AreaNegocio", "MovDocumento", True, True)
   Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib & " AND MovDocumento.IdTipoValLib = " & IdTipoValLib
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY IdEntidad, Codigo "
   Q1 = Q1 & " ORDER BY IdEntidad, Count(*) Desc"

   Set Rs = OpenRs(DbMain, Q1)

   IdEntidad = 0
   Do While Not Rs.EOF

      If vFld(Rs("IdEntidad")) <> IdEntidad Then
         IdEntidad = vFld(Rs("IdEntidad"))

         Q1 = "UPDATE Entidades SET " & CodANegFldName & " = '" & vFld(Rs("Codigo")) & "'"
         Q1 = Q1 & " WHERE IdEntidad = " & vFld(Rs("IdEntidad")) & " AND (" & CodANegFldName & " IS NULL OR " & CodANegFldName & " = '') "
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
         Call ExecSQL(DbMain, Q1)

      End If

      Rs.MoveNext

   Loop

   Call CloseRs(Rs)

   'centros de costo
   Q1 = "SELECT IdEntidad, CentroCosto.Codigo, Count(*) as N "
   Q1 = Q1 & " FROM (Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " INNER JOIN CentroCosto ON MovDocumento.IdCCosto = CentroCosto.IdCCosto "
   Q1 = Q1 & JoinEmpAno(gDbType, "CentroCosto", "MovDocumento", True, True)
   Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib & " AND MovDocumento.IdTipoValLib = " & IdTipoValLib
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY IdEntidad, Codigo "
   Q1 = Q1 & " ORDER BY IdEntidad, Count(*) Desc"

   Set Rs = OpenRs(DbMain, Q1)

   IdEntidad = 0
   Do While Not Rs.EOF

      If vFld(Rs("IdEntidad")) <> IdEntidad Then
         IdEntidad = vFld(Rs("IdEntidad"))

         Q1 = "UPDATE Entidades SET " & CodCCostoFldName & " = '" & vFld(Rs("Codigo")) & "'"
         Q1 = Q1 & " WHERE IdEntidad = " & vFld(Rs("IdEntidad")) & " AND (" & CodCCostoFldName & " IS NULL OR " & CodCCostoFldName & " = '') "
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
         Call ExecSQL(DbMain, Q1)

      End If

      Rs.MoveNext

   Loop

   Call CloseRs(Rs)

End Sub
'Esta función llena la configuración de las cuentas asociadas a las entidades para la importación de libro de compras o ventas del SII, en base a los documentos del año anterior
Public Sub FillCuentasUtilizadasAnoAnt(ByVal TipoLib As Integer, ByVal IdTipoValLib As Integer, ByVal CodCtaFldName As String, ByVal CodCCostoFldName As String, ByVal CodANegFldName As String, ByVal CodCuentaDef As String)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdEntidad As Long
   Dim fname As String
   
   
   If Not gEmpresa.TieneAnoAnt Then
      Exit Sub
   End If
   
   If gEmprSeparadas Then
   
#If DATACON = 1 Then       'Access
   
      'linkeamos las tablas de Documento y MovDocumento del año anterior
      
      fname = gDbPath & "\Empresas\" & gEmpresa.Ano - 1 & "\" & gEmpresa.Rut & ".mdb"
      
      If Not ExistFile(fname) Then
         Exit Sub
      End If
      
      Call LinkMdbTable(DbMain, fname, "Documento", "DocumentoAnt", , , gEmpresa.ConnStr)
      Call LinkMdbTable(DbMain, fname, "MovDocumento", "MovDocumentoAnt", , , gEmpresa.ConnStr)
      Call LinkMdbTable(DbMain, fname, "Cuentas", "CuentasAnt", , , gEmpresa.ConnStr)
      Call LinkMdbTable(DbMain, fname, "AreaNegocio", "AreaNegocioAnt", , , gEmpresa.ConnStr)
      Call LinkMdbTable(DbMain, fname, "CentroCosto", "CentroCostoAnt", , , gEmpresa.ConnStr)
            
#End If

      Q1 = "SELECT IdEntidad, CuentasAnt.Codigo, Count(*) as N "
      Q1 = Q1 & " FROM ((DocumentoAnt INNER JOIN MovDocumentoAnt ON DocumentoAnt.IdDoc = MovDocumentoAnt.IdDoc) "
      Q1 = Q1 & " INNER JOIN CuentasAnt ON MovDocumentoAnt.IdCuenta = CuentasAnt.IdCuenta) "
      Q1 = Q1 & " INNER JOIN Cuentas ON CuentasAnt.Codigo = Cuentas.Codigo "
      Q1 = Q1 & " WHERE DocumentoAnt.TipoLib = " & TipoLib & " AND MovDocumentoAnt.IdTipoValLib = " & IdTipoValLib
      Q1 = Q1 & " GROUP BY IdEntidad, CuentasAnt.Codigo "
      Q1 = Q1 & " ORDER BY IdEntidad, Count(*) Desc"
     
   Else
   
      Q1 = "SELECT IdEntidad, Cuentas.Codigo, Count(*) as N "
      Q1 = Q1 & " FROM ((Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
      Q1 = Q1 & " INNER JOIN Cuentas ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento") & " )"
      Q1 = Q1 & " INNER JOIN Cuentas As CuentasNew ON Cuentas.Codigo = CuentasNew.Codigo "
      Q1 = Q1 & " AND Cuentas.IdEmpresa = CuentasNew.IdEmpresa AND Cuentas.Ano = CuentasNew.Ano - 1"
      Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib & " AND MovDocumento.IdTipoValLib = " & IdTipoValLib
      Q1 = Q1 & " AND MovDocumento.IdEmpresa = " & gEmpresa.id & " AND MovDocumento.Ano = " & gEmpresa.Ano - 1
      Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano - 1
      Q1 = Q1 & " GROUP BY IdEntidad, Cuentas.Codigo "
      Q1 = Q1 & " ORDER BY IdEntidad, Count(*) Desc"
   
   End If
   
   Set Rs = OpenRs(DbMain, Q1)
   
   IdEntidad = 0
   Do While Not Rs.EOF
   
      If vFld(Rs("IdEntidad")) <> IdEntidad Then
         IdEntidad = vFld(Rs("IdEntidad"))
   
         If vFld(Rs("Codigo")) <> CodCuentaDef Then
            Q1 = "UPDATE Entidades SET " & CodCtaFldName & " = '" & vFld(Rs("Codigo")) & "'"
            Q1 = Q1 & " WHERE IdEntidad = " & vFld(Rs("IdEntidad")) & " AND (" & CodCtaFldName & " IS NULL OR " & CodCtaFldName & " = '') "
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
            Call ExecSQL(DbMain, Q1)
         End If
         
      End If
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   If gEmprSeparadas Then
      Q1 = "SELECT IdEntidad, AreaNegocioAnt.Codigo, Count(*) as N "
      Q1 = Q1 & " FROM ((DocumentoAnt INNER JOIN MovDocumentoAnt ON DocumentoAnt.IdDoc = MovDocumentoAnt.IdDoc) "
      Q1 = Q1 & " INNER JOIN AreaNegocioAnt ON MovDocumentoAnt.IdAreaNeg = AreaNegocioAnt.IdAreaNegocio) "
      Q1 = Q1 & " INNER JOIN AreaNegocio ON AreaNegocioAnt.Codigo = AreaNegocio.Codigo "
      Q1 = Q1 & " WHERE DocumentoAnt.TipoLib = " & TipoLib & " AND MovDocumentoAnt.IdTipoValLib = " & IdTipoValLib
      Q1 = Q1 & " GROUP BY IdEntidad, AreaNegocioAnt.Codigo "
      Q1 = Q1 & " ORDER BY IdEntidad, Count(*) Desc"
      
   Else
      Q1 = "SELECT IdEntidad, AreaNegocio.Codigo, Count(*) as N "
      Q1 = Q1 & " FROM ((Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & "  AND Documento.IdEmpresa = MovDocumento.IdEmpresa AND Documento.Ano = MovDocumento.Ano )"
      Q1 = Q1 & " INNER JOIN AreaNegocio ON MovDocumento.IdAreaNeg = AreaNegocio.IdAreaNegocio "
      Q1 = Q1 & "  AND AreaNegocio.IdEmpresa = MovDocumento.IdEmpresa )"
      Q1 = Q1 & " INNER JOIN AreaNegocio As AreaNegocioNew ON AreaNegocio.Codigo = AreaNegocioNew.Codigo "
      Q1 = Q1 & "  AND AreaNegocio.IdEmpresa = AreaNegocioNew.IdEmpresa "
      Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib & " AND MovDocumento.IdTipoValLib = " & IdTipoValLib
      Q1 = Q1 & "  AND MovDocumento.IdEmpresa = " & gEmpresa.id & " AND MovDocumento.Ano = " & gEmpresa.Ano - 1
      Q1 = Q1 & "  AND AreaNegocio.IdEmpresa = " & gEmpresa.id
      Q1 = Q1 & " GROUP BY IdEntidad, AreaNegocio.Codigo "
      Q1 = Q1 & " ORDER BY IdEntidad, Count(*) Desc"

   End If
   
   Set Rs = OpenRs(DbMain, Q1)

   IdEntidad = 0
   Do While Not Rs.EOF

      If vFld(Rs("IdEntidad")) <> IdEntidad Then
         IdEntidad = vFld(Rs("IdEntidad"))

         Q1 = "UPDATE Entidades SET " & CodANegFldName & " = '" & vFld(Rs("Codigo")) & "'"
         Q1 = Q1 & " WHERE IdEntidad = " & vFld(Rs("IdEntidad")) & " AND (" & CodANegFldName & " IS NULL OR " & CodANegFldName & " = '') "
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
         Call ExecSQL(DbMain, Q1)

      End If

      Rs.MoveNext

   Loop

   Call CloseRs(Rs)

   If gEmprSeparadas Then
      Q1 = "SELECT IdEntidad, CentroCostoAnt.Codigo, Count(*) as N "
      Q1 = Q1 & " FROM ((DocumentoAnt INNER JOIN MovDocumentoAnt ON DocumentoAnt.IdDoc = MovDocumentoAnt.IdDoc) "
      Q1 = Q1 & " INNER JOIN CentroCostoAnt ON MovDocumentoAnt.IdCCosto = CentroCostoAnt.IdCCosto) "
      Q1 = Q1 & " INNER JOIN CentroCosto ON CentroCostoAnt.Codigo = CentroCosto.Codigo "
      Q1 = Q1 & " WHERE DocumentoAnt.TipoLib = " & TipoLib & " AND MovDocumentoAnt.IdTipoValLib = " & IdTipoValLib
      Q1 = Q1 & " GROUP BY IdEntidad, CentroCostoAnt.Codigo "
      Q1 = Q1 & " ORDER BY IdEntidad, Count(*) Desc"
      
   Else
      Q1 = "SELECT IdEntidad, CentroCosto.Codigo, Count(*) as N "
      Q1 = Q1 & " FROM ((Documento INNER JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc "
      Q1 = Q1 & "  AND Documento.IdEmpresa = MovDocumento.IdEmpresa AND Documento.Ano = MovDocumento.Ano )"
      Q1 = Q1 & " INNER JOIN CentroCosto ON MovDocumento.IdAreaNeg = CentroCosto.IdCCosto "
      Q1 = Q1 & "  AND CentroCosto.IdEmpresa = MovDocumento.IdEmpresa )"
      Q1 = Q1 & " INNER JOIN CentroCosto As CentroCostoNew ON CentroCosto.Codigo = CentroCostoNew.Codigo "
      Q1 = Q1 & "  AND CentroCosto.IdEmpresa = CentroCostoNew.IdEmpresa "
      Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib & " AND MovDocumento.IdTipoValLib = " & IdTipoValLib
      Q1 = Q1 & "  AND MovDocumento.IdEmpresa = " & gEmpresa.id & " AND MovDocumento.Ano = " & gEmpresa.Ano - 1
      Q1 = Q1 & "  AND CentroCosto.IdEmpresa = " & gEmpresa.id
      Q1 = Q1 & " GROUP BY IdEntidad, CentroCosto.Codigo "
      Q1 = Q1 & " ORDER BY IdEntidad, Count(*) Desc"

   End If
   
   Set Rs = OpenRs(DbMain, Q1)

   IdEntidad = 0
   Do While Not Rs.EOF

      If vFld(Rs("IdEntidad")) <> IdEntidad Then
         IdEntidad = vFld(Rs("IdEntidad"))

         Q1 = "UPDATE Entidades SET " & CodCCostoFldName & " = '" & vFld(Rs("Codigo")) & "'"
         Q1 = Q1 & " WHERE IdEntidad = " & vFld(Rs("IdEntidad")) & " AND (" & CodCCostoFldName & " IS NULL OR " & CodCCostoFldName & " = '') "
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
         Call ExecSQL(DbMain, Q1)

      End If

      Rs.MoveNext

   Loop

   Call CloseRs(Rs)

   If gEmprSeparadas Then
#If DATACON = 1 Then       'Access
      Call UnLinkTable(DbMain, "DocumentoAnt")
      Call UnLinkTable(DbMain, "MovDocumentoAnt")
      Call UnLinkTable(DbMain, "CuentasAnt")
      Call UnLinkTable(DbMain, "AreaNegocioAnt")
      Call UnLinkTable(DbMain, "CentroCostoAnt")
#End If
   End If
   
End Sub

Public Function LoadCodCuentasDefLibros(ByVal TipoLib As Integer, CodCtaAfecto As String, CodCtaExento As String, CodCtaTotal As String) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Item As String
   
   LoadCodCuentasDefLibros = False
            
   If TipoLib > 0 Then
   
      Q1 = "SELECT Codigo, TipoValor "
      Q1 = Q1 & " FROM CuentasBasicas  "
      Q1 = Q1 & " INNER JOIN Cuentas ON CuentasBasicas.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & " AND Cuentas.IdEmpresa = CuentasBasicas.IdEmpresa AND Cuentas.Ano = CuentasBasicas.Ano "
      Q1 = Q1 & " WHERE TipoLib = " & TipoLib
      Q1 = Q1 & " AND CuentasBasicas.IdEmpresa = " & gEmpresa.id & " AND CuentasBasicas.Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY TipoValor, Id "
      
      Set Rs = OpenRs(DbMain, Q1)
   
      Do While Rs.EOF = False
                      
         Select Case vFld(Rs("TipoValor"))
         
            Case LIBVENTAS_AFECTO, LIBCOMPRAS_AFECTO
            
               If CodCtaAfecto = "" Then
                  CodCtaAfecto = vFld(Rs("Codigo"))
               End If
               
            Case LIBVENTAS_EXENTO, LIBCOMPRAS_EXENTO
            
               If CodCtaExento = "" Then
                  CodCtaExento = vFld(Rs("Codigo"))
               End If
            
            Case LIBVENTAS_TOTAL, LIBCOMPRAS_TOTAL
            
               If CodCtaTotal = "" Then
                  CodCtaTotal = vFld(Rs("Codigo"))
               End If
               
         End Select
                           
         Rs.MoveNext
         
      Loop
      Call CloseRs(Rs)
      
   End If
   
   If CodCtaAfecto = "" Or CodCtaExento = "" Or CodCtaTotal = "" Then
      MsgBox1 "Falta definir las cuentas básicas para los libros de compras, ventas y/o retenciones." & vbCrLf & vbCrLf & "Utilice el menú Configuración >> Configuración inicial >> Definir cuentas básicas", vbExclamation
      Exit Function
   End If
         
   LoadCodCuentasDefLibros = True
End Function
