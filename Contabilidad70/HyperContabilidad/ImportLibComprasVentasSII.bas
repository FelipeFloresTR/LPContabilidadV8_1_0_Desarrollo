Attribute VB_Name = "ImportLibComprasVentasSII"
Option Explicit

Const IMP_33 = ",14,17,18,19,23,24,25,26,27,28,29,35,44,45,46,50,51,52,53,271,"
Const IMP_45_46 = ",15,17,30,301,31,32,321,33,331,34,341,36,361,37,371,38,39,41,47,48,481,49,60,"
Const IMP_SUMA_60_61_55_56 = ",14,17,18,19,23,24,25,26,27,28,29,35,44,45,46,50,51,52,53,271,"
Const IMP_RESTA_60_61_55_56 = ",15,30,301,31,32,321,33,331,34,341,36,361,37,371,38,39,41,47,48,481,"


Const ES_IVARETENIDO = ",36,48,32,37,30,33,31,34,47,38,41,39,15,"
Const NO_ES_IVARETENIDO = ",24,25,26,271,27,28,29,19,17,18,23,44,46,45,14,35,50,51,52,53,"

Const ES_DOC_EXPORT = ",101,104,106,110,111,112,"


Dim lTipoLib As Integer

'Nro;Tipo Doc;Tipo Compra;RUT Proveedor;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;
Dim lIdDoc As Long
Dim lCodDocDTESII As String
Dim lTipoDoc As String
Dim lDTE As Boolean
Dim lDelGiro As Boolean
Dim lRUTEntidad As String
Dim lRazonSocialEntidad As String
Dim lIdEntidad As Long
Dim lNumDoc As String
Dim lFechaEmision As Long
Dim lFechaRec As Long
Dim lFechaAcuse As Long
Dim lFechaReclamo As Long
Dim lDescrip As String

Dim lMovEdited As Boolean
      
'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
Dim lMontoExento As Double
Dim lMontoAfecto As Double
Dim lMontoIVARec As Double
Dim lMontoIVANoRec As Double
Dim lCodIVANoRec As Integer

'para ventas
Dim lMontoIVA As Double
Dim lMontoIVARetTotal As Double     'IVA Retenido total
Dim lMontoIVARetParcial As Double   'IVA Retenido parciaL
Dim lMontoIVAPropio As Double       'IVA Propio
Dim lMontoIVATerceros As Double     'IVA Terceros

Dim lCredEmpContructora As Double
Dim lMontoTotalPeriodo As Double

Dim lIdSucursal As Long


'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
Dim lMontoTotal As Double
Dim lMontoNetoActFijo As Double
Dim lMontoIVAActFijo As Double
Dim lMontoIVAUsoComun As Double

      
'Impto. Sin Derecho a Credito;IVA No Retenido;
Dim lMontoImpSinDerechoCred As Double
Dim lMontoIVANoRet As Double

'Otros Impuestos
Dim lMontoOtroImp As Double
Dim lCodOtroImp As String
Dim lTasaOtroImp As Single

'Docuemtno de Referencia
Dim lTipoDocRef As Integer
Dim lNumDocRef As String
Dim lIdDocRef As Long
Dim lDTEDocRef As Integer
   

      
'Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;
Dim lTabacosPuros As Double
Dim lTabacosCigarrillos As Double
Dim lTabacosElaborados As Double
      
'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
Dim lNCE_NDE_FactCompra As String
Dim lOtrosImp(15) As ImpAdicDoc_t

'cuentas default
Dim lIdCtaAfecto As Long
Dim lIdCtaExento As Long
Dim lIdCtaTotal As Long

Dim lIdCtaIVA As Long
Dim lIdCtaIVAIrrec As Long
Dim lIdCtaOtrosImp As Long
Dim lIdCtaOtrosImpFacCompra As Long

'cuentas Entidad
Dim lIdCtaAfectoEntidad As Long
Dim lIdCtaExentoEntidad As Long
Dim lIdCtaTotalEntidad As Long
Dim lDefCtasProveedor As Boolean   'indoca si el proveedor tiene una cuenta definida distinta que la de omisión

'area de negocio Entidad
Dim lIdAreaNegAfectoEntidad As Long
Dim lIdAreaNegExentoEntidad As Long
Dim lIdAreaNegTotalEntidad As Long

'centro de costo entidad
Dim lIdCCostoAfectoEntidad As Long
Dim lIdCCostoExentoEntidad As Long
Dim lIdCCostoTotalEntidad As Long

Dim vError As Long ' indica si se prodruce el error de existencia de documento tema 1 2738156


Public Function Import_LibroComprasSII(Frm As Form, ByVal FName As String, Ano As Integer, Mes As Integer) As Boolean
   Dim i As Integer, j As Integer, k As Integer, n As Integer, r As Integer, l As Integer, F As Integer
   Dim p As Long
   Dim FNameLogImp As String
   Dim FNameTmp As String
   Dim MaxRegLibComp As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Buf As String
   Dim CorrelativoDoc As Integer
   Dim NewDoc As Boolean
   Dim Txt As String
   Dim DocErr As Integer
   Dim NDocsConError As Integer, NDocsOK As Integer
   Dim MsgDocsOK As String
   Dim RazonSocial As String, NotValidRut As Boolean
   Dim FechaAcuse8Dias As Long
   Dim MsgAcuse As Boolean
   Dim MsgTot As String
   Dim Rc As Integer
   Dim IdTipoValLib As Integer
   Dim Lineas() As String
   Dim Resp As String
   
   lTipoLib = LIB_COMPRAS
       
   Import_LibroComprasSII = False   'error
   On Error Resume Next
   
   FNameLogImp = gImportPath & "\Log\ImpLibCompSII-" & Format(Now, "yyyymmdd") & ".log"
   
   MaxRegLibComp = 3501
   On Error Resume Next
   
   '******************** ADO 2678537 Validación RC y RV (Victor Morales) 15-11-2021 *******************
   
   Lineas = Split(FName, "_")

   If UBound(Lineas) = 4 Then
      For i = LBound(Lineas) To UBound(Lineas)
        Select Case i

            Case 1:
                    If Mid(UCase(Lineas(i)), 1, 3) <> "COM" Then
                        If MsgBox1("El archivo no es de COMPRA es de " & Lineas(1) & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Exit Function
                        End If
                    End If
            Case 3:
                    If Mid(Lineas(i), 1, Len(Lineas(i)) - 2) <> gEmpresa.Rut Or Mid(Lineas(i), Len(Lineas(i)), 1) <> DV_Rut(gEmpresa.Rut) Then
                        If MsgBox1("El RUT del archivo capturado no coincide con la empresa que se está trabajando" & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Exit Function
                        End If
                    End If
            Case 4:
                    If Mid(Lineas(i), 1, 4) <> gEmpresa.Ano Or Mid(Lineas(i), 5, 2) <> Mes Then
                        If MsgBox1("La fecha " & Mid(Lineas(i), 1, 4) & Mid(Lineas(i), 5, 2) & " del archivo no corresponde al periodo que se está capturando " & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Exit Function
                        End If
                    End If
            Case Else:

        End Select
       Next
   Else

        If MsgBox1("Archivo de captura no cumple con requisitos de formato " & vbCrLf & "Rut " & gEmpresa.Rut & vbCrLf & "Mes " & Mes & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
        End If
   End If
   
   '****************************************************************************************************
   
   Rc = LineCount(FName, MaxRegLibComp)
   If ERR Then
      MsgErr FName
      Import_LibroComprasSII = -ERR
      Exit Function
   End If
   
   If gDbType = SQL_ACCESS Then
    If Rc < 0 Then
       MsgBox1 "El archivo tiene más registros de los que soporta el sistema para el libro de compras o ventas (máx. " & MaxRegLibComp - 1 & " Filas)." & vbCrLf & vbCrLf & "No es posible realizar el proceso de importación.", vbExclamation
       Exit Function
    End If
   End If
   
   If Not LoadCuentasDef Then
      Exit Function
   End If
   
   
   'convertimos el archivo de Unix a DOS porque así lo entrega el SII
   Rc = ConvUnix2DosFile(FName, FName & "_")
   
   If Rc < 0 Then    'hubo un error al leer el archivo
      Exit Function
   ElseIf Rc = 0 Then   ' no lo convirtió porque ya está en formato DOS, se usa archivo original
      FNameTmp = ""
   ElseIf Rc > 0 Then   'lo convirtió y generó arcgivo temporal con guión bajo al final. Leemos del archivo temporal y luego lo borramos
      FName = FName & "_"
      FNameTmp = FName
   End If
      
   'abrimos el archivo
   Fd = FreeFile
   Open FName For Input As #Fd
   If ERR Then
      MsgErr FName
      Import_LibroComprasSII = -ERR
      If FNameTmp <> "" Then   'generamos uno temporal
         Kill FNameTmp            'lo borramos
      End If
      Exit Function
   End If
      
   r = 0
   Sep = ";"
   NewDoc = False
   NDocsOK = 0
   CorrelativoDoc = 0
   MsgAcuse = False
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
         
   'Campos: Nro;Tipo Doc;Tipo Compra;RUT Proveedor;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;Impto. Sin Derecho a Credito;IVA No Retenido;Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         
   Do Until EOF(Fd)
         
      Line Input #Fd, Buf
      l = l + 1
               
      Buf = Trim(Buf)
      
      '1er registro con nombres de campos
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 Then
         GoTo NextRec
      End If
      
      p = 1
   
      'ahora leemos los documentos y los insertamos uno por uno
      n = vFmt(NextField2(Buf, p, Sep))         'número correlativo de registro. Viene en blanco si continúa documento anterior
      If n = 0 Then
         If CorrelativoDoc = 0 Then
            Call AddLogImp(FNameLogImp, FName, l, "Falta número de registro.")
            DocErr = True
         Else   'es continuación del documento anterior
            NewDoc = False
            r = r + 1         'registro i-esimo del documento NReg
         End If
         
      Else                    'documento nuevo
      
         If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
            If Not DocErr Then
               If ValidaTotalesCompras(MsgTot) Then
                  If GenDocumento Then       'puede que no se agregue si ya existe
                     NDocsOK = NDocsOK + 1
                  Else
                     Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: Documento ya existe.")
                  End If
               Else
                  DocErr = True
                  NDocsConError = NDocsConError + 1
                  Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, MsgTot)
                                                         
               End If
            Else
               NDocsConError = NDocsConError + 1
            End If
      
         End If
         
         DocErr = False
         CorrelativoDoc = n
         r = 0                'primer registro del documento NReg
         
         'Inicializaciones
               
         'Nro;Tipo Doc;Tipo Compra;RUT Proveedor;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;
         lCodDocDTESII = ""
         lTipoDoc = 0
         lDTE = False
         lDelGiro = 0
         lRUTEntidad = ""
         lRazonSocialEntidad = ""
         lIdEntidad = 0
         RazonSocial = ""
         NotValidRut = False
         lNumDoc = ""
         lFechaEmision = 0
         lFechaRec = 0
         lFechaAcuse = 0
         FechaAcuse8Dias = 0
         
         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
         lMontoExento = 0
         lMontoAfecto = 0
         lMontoIVARec = 0
         lMontoIVANoRec = 0
         lCodIVANoRec = 0
         
         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
         lMontoTotal = 0
         lMontoNetoActFijo = 0
         lMontoIVAActFijo = 0
         lMontoIVAUsoComun = 0
         
         'Impto. Sin Derecho a Credito;IVA No Retenido;
         lMontoImpSinDerechoCred = 0
         lMontoIVANoRet = 0
         
         lMontoOtroImp = 0   'totalizado para encabezado del doc
         
         'Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;
         lTabacosPuros = 0
         lTabacosCigarrillos = 0
         lTabacosElaborados = 0
         
         lIdDocRef = 0
         lTipoDocRef = 0
         lNumDocRef = ""
         lDTEDocRef = 0

         
         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         lNCE_NDE_FactCompra = ""
         
         For k = 0 To UBound(lOtrosImp)
            lOtrosImp(k).CodSIIDTE = ""
            lOtrosImp(k).Tasa = 0
            lOtrosImp(k).Valor = 0
         Next k
         
'         lIVANoRec = False
'         For k = 0 To UBound(lCodIVANoRec)
'            lCodIVANoRec(k) = 0
'            lMontoIVANoRec(k) = 0
'         Next k
   
      End If
        
        
      If r = 0 Then  'nuevo documento
      
         lCodDocDTESII = Trim(NextField2(Buf, p, Sep))
         If lCodDocDTESII = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta el tipo de documento.")
            DocErr = True
         End If
         
         lDTE = True
         lTipoDoc = GetTipoDocFromCodDocDTESII(lTipoLib, lCodDocDTESII)
         If lTipoDoc = 0 Then
            lDTE = False
            If lCodDocDTESII = "30" Then
               lTipoDoc = FindTipoDoc(lTipoLib, "FAC")
            ElseIf lCodDocDTESII = "32" Then
               lTipoDoc = FindTipoDoc(lTipoLib, "FCE")
            Else
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Tipo de documento inválido.")
               DocErr = True
            End If
         End If
         
         
         'tipo de compra: Del Giro
         lDelGiro = IIf(Trim(NextField2(Buf, p, Sep)) = "Del Giro", True, False)
         
         'RUT Proveedor
         lRUTEntidad = Trim(NextField2(Buf, p, Sep))
         If Not ValidRut(lRUTEntidad) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "RUT inválido.")
            DocErr = True
         End If
            
         'Razón Social
         lRazonSocialEntidad = Trim(NextField2(Buf, p, Sep))
         If lRazonSocialEntidad = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta razón social.")
            DocErr = True
         End If
         
         lRazonSocialEntidad = Utf82Ansi(Utf82Ansi(lRazonSocialEntidad))  ' se hace dos veces porque al leerlo como texto duplica los caracteres (raro)
                 
         lIdEntidad = GetIdEntidad(lRUTEntidad, RazonSocial, NotValidRut)
         
         'Folio
         lNumDoc = Trim(NextField2(Buf, p, Sep))
         If lNumDoc = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta el número de documento.")
            DocErr = True
         End If
   
         'Fecha Docto (Fecha Emision)
         lFechaEmision = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         If lFechaEmision = 0 Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: Falta la fecha del Documento.")
            DocErr = True
         End If
         
         
         'Fecha Recepcion (esta fecha es de recepción en el SII y no nos sirve)
         lFechaRec = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
'         If Month(lFechaRec) <> Mes Or Year(lFechaRec) <> Ano Then
'            Call AddLogImp(FNameLogImp, fname, l, "Fecha de recepción no corresponde al mes-año seleccionado.")
'            DocErr = True
'         End If
         
         'Fecha Acuse  (esta será la fecha de recepción para nuestro sistema, de acuerdo a lo indicado por Thomson)
         lFechaAcuse = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         If lFechaAcuse = 0 Then
            
            If Not MsgAcuse Then
               If MsgBox1("Se han detectado que uno o más documentos no poseen fecha de acuse de recibo, por lo que podrían existir documentos que deberían ser contabilizados en otro mes." & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  GoTo EndFnc
               End If
               MsgAcuse = True
            End If
            
            '8 días para el acuse
            FechaAcuse8Dias = DateAdd("d", 8, lFechaRec)
            If month(FechaAcuse8Dias) <> Mes Or Year(FechaAcuse8Dias) <> Ano Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "El documento no puede ser contabilizado en el mes-año seleccionado, ya que la fecha de acuse de recibo debería corresponder a otro mes.")
               DocErr = True
            End If
            
            lFechaAcuse = FechaAcuse8Dias   'dejamos esta como fecha de acuse

            
         ElseIf month(lFechaAcuse) <> Mes Or Year(lFechaAcuse) <> Ano Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "La fecha del acuse de recibo no corresponde al mes-año seleccionado.")
            DocErr = True
         End If
              
         
         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
         lMontoExento = vFmt(NextField2(Buf, p, Sep))
         lMontoAfecto = vFmt(NextField2(Buf, p, Sep))
         lMontoIVARec = vFmt(NextField2(Buf, p, Sep))
         lMontoIVANoRec = vFmt(NextField2(Buf, p, Sep))
         lCodIVANoRec = vFmt(NextField2(Buf, p, Sep))
         
         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
         
         
         
         lMontoTotal = vFmt(NextField2(Buf, p, Sep))
         
        
         
         lMontoNetoActFijo = vFmt(NextField2(Buf, p, Sep))
         lMontoIVAActFijo = vFmt(NextField2(Buf, p, Sep))
         lMontoIVAUsoComun = vFmt(NextField2(Buf, p, Sep))
         
         lMontoAfecto = lMontoAfecto - lMontoNetoActFijo    'Victor Morales 14 nov 2019
         lMontoIVARec = lMontoIVARec - lMontoIVAActFijo
         
         'Impto. Sin Derecho a Credito;IVA No Retenido;
         lMontoImpSinDerechoCred = vFmt(NextField2(Buf, p, Sep))
         lMontoIVANoRet = vFmt(NextField2(Buf, p, Sep))
         
         'Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;
         lTabacosPuros = vFmt(NextField2(Buf, p, Sep))
         lTabacosCigarrillos = vFmt(NextField2(Buf, p, Sep))
         lTabacosElaborados = vFmt(NextField2(Buf, p, Sep))
         
         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         lNCE_NDE_FactCompra = Trim(NextField2(Buf, p, Sep))
            
      Else
      
         If lCodDocDTESII <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza tipo de documento con datos documento actual.")
            DocErr = True
         End If
         
         'tipo de compra: Del Giro
         Txt = NextField2(Buf, p, Sep)
         
         'RUT Proveedor
         If lRUTEntidad <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza RUT con datos documento actual.")
            DocErr = True
         End If
            
         'Razón Social
         If lRazonSocialEntidad <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Razón Social con datos documento actual.")
            DocErr = True
         End If
         
         'Folio
         If lNumDoc <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Nro. Documento con datos documento actual.")
            DocErr = True
         End If
   
         'Fecha Docto
         If lFechaEmision <> Int(GetDate(Trim(NextField2(Buf, p, Sep)))) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Fecha Documento con datos documento actual.")
            DocErr = True
         End If
         
         For F = 8 To 24
            Txt = NextField2(Buf, p, Sep)   'saltamos los campos hasta llegar a los otros impuestos codificados
         Next F
     
      End If
      
      lOtrosImp(r).CodSIIDTE = Trim(NextField2(Buf, p, Sep))
      lOtrosImp(r).Valor = vFmt(NextField2(Buf, p, Sep))
      lOtrosImp(r).Tasa = vFmt(NextField2(Buf, p, Sep))
      
      If lOtrosImp(r).CodSIIDTE <> "" And lOtrosImp(r).CodSIIDTE <> "0" Then
         IdTipoValLib = GetTipoTipoValLibFromCodSIIDTE(lTipoLib, lOtrosImp(r).CodSIIDTE)
         If IdTipoValLib <= 0 Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Código otro impuesto inválido.")
            DocErr = True
         
         ElseIf Not ValidaOtroImp(lOtrosImp(r).CodSIIDTE) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Código de impuesto inválido.")
            DocErr = True
         End If
      End If
      
NextRec:

   Loop
   
   Close #Fd

   
   'agregamos el último documento, si hay
   If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
      If Not DocErr Then
          If ValidaTotalesCompras(MsgTot) Then
            If GenDocumento Then       'puede que no se agregue si ya existe
               NDocsOK = NDocsOK + 1
            Else
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Documento ya existe.")
            End If
         Else
            DocErr = True
            NDocsConError = NDocsConError + 1
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, MsgTot)
         End If
      Else
         NDocsConError = NDocsConError + 1
      End If
   
   End If
   
EndFnc:

   If NDocsOK > 1 Then
      MsgDocsOK = "Se importaron " & NDocsOK & " documentos nuevos."
   ElseIf NDocsOK = 1 Then
      MsgDocsOK = "Se importó un documento nuevo."
   Else
      MsgDocsOK = "No se importaron documentos nuevos."
   End If
      
   If NDocsConError = 0 Then
      MsgBox1 "Proceso de importación finalizado." & vbNewLine & vbNewLine & MsgDocsOK, vbInformation + vbOKOnly
      
   Else
      If NDocsConError = 1 Then
         MsgBox1 "Se encontró un documento con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
      Else
         MsgBox1 "Se encontraron " & NDocsConError & " documentos con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
      End If
      
      If MsgBox1("¿Desea revisar el log de importación " & FNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Frm.hWnd, "open", FNameLogImp, "", "", SW_SHOW)
      End If
      
   End If
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
   
   If FNameTmp <> "" Then   'generamos uno temporal
      Kill FNameTmp            'lo borramos
   End If

End Function
'2862611
'SIN DLL
Public Function Import_LibroComprasSIIAuto(Frm As Form, ByVal FName As String, Ano As Integer, Mes As Integer, Info() As String) As Boolean
   Dim i As Integer, j As Integer, k As Integer, n As Integer, r As Integer, l As Integer, F As Integer
   Dim p As Long
   Dim FNameLogImp As String
   Dim FNameTmp As String
   Dim MaxRegLibComp As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Buf As String
   Dim CorrelativoDoc As Integer
   Dim NewDoc As Boolean
   Dim Txt As String
   Dim DocErr As Integer
   Dim NDocsConError As Integer, NDocsOK As Integer
   Dim MsgDocsOK As String
   Dim RazonSocial As String, NotValidRut As Boolean
   Dim FechaAcuse8Dias As Long
   Dim MsgAcuse As Boolean
   Dim MsgTot As String
   Dim Rc As Integer
   Dim IdTipoValLib As Integer
   Dim Lineas() As String
   Dim Resp As String
   'Dim CSII As TR_CCONECTSII.TR_CCONECTSII
   'Dim Info() As String
   
   'Set CSII = New TR_CCONECTSII.TR_CCONECTSII
   'Info = CSII.GetCompras("11108309", "6", "3tres", "2022", "5")
   'Info = CSII.GetBoletasRecibidas("77765060", "2", "romal1", "2022", "10")
   'Info = CSII.GetBoletasEmitidas("77765060", "2", "romal1", "2022", "10")
   'Info = CSII.GetCompras(vFmtCID(Rut), DV_Rut(vFmtCID(Rut)), Clave, Ano, Mes)
   'Info = CSII.GetBoletasEmitidas("17533256", "1", "Fpriet4512", "2018", "05")
   
   lTipoLib = LIB_COMPRAS
       
   Import_LibroComprasSIIAuto = False   'error
   On Error Resume Next
   
   FNameLogImp = gImportPath & "\Log\ImpLibCompSII-" & Format(Now, "yyyymmdd") & ".log"
   
   MaxRegLibComp = 3501
   On Error Resume Next
   
   '******************** ADO 2678537 Validación RC y RV (Victor Morales) 15-11-2021 *******************
   
'   Lineas = Split(fname, "_")

'   If UBound(Lineas) = 4 Then
'      For i = LBound(Lineas) To UBound(Lineas)
'        Select Case i
'
'            Case 1:
'                    If Mid(UCase(Lineas(i)), 1, 3) <> "COM" Then
'                        If MsgBox1("El archivo no es de COMPRA es de " & Lineas(1) & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                            Exit Function
'                        End If
'                    End If
'            Case 3:
'                    If Mid(Lineas(i), 1, Len(Lineas(i)) - 2) <> gEmpresa.rut Or Mid(Lineas(i), Len(Lineas(i)), 1) <> DV_Rut(gEmpresa.rut) Then
'                        If MsgBox1("El RUT del archivo capturado no coincide con la empresa que se está trabajando" & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                            Exit Function
'                        End If
'                    End If
'            Case 4:
'                    If Mid(Lineas(i), 1, 4) <> gEmpresa.Ano Or Mid(Lineas(i), 5, 2) <> Mes Then
'                        If MsgBox1("La fecha " & Mid(Lineas(i), 1, 4) & Mid(Lineas(i), 5, 2) & " del archivo no corresponde al periodo que se está capturando " & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                            Exit Function
'                        End If
'                    End If
'            Case Else:
'
'        End Select
'       Next
'   Else
'
'        If MsgBox1("Archivo de captura no cumple con requisitos de formato " & vbCrLf & "Rut " & gEmpresa.rut & vbCrLf & "Mes " & Mes & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'            Exit Function
'        End If
'   End If
   
   '****************************************************************************************************
   
'   Rc = LineCount(fname, MaxRegLibComp)
'   If ERR Then
'      MsgErr fname
'      Import_LibroComprasSIIAuto = -ERR
'      Exit Function
'   End If
   
   If gDbType = SQL_ACCESS Then
    If MaxRegLibComp < UBound(Info) Then
       MsgBox1 "El archivo tiene más registros de los que soporta el sistema para el libro de compras o ventas (máx. " & MaxRegLibComp - 1 & " Filas)." & vbCrLf & vbCrLf & "No es posible realizar el proceso de importación.", vbExclamation
       Exit Function
    End If
   End If
   
   If Not LoadCuentasDef Then
      Exit Function
   End If
   
   
   'convertimos el archivo de Unix a DOS porque así lo entrega el SII
'   Rc = ConvUnix2DosFile(fname, fname & "_")
'
'   If Rc < 0 Then    'hubo un error al leer el archivo
'      Exit Function
'   ElseIf Rc = 0 Then   ' no lo convirtió porque ya está en formato DOS, se usa archivo original
'      FNameTmp = ""
'   ElseIf Rc > 0 Then   'lo convirtió y generó arcgivo temporal con guión bajo al final. Leemos del archivo temporal y luego lo borramos
'      fname = fname & "_"
'      FNameTmp = fname
'   End If
'
'   'abrimos el archivo
'   Fd = FreeFile
'   Open fname For Input As #Fd
'   If ERR Then
'      MsgErr fname
'      Import_LibroComprasSIIAuto = -ERR
'      If FNameTmp <> "" Then   'generamos uno temporal
'         Kill FNameTmp            'lo borramos
'      End If
'      Exit Function
'   End If
      
   r = 0
   Sep = ";"
   NewDoc = False
   NDocsOK = 0
   CorrelativoDoc = 0
   MsgAcuse = False
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
         
   'Campos: Nro;Tipo Doc;Tipo Compra;RUT Proveedor;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;Impto. Sin Derecho a Credito;IVA No Retenido;Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         
         
   For i = LBound(Info) To UBound(Info)
   'Do Until EOF(Fd)
         
      'Line Input #Fd, Buf
      l = LBound(Info)
               
      Buf = Trim(Info(i))
      
'      '1er registro con nombres de campos
'      If Buf = "" Then
'         GoTo NextRec
'      ElseIf l = 1 Then
'         GoTo NextRec
'      End If
      
      p = 1
   
      'ahora leemos los documentos y los insertamos uno por uno
      n = vFmt(NextField2(Buf, p, Sep))         'número correlativo de registro. Viene en blanco si continúa documento anterior
      If n = 33 Then
         n = 33
      End If
      If n = 0 Then
         If CorrelativoDoc = 0 Then
            Call AddLogImp(FNameLogImp, FName, l, "Falta número de registro.")
            DocErr = True
         Else   'es continuación del documento anterior
            NewDoc = False
            r = r + 1         'registro i-esimo del documento NReg
         End If
         
      Else                    'documento nuevo
      
         If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
            If Not DocErr Then
               If ValidaTotalesCompras(MsgTot) Then
                  If GenDocumento Then       'puede que no se agregue si ya existe
                     NDocsOK = NDocsOK + 1
                  Else
                     Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: Documento ya existe.")
                  End If
               Else
                  DocErr = True
                  NDocsConError = NDocsConError + 1
                  Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, MsgTot)
                                                         
               End If
            Else
               NDocsConError = NDocsConError + 1
            End If
      
         End If
         
         DocErr = False
         CorrelativoDoc = n
         r = 0                'primer registro del documento NReg
         
         'Inicializaciones
               
         'Nro;Tipo Doc;Tipo Compra;RUT Proveedor;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;
         lCodDocDTESII = ""
         lTipoDoc = 0
         lDTE = False
         lDelGiro = 0
         lRUTEntidad = ""
         lRazonSocialEntidad = ""
         lIdEntidad = 0
         RazonSocial = ""
         NotValidRut = False
         lNumDoc = ""
         lFechaEmision = 0
         lFechaRec = 0
         lFechaAcuse = 0
         FechaAcuse8Dias = 0
         
         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
         lMontoExento = 0
         lMontoAfecto = 0
         lMontoIVARec = 0
         lMontoIVANoRec = 0
         lCodIVANoRec = 0
         
         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
         lMontoTotal = 0
         lMontoNetoActFijo = 0
         lMontoIVAActFijo = 0
         lMontoIVAUsoComun = 0
         
         'Impto. Sin Derecho a Credito;IVA No Retenido;
         lMontoImpSinDerechoCred = 0
         lMontoIVANoRet = 0
         
         lMontoOtroImp = 0   'totalizado para encabezado del doc
         
         'Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;
         lTabacosPuros = 0
         lTabacosCigarrillos = 0
         lTabacosElaborados = 0
         
         lIdDocRef = 0
         lTipoDocRef = 0
         lNumDocRef = ""
         lDTEDocRef = 0

         
         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         lNCE_NDE_FactCompra = ""
         
         For k = 0 To UBound(lOtrosImp)
            lOtrosImp(k).CodSIIDTE = ""
            lOtrosImp(k).Tasa = 0
            lOtrosImp(k).Valor = 0
         Next k
         
'         lIVANoRec = False
'         For k = 0 To UBound(lCodIVANoRec)
'            lCodIVANoRec(k) = 0
'            lMontoIVANoRec(k) = 0
'         Next k
   
      End If
        
        
      If r = 0 Then  'nuevo documento
      
         lCodDocDTESII = Trim(NextField2(Buf, p, Sep))
         If lCodDocDTESII = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta el tipo de documento.")
            DocErr = True
         End If
         
         lDTE = True
         lTipoDoc = GetTipoDocFromCodDocDTESII(lTipoLib, lCodDocDTESII)
         If lTipoDoc = 0 Then
            lDTE = False
            If lCodDocDTESII = "30" Then
               lTipoDoc = FindTipoDoc(lTipoLib, "FAC")
            ElseIf lCodDocDTESII = "32" Then
               lTipoDoc = FindTipoDoc(lTipoLib, "FCE")
            Else
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Tipo de documento inválido.")
               DocErr = True
            End If
         End If
         
         
         'tipo de compra: Del Giro
         lDelGiro = IIf(Trim(NextField2(Buf, p, Sep)) = "Del Giro", True, False)
         
         'RUT Proveedor
         lRUTEntidad = Trim(NextField2(Buf, p, Sep))
         If Not ValidRut(lRUTEntidad) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "RUT inválido.")
            DocErr = True
         End If
            
         'Razón Social
         lRazonSocialEntidad = Trim(NextField2(Buf, p, Sep))
         If lRazonSocialEntidad = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta razón social.")
            DocErr = True
         End If
         
         lRazonSocialEntidad = Utf82Ansi(Utf82Ansi(lRazonSocialEntidad))  ' se hace dos veces porque al leerlo como texto duplica los caracteres (raro)
                 
         lIdEntidad = GetIdEntidad(lRUTEntidad, RazonSocial, NotValidRut)
         
         'Folio
         lNumDoc = Trim(NextField2(Buf, p, Sep))
         If lNumDoc = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta el número de documento.")
            DocErr = True
         End If
   
         'Fecha Docto (Fecha Emision)
         lFechaEmision = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         If lFechaEmision = 0 Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: Falta la fecha del Documento.")
            DocErr = True
         End If
         
         
         'Fecha Recepcion (esta fecha es de recepción en el SII y no nos sirve)
         lFechaRec = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
'         If Month(lFechaRec) <> Mes Or Year(lFechaRec) <> Ano Then
'            Call AddLogImp(FNameLogImp, fname, l, "Fecha de recepción no corresponde al mes-año seleccionado.")
'            DocErr = True
'         End If
         
         'Fecha Acuse  (esta será la fecha de recepción para nuestro sistema, de acuerdo a lo indicado por Thomson)
         lFechaAcuse = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         If lFechaAcuse = 0 Then
            
            If Not MsgAcuse Then
               If MsgBox1("Se han detectado que uno o más documentos no poseen fecha de acuse de recibo, por lo que podrían existir documentos que deberían ser contabilizados en otro mes." & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  GoTo EndFnc
               End If
               MsgAcuse = True
            End If
            
            '8 días para el acuse
            FechaAcuse8Dias = DateAdd("d", 8, lFechaRec)
            If month(FechaAcuse8Dias) <> Mes Or Year(FechaAcuse8Dias) <> Ano Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "El documento no puede ser contabilizado en el mes-año seleccionado, ya que la fecha de acuse de recibo debería corresponder a otro mes.")
               DocErr = True
            End If
            
            lFechaAcuse = FechaAcuse8Dias   'dejamos esta como fecha de acuse

            
         ElseIf month(lFechaAcuse) <> Mes Or Year(lFechaAcuse) <> Ano Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "La fecha del acuse de recibo no corresponde al mes-año seleccionado.")
            DocErr = True
         End If
              
         
         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
         lMontoExento = vFmt(NextField2(Buf, p, Sep))
         lMontoAfecto = vFmt(NextField2(Buf, p, Sep))
         lMontoIVARec = vFmt(NextField2(Buf, p, Sep))
         lMontoIVANoRec = vFmt(NextField2(Buf, p, Sep))
         lCodIVANoRec = vFmt(NextField2(Buf, p, Sep))
         
         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
         lMontoTotal = vFmt(NextField2(Buf, p, Sep))
         lMontoNetoActFijo = vFmt(NextField2(Buf, p, Sep))
         lMontoIVAActFijo = vFmt(NextField2(Buf, p, Sep))
         lMontoIVAUsoComun = vFmt(NextField2(Buf, p, Sep))
         
         lMontoAfecto = lMontoAfecto - lMontoNetoActFijo    'Victor Morales 14 nov 2019
         lMontoIVARec = lMontoIVARec - lMontoIVAActFijo
         
         'Impto. Sin Derecho a Credito;IVA No Retenido;
         lMontoImpSinDerechoCred = vFmt(NextField2(Buf, p, Sep))
         lMontoIVANoRet = vFmt(NextField2(Buf, p, Sep))
         
         'Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;
         lTabacosPuros = vFmt(NextField2(Buf, p, Sep))
         lTabacosCigarrillos = vFmt(NextField2(Buf, p, Sep))
         lTabacosElaborados = vFmt(NextField2(Buf, p, Sep))
         
         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         lNCE_NDE_FactCompra = Trim(NextField2(Buf, p, Sep))
            
      Else
      
         If lCodDocDTESII <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza tipo de documento con datos documento actual.")
            DocErr = True
         End If
         
         'tipo de compra: Del Giro
         Txt = NextField2(Buf, p, Sep)
         
         'RUT Proveedor
         If lRUTEntidad <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza RUT con datos documento actual.")
            DocErr = True
         End If
            
         'Razón Social
         If lRazonSocialEntidad <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Razón Social con datos documento actual.")
            DocErr = True
         End If
         
         'Folio
         If lNumDoc <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Nro. Documento con datos documento actual.")
            DocErr = True
         End If
   
         'Fecha Docto
         If lFechaEmision <> Int(GetDate(Trim(NextField2(Buf, p, Sep)))) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Fecha Documento con datos documento actual.")
            DocErr = True
         End If
         
         For F = 8 To 24
            Txt = NextField2(Buf, p, Sep)   'saltamos los campos hasta llegar a los otros impuestos codificados
         Next F
     
      End If
      
      lOtrosImp(r).CodSIIDTE = Trim(NextField2(Buf, p, Sep))
      lOtrosImp(r).Valor = vFmt(NextField2(Buf, p, Sep))
      lOtrosImp(r).Tasa = vFmt(NextField2(Buf, p, Sep))
      
      If lOtrosImp(r).CodSIIDTE <> "" And lOtrosImp(r).CodSIIDTE <> "0" Then
         IdTipoValLib = GetTipoTipoValLibFromCodSIIDTE(lTipoLib, lOtrosImp(r).CodSIIDTE)
         If IdTipoValLib <= 0 Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Código otro impuesto inválido.")
            DocErr = True
         
         ElseIf Not ValidaOtroImp(lOtrosImp(r).CodSIIDTE) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Código de impuesto inválido.")
            DocErr = True
         End If
      End If
      
'NextRec:

   Next
   'Loop
   
   Close #Fd

   
   'agregamos el último documento, si hay
   If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
      If Not DocErr Then
          If ValidaTotalesCompras(MsgTot) Then
            If GenDocumento Then       'puede que no se agregue si ya existe
               NDocsOK = NDocsOK + 1
            Else
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Documento ya existe.")
            End If
         Else
            DocErr = True
            NDocsConError = NDocsConError + 1
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, MsgTot)
         End If
      Else
         NDocsConError = NDocsConError + 1
      End If
   
   End If
   
EndFnc:

   If NDocsOK > 1 Then
      MsgDocsOK = "Se importaron " & NDocsOK & " documentos nuevos."
   ElseIf NDocsOK = 1 Then
      MsgDocsOK = "Se importó un documento nuevo."
   Else
      MsgDocsOK = "No se importaron documentos nuevos."
   End If
      
   If NDocsConError = 0 Then
      MsgBox1 "Proceso de importación finalizado." & vbNewLine & vbNewLine & MsgDocsOK, vbInformation + vbOKOnly
      
   Else
      If NDocsConError = 1 Then
         MsgBox1 "Se encontró un documento con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
      Else
         MsgBox1 "Se encontraron " & NDocsConError & " documentos con errores en la integración con el SII." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
      End If
      
      If MsgBox1("¿Desea revisar el log de importación " & FNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Frm.hWnd, "open", FNameLogImp, "", "", SW_SHOW)
      End If
      
   End If
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
   
   If FNameTmp <> "" Then   'generamos uno temporal
      Kill FNameTmp            'lo borramos
   End If

End Function
'Fin 2862611
'2862611
' CON DLL Descomentar interior ademas de agregar la DLL a los componentes
Public Function Import_LibroComprasSIIAuto1(Frm As Form, ByVal FName As String, Ano As Integer, Mes As Integer, Rut As String, Clave As String) As Boolean
'   Dim i As Integer, j As Integer, k As Integer, n As Integer, r As Integer, l As Integer, F As Integer
'   Dim p As Long
'   Dim FNameLogImp As String
'   Dim FNameTmp As String
'   Dim MaxRegLibComp As Integer
'   Dim Fd As Long
'   Dim Sep As String
'   Dim Buf As String
'   Dim CorrelativoDoc As Integer
'   Dim NewDoc As Boolean
'   Dim Txt As String
'   Dim DocErr As Integer
'   Dim NDocsConError As Integer, NDocsOK As Integer
'   Dim MsgDocsOK As String
'   Dim RazonSocial As String, NotValidRut As Boolean
'   Dim FechaAcuse8Dias As Long
'   Dim MsgAcuse As Boolean
'   Dim MsgTot As String
'   Dim Rc As Integer
'   Dim IdTipoValLib As Integer
'   Dim Lineas() As String
'   Dim Resp As String
'   Dim CSII As TR_CCONECTSII.TR_CCONECTSII
'   Dim Info() As String
'
'   Set CSII = New TR_CCONECTSII.TR_CCONECTSII
'   'Info = CSII.GetCompras("11108309", "6", "3tres", "2022", "5")
'   'Info = CSII.GetBoletasRecibidas("77765060", "2", "romal1", "2022", "10")
'   'Info = CSII.GetBoletasEmitidas("77765060", "2", "romal1", "2022", "10")
'   Info = CSII.GetCompras(vFmtCID(Rut), DV_Rut(vFmtCID(Rut)), Clave, Ano, Mes)
'   'Info = CSII.GetBoletasEmitidas("17533256", "1", "Fpriet4512", "2018", "05")
'
'   lTipoLib = LIB_COMPRAS
'
'   Import_LibroComprasSIIAuto1 = False   'error
'   On Error Resume Next
'
'   FNameLogImp = gImportPath & "\Log\ImpLibCompSII-" & Format(Now, "yyyymmdd") & ".log"
'
'   MaxRegLibComp = 3501
'   On Error Resume Next
'
'   '******************** ADO 2678537 Validación RC y RV (Victor Morales) 15-11-2021 *******************
'
'   Lineas = Split(fname, "_")
'
''   If UBound(Lineas) = 4 Then
''      For i = LBound(Lineas) To UBound(Lineas)
''        Select Case i
''
''            Case 1:
''                    If Mid(UCase(Lineas(i)), 1, 3) <> "COM" Then
''                        If MsgBox1("El archivo no es de COMPRA es de " & Lineas(1) & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
''                            Exit Function
''                        End If
''                    End If
''            Case 3:
''                    If Mid(Lineas(i), 1, Len(Lineas(i)) - 2) <> gEmpresa.rut Or Mid(Lineas(i), Len(Lineas(i)), 1) <> DV_Rut(gEmpresa.rut) Then
''                        If MsgBox1("El RUT del archivo capturado no coincide con la empresa que se está trabajando" & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
''                            Exit Function
''                        End If
''                    End If
''            Case 4:
''                    If Mid(Lineas(i), 1, 4) <> gEmpresa.Ano Or Mid(Lineas(i), 5, 2) <> Mes Then
''                        If MsgBox1("La fecha " & Mid(Lineas(i), 1, 4) & Mid(Lineas(i), 5, 2) & " del archivo no corresponde al periodo que se está capturando " & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
''                            Exit Function
''                        End If
''                    End If
''            Case Else:
''
''        End Select
''       Next
''   Else
''
''        If MsgBox1("Archivo de captura no cumple con requisitos de formato " & vbCrLf & "Rut " & gEmpresa.rut & vbCrLf & "Mes " & Mes & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
''            Exit Function
''        End If
''   End If
'
'   '****************************************************************************************************
'
''   Rc = LineCount(fname, MaxRegLibComp)
''   If ERR Then
''      MsgErr fname
''      Import_LibroComprasSIIAuto1 = -ERR
''      Exit Function
''   End If
'
'   If gDbType = SQL_ACCESS Then
'    If MaxRegLibComp < UBound(Info) Then
'       MsgBox1 "El archivo tiene más registros de los que soporta el sistema para el libro de compras o ventas (máx. " & MaxRegLibComp - 1 & " Filas)." & vbCrLf & vbCrLf & "No es posible realizar el proceso de importación.", vbExclamation
'       Exit Function
'    End If
'   End If
'
'   If Not LoadCuentasDef Then
'      Exit Function
'   End If
'
'
'   'convertimos el archivo de Unix a DOS porque así lo entrega el SII
''   Rc = ConvUnix2DosFile(fname, fname & "_")
''
''   If Rc < 0 Then    'hubo un error al leer el archivo
''      Exit Function
''   ElseIf Rc = 0 Then   ' no lo convirtió porque ya está en formato DOS, se usa archivo original
''      FNameTmp = ""
''   ElseIf Rc > 0 Then   'lo convirtió y generó arcgivo temporal con guión bajo al final. Leemos del archivo temporal y luego lo borramos
''      fname = fname & "_"
''      FNameTmp = fname
''   End If
''
''   'abrimos el archivo
''   Fd = FreeFile
''   Open fname For Input As #Fd
''   If ERR Then
''      MsgErr fname
''      Import_LibroComprasSIIAuto1 = -ERR
''      If FNameTmp <> "" Then   'generamos uno temporal
''         Kill FNameTmp            'lo borramos
''      End If
''      Exit Function
''   End If
'
'   r = 0
'   Sep = ";"
'   NewDoc = False
'   NDocsOK = 0
'   CorrelativoDoc = 0
'   MsgAcuse = False
'
'   If gDbType = SQL_SERVER Then
'      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
'   End If
'
'   'Campos: Nro;Tipo Doc;Tipo Compra;RUT Proveedor;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;Impto. Sin Derecho a Credito;IVA No Retenido;Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
'
'
'   For i = LBound(Info) To UBound(Info)
'   'Do Until EOF(Fd)
'
'      'Line Input #Fd, Buf
'      l = LBound(Info)
'
'      Buf = Trim(Info(i))
'
''      '1er registro con nombres de campos
''      If Buf = "" Then
''         GoTo NextRec
''      ElseIf l = 1 Then
''         GoTo NextRec
''      End If
'
'      p = 1
'
'      'ahora leemos los documentos y los insertamos uno por uno
'      n = vFmt(NextField2(Buf, p, Sep))         'número correlativo de registro. Viene en blanco si continúa documento anterior
'      If n = 33 Then
'         n = 33
'      End If
'      If n = 0 Then
'         If CorrelativoDoc = 0 Then
'            Call AddLogImp(FNameLogImp, fname, l, "Falta número de registro.")
'            DocErr = True
'         Else   'es continuación del documento anterior
'            NewDoc = False
'            r = r + 1         'registro i-esimo del documento NReg
'         End If
'
'      Else                    'documento nuevo
'
'         If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
'            If Not DocErr Then
'               If ValidaTotalesCompras(MsgTot) Then
'                  If GenDocumento Then       'puede que no se agregue si ya existe
'                     NDocsOK = NDocsOK + 1
'                  Else
'                     Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Advertencia: Documento ya existe.")
'                  End If
'               Else
'                  DocErr = True
'                  NDocsConError = NDocsConError + 1
'                  Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, MsgTot)
'
'               End If
'            Else
'               NDocsConError = NDocsConError + 1
'            End If
'
'         End If
'
'         DocErr = False
'         CorrelativoDoc = n
'         r = 0                'primer registro del documento NReg
'
'         'Inicializaciones
'
'         'Nro;Tipo Doc;Tipo Compra;RUT Proveedor;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;
'         lCodDocDTESII = ""
'         lTipoDoc = 0
'         lDTE = False
'         lDelGiro = 0
'         lRUTEntidad = ""
'         lRazonSocialEntidad = ""
'         lIdEntidad = 0
'         RazonSocial = ""
'         NotValidRut = False
'         lNumDoc = ""
'         lFechaEmision = 0
'         lFechaRec = 0
'         lFechaAcuse = 0
'         FechaAcuse8Dias = 0
'
'         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
'         lMontoExento = 0
'         lMontoAfecto = 0
'         lMontoIVARec = 0
'         lMontoIVANoRec = 0
'         lCodIVANoRec = 0
'
'         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
'         lMontoTotal = 0
'         lMontoNetoActFijo = 0
'         lMontoIVAActFijo = 0
'         lMontoIVAUsoComun = 0
'
'         'Impto. Sin Derecho a Credito;IVA No Retenido;
'         lMontoImpSinDerechoCred = 0
'         lMontoIVANoRet = 0
'
'         lMontoOtroImp = 0   'totalizado para encabezado del doc
'
'         'Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;
'         lTabacosPuros = 0
'         lTabacosCigarrillos = 0
'         lTabacosElaborados = 0
'
'         lIdDocRef = 0
'         lTipoDocRef = 0
'         lNumDocRef = ""
'         lDTEDocRef = 0
'
'
'         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
'         lNCE_NDE_FactCompra = ""
'
'         For k = 0 To UBound(lOtrosImp)
'            lOtrosImp(k).CodSIIDTE = ""
'            lOtrosImp(k).Tasa = 0
'            lOtrosImp(k).Valor = 0
'         Next k
'
''         lIVANoRec = False
''         For k = 0 To UBound(lCodIVANoRec)
''            lCodIVANoRec(k) = 0
''            lMontoIVANoRec(k) = 0
''         Next k
'
'      End If
'
'
'      If r = 0 Then  'nuevo documento
'
'         lCodDocDTESII = Trim(NextField2(Buf, p, Sep))
'         If lCodDocDTESII = "" Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Falta el tipo de documento.")
'            DocErr = True
'         End If
'
'         lDTE = True
'         lTipoDoc = GetTipoDocFromCodDocDTESII(lTipoLib, lCodDocDTESII)
'         If lTipoDoc = 0 Then
'            lDTE = False
'            If lCodDocDTESII = "30" Then
'               lTipoDoc = FindTipoDoc(lTipoLib, "FAC")
'            ElseIf lCodDocDTESII = "32" Then
'               lTipoDoc = FindTipoDoc(lTipoLib, "FCE")
'            Else
'               Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Tipo de documento inválido.")
'               DocErr = True
'            End If
'         End If
'
'
'         'tipo de compra: Del Giro
'         lDelGiro = IIf(Trim(NextField2(Buf, p, Sep)) = "Del Giro", True, False)
'
'         'RUT Proveedor
'         lRUTEntidad = Trim(NextField2(Buf, p, Sep))
'         If Not ValidRut(lRUTEntidad) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "RUT inválido.")
'            DocErr = True
'         End If
'
'         'Razón Social
'         lRazonSocialEntidad = Trim(NextField2(Buf, p, Sep))
'         If lRazonSocialEntidad = "" Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Falta razón social.")
'            DocErr = True
'         End If
'
'         lRazonSocialEntidad = Utf82Ansi(Utf82Ansi(lRazonSocialEntidad))  ' se hace dos veces porque al leerlo como texto duplica los caracteres (raro)
'
'         lIdEntidad = GetIdEntidad(lRUTEntidad, RazonSocial, NotValidRut)
'
'         'Folio
'         lNumDoc = Trim(NextField2(Buf, p, Sep))
'         If lNumDoc = "" Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Falta el número de documento.")
'            DocErr = True
'         End If
'
'         'Fecha Docto (Fecha Emision)
'         lFechaEmision = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
'         If lFechaEmision = 0 Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Advertencia: Falta la fecha del Documento.")
'            DocErr = True
'         End If
'
'
'         'Fecha Recepcion (esta fecha es de recepción en el SII y no nos sirve)
'         lFechaRec = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
''         If Month(lFechaRec) <> Mes Or Year(lFechaRec) <> Ano Then
''            Call AddLogImp(FNameLogImp, fname, l, "Fecha de recepción no corresponde al mes-año seleccionado.")
''            DocErr = True
''         End If
'
'         'Fecha Acuse  (esta será la fecha de recepción para nuestro sistema, de acuerdo a lo indicado por Thomson)
'         lFechaAcuse = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
'         If lFechaAcuse = 0 Then
'
'            If Not MsgAcuse Then
'               If MsgBox1("Se han detectado que uno o más documentos no poseen fecha de acuse de recibo, por lo que podrían existir documentos que deberían ser contabilizados en otro mes." & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                  GoTo EndFnc
'               End If
'               MsgAcuse = True
'            End If
'
'            '8 días para el acuse
'            FechaAcuse8Dias = DateAdd("d", 8, lFechaRec)
'            If month(FechaAcuse8Dias) <> Mes Or Year(FechaAcuse8Dias) <> Ano Then
'               Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "El documento no puede ser contabilizado en el mes-año seleccionado, ya que la fecha de acuse de recibo debería corresponder a otro mes.")
'               DocErr = True
'            End If
'
'            lFechaAcuse = FechaAcuse8Dias   'dejamos esta como fecha de acuse
'
'
'         ElseIf month(lFechaAcuse) <> Mes Or Year(lFechaAcuse) <> Ano Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "La fecha del acuse de recibo no corresponde al mes-año seleccionado.")
'            DocErr = True
'         End If
'
'
'         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
'         lMontoExento = vFmt(NextField2(Buf, p, Sep))
'         lMontoAfecto = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVARec = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVANoRec = vFmt(NextField2(Buf, p, Sep))
'         lCodIVANoRec = vFmt(NextField2(Buf, p, Sep))
'
'         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
'         lMontoTotal = vFmt(NextField2(Buf, p, Sep))
'         lMontoNetoActFijo = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVAActFijo = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVAUsoComun = vFmt(NextField2(Buf, p, Sep))
'
'         lMontoAfecto = lMontoAfecto - lMontoNetoActFijo    'Victor Morales 14 nov 2019
'         lMontoIVARec = lMontoIVARec - lMontoIVAActFijo
'
'         'Impto. Sin Derecho a Credito;IVA No Retenido;
'         lMontoImpSinDerechoCred = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVANoRet = vFmt(NextField2(Buf, p, Sep))
'
'         'Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;
'         lTabacosPuros = vFmt(NextField2(Buf, p, Sep))
'         lTabacosCigarrillos = vFmt(NextField2(Buf, p, Sep))
'         lTabacosElaborados = vFmt(NextField2(Buf, p, Sep))
'
'         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
'         lNCE_NDE_FactCompra = Trim(NextField2(Buf, p, Sep))
'
'      Else
'
'         If lCodDocDTESII <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza tipo de documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         'tipo de compra: Del Giro
'         Txt = NextField2(Buf, p, Sep)
'
'         'RUT Proveedor
'         If lRUTEntidad <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza RUT con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Razón Social
'         If lRazonSocialEntidad <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Razón Social con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Folio
'         If lNumDoc <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Nro. Documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Fecha Docto
'         If lFechaEmision <> Int(GetDate(Trim(NextField2(Buf, p, Sep)))) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Fecha Documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         For F = 8 To 24
'            Txt = NextField2(Buf, p, Sep)   'saltamos los campos hasta llegar a los otros impuestos codificados
'         Next F
'
'      End If
'
'      lOtrosImp(r).CodSIIDTE = Trim(NextField2(Buf, p, Sep))
'      lOtrosImp(r).Valor = vFmt(NextField2(Buf, p, Sep))
'      lOtrosImp(r).Tasa = vFmt(NextField2(Buf, p, Sep))
'
'      If lOtrosImp(r).CodSIIDTE <> "" And lOtrosImp(r).CodSIIDTE <> "0" Then
'         IdTipoValLib = GetTipoTipoValLibFromCodSIIDTE(lTipoLib, lOtrosImp(r).CodSIIDTE)
'         If IdTipoValLib <= 0 Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Código otro impuesto inválido.")
'            DocErr = True
'
'         ElseIf Not ValidaOtroImp(lOtrosImp(r).CodSIIDTE) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Código de impuesto inválido.")
'            DocErr = True
'         End If
'      End If
'
''NextRec:
'
'   Next
'   'Loop
'
'   Close #Fd
'
'
'   'agregamos el último documento, si hay
'   If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
'      If Not DocErr Then
'          If ValidaTotalesCompras(MsgTot) Then
'            If GenDocumento Then       'puede que no se agregue si ya existe
'               NDocsOK = NDocsOK + 1
'            Else
'               Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Documento ya existe.")
'            End If
'         Else
'            DocErr = True
'            NDocsConError = NDocsConError + 1
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, MsgTot)
'         End If
'      Else
'         NDocsConError = NDocsConError + 1
'      End If
'
'   End If
'
'EndFnc:
'
'   If NDocsOK > 1 Then
'      MsgDocsOK = "Se importaron " & NDocsOK & " documentos nuevos."
'   ElseIf NDocsOK = 1 Then
'      MsgDocsOK = "Se importó un documento nuevo."
'   Else
'      MsgDocsOK = "No se importaron documentos nuevos."
'   End If
'
'   If NDocsConError = 0 Then
'      MsgBox1 "Proceso de importación finalizado." & vbNewLine & vbNewLine & MsgDocsOK, vbInformation + vbOKOnly
'
'   Else
'      If NDocsConError = 1 Then
'         MsgBox1 "Se encontró un documento con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
'      Else
'         MsgBox1 "Se encontraron " & NDocsConError & " documentos con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
'      End If
'
'      If MsgBox1("¿Desea revisar el log de importación " & FNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'         Call ShellExecute(Frm.hWnd, "open", FNameLogImp, "", "", SW_SHOW)
'      End If
'
'   End If
'
'   If gDbType = SQL_SERVER Then
'      Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
'   End If
'
'   If FNameTmp <> "" Then   'generamos uno temporal
'      Kill FNameTmp            'lo borramos
'   End If

End Function
'Fin 2862611
'2862611
'SIN DLL
Public Function Import_LibroVentasSIIAuto(Frm As Form, ByVal FName As String, Ano As Integer, Mes As Integer, Info() As String) As Boolean
   Dim i As Integer, j As Integer, k As Integer, n As Integer, r As Integer, l As Integer, F As Integer
   Dim p As Long
   Dim FNameLogImp As String
   Dim FNameTmp As String
   Dim MaxRegLibComp As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Buf As String
   Dim CorrelativoDoc As Integer
   Dim NewDoc As Boolean
   Dim Txt As String
   Dim DocErr As Integer
   Dim NDocsConError As Integer, NDocsOK As Integer
   Dim MsgDocsOK As String
   Dim RazonSocial As String, NotValidRut As Boolean
   Dim FechaAcuse8Dias As Long
   Dim MsgAcuse As Boolean
   Dim MsgTot As String
   Dim Rc As Integer
   Dim AuxStr As String, AuxMonto As Double
   Dim CodSucursal As String
   Dim IdTipoValLib As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Lineas() As String
   Dim Resp As String

   vError = 0
   
   lTipoLib = LIB_VENTAS
       
   Import_LibroVentasSIIAuto = False   'error
   On Error Resume Next
   
   FNameLogImp = gImportPath & "\Log\ImpLibCompSII-" & Format(Now, "yyyymmdd") & ".log"
   
   MaxRegLibComp = 3501
   On Error Resume Next
   
   
      '******************** ADO 2678537 Validación RC y RV (Victor Morales) 15-11-2021 *******************
   
   Lineas = Split(FName, "_")

'   If UBound(Lineas) = 3 Then
'      For i = LBound(Lineas) To UBound(Lineas)
'        Select Case i
'
'            Case 1:
'                    If Mid(UCase(Lineas(i)), 1, 3) <> "VEN" Then
'                        If MsgBox1("El archivo no es de VENTA es de " & Lineas(1) & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                            Exit Function
'                        End If
'                    End If
'            Case 2:
'                    If Mid(Lineas(i), 1, Len(Lineas(i)) - 2) <> gEmpresa.Rut Or Mid(Lineas(i), Len(Lineas(i)), 1) <> DV_Rut(gEmpresa.Rut) Then
'                        If MsgBox1("El RUT del archivo capturado no coincide con la empresa que se está trabajando" & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                            Exit Function
'                        End If
'                    End If
'            Case 3:
'                    If Mid(Lineas(i), 1, 4) <> gEmpresa.Ano Or Mid(Lineas(i), 5, 2) <> Mes Then
'                        If MsgBox1("La fecha " & Mid(Lineas(i), 1, 4) & Mid(Lineas(i), 5, 2) & " del archivo no corresponde al periodo que se está capturando " & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                            Exit Function
'                        End If
'                    End If
'            Case Else:
'
'        End Select
'       Next
'   Else
'        If MsgBox1("Archivo de captura no cumple con requisitos de formato " & vbCrLf & "Rut " & gEmpresa.Rut & vbCrLf & "Mes " & Mes & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'            Exit Function
'        End If
'   End If
'
'   '****************************************************************************************************
'
'   Rc = LineCount(fname, MaxRegLibComp)
'   If ERR Then
'      MsgErr fname
'      Import_LibroVentasSIIAuto = -ERR
'      Exit Function
'   End If
   
   If gDbType = SQL_ACCESS Then
    If MaxRegLibComp < UBound(Info) Then
       MsgBox1 "El archivo tiene más registros de los que soporta el sistema para el libro de compras o ventas (máx. " & MaxRegLibComp - 1 & " Filas)." & vbCrLf & vbCrLf & "No es posible realizar el proceso de importación.", vbExclamation
       Exit Function
    End If
   End If
   
   If Not LoadCuentasDef Then
      Exit Function
   End If
   
   
'   'convertimos el archivo de Unix a DOS porque así lo entrega el SII
'   Rc = ConvUnix2DosFile(fname, fname & "_")
'
'   If Rc < 0 Then    'hubo un error al leer el archivo
'      Exit Function
'   ElseIf Rc = 0 Then   ' no lo convirtió porque ya está en formato DOS, se usa archivo original
'      FNameTmp = ""
'   ElseIf Rc > 0 Then   'lo convirtió y generó arcgivo temporal con guión bajo al final. Leemos del archivo temporal y luego lo borramos
'      fname = fname & "_"
'      FNameTmp = fname
'   End If
'
'   'abrimos el archivo
'   Fd = FreeFile
'   Open fname For Input As #Fd
'   If ERR Then
'      MsgErr fname
'      Import_LibroVentasSIIAuto = -ERR
'      If FNameTmp <> "" Then   'generamos uno temporal
'         Kill FNameTmp            'lo borramos
'      End If
'      Exit Function
'   End If
      
   r = 0
   Sep = ";"
   NewDoc = False
   NDocsOK = 0
   CorrelativoDoc = 0
   MsgAcuse = False
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
         
   'Campos: Nro;Tipo Doc;Tipo Compra;RUT Cliente;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;Impto. Sin Derecho a Credito;IVA No Retenido;Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         
   For i = LBound(Info) To UBound(Info)
   'Do Until EOF(Fd)
         
      'Line Input #Fd, Buf
      l = LBound(Info)
               
      Buf = Trim(Info(i))
      'Buf = Trim(Buf)
      
'      '1er registro con nombres de campos
'      If Buf = "" Then
'         GoTo NextRec
'      ElseIf l = 1 Then
'         GoTo NextRec
'      End If
      
      p = 1
   
      'ahora leemos los documentos y los insertamos uno por uno
      n = vFmt(NextField2(Buf, p, Sep))         'número correlativo de registro. Viene en blanco si continúa documento anterior
      If n = 0 Then
         If CorrelativoDoc = 0 Then
            Call AddLogImp(FNameLogImp, FName, l, "Falta número de registro.")
            DocErr = True
         Else   'es continuación del documento anterior
            NewDoc = False
            r = r + 1         'registro i-esimo del documento NReg
         End If
         
      Else                    'documento nuevo
      
         If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
            If Not DocErr Then
               If ValidaTotalesVentas(MsgTot) Then
                  If GenDocumento Then       'puede que no se agregue si ya existe
                     NDocsOK = NDocsOK + 1
                     If MsgTot <> "" Then
                        Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: " & MsgTot)
                     End If
                  Else
                     Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: Documento ya existe.")
                  End If
               Else
                  DocErr = True
                  NDocsConError = NDocsConError + 1
                  Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, MsgTot)
                                                         
               End If
            Else
               NDocsConError = NDocsConError + 1
            End If
      
         End If
         
         DocErr = False
         CorrelativoDoc = n
         r = 0                'primer registro del documento NReg
         
         'Inicializaciones
               
         'Nro;Tipo Doc;Tipo Compra;RUT Cliente;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;
         lCodDocDTESII = ""
         lTipoDoc = 0
         lDTE = False
         lDelGiro = 0
         lRUTEntidad = ""
         lRazonSocialEntidad = ""
         lIdEntidad = 0
         RazonSocial = ""
         NotValidRut = False
         lNumDoc = ""
         lFechaEmision = 0
         lFechaRec = 0
         lFechaAcuse = 0
         lFechaReclamo = 0
         FechaAcuse8Dias = 0
         lIdSucursal = 0
         
         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
         lMontoExento = 0
         lMontoAfecto = 0
         lMontoIVA = 0
         
         lMontoIVARetTotal = 0    'IVA Retenido total
         lMontoIVARetParcial = 0  'IVA Retenido parciaL
         lMontoIVANoRet = 0       'IVA NO Retenido
         lMontoIVAPropio = 0      'IVA Propio
         lMontoIVATerceros = 0    'IVA Terceros
         
         lCredEmpContructora = 0
         lMontoTotalPeriodo = 0
         
         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
         lMontoTotal = 0
         
         lTipoDocRef = 0
         lNumDocRef = ""
         lIdDocRef = 0
         lDTEDocRef = 0
   
                     
         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         lNCE_NDE_FactCompra = ""
         
         For k = 0 To UBound(lOtrosImp)
            lOtrosImp(k).CodSIIDTE = ""
            lOtrosImp(k).Tasa = 0
            lOtrosImp(k).Valor = 0
         Next k
         
      End If
        
        
      If r = 0 Then  'nuevo documento
      
         lCodDocDTESII = Trim(NextField2(Buf, p, Sep))
         If lCodDocDTESII = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta el tipo de documento.")
            DocErr = True
         End If
         
         lDTE = True
         lTipoDoc = GetTipoDocFromCodDocDTESII(lTipoLib, lCodDocDTESII)
         If lTipoDoc = 0 Then
            lDTE = False
            lTipoDoc = GetTipoDocFromCodDocSII(lTipoLib, lCodDocDTESII)
            If lTipoDoc = 0 Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Tipo de documento inválido.")
               DocErr = True
            End If
         End If
         
         
         'tipo de compra: Del Giro
         lDelGiro = IIf(Trim(NextField2(Buf, p, Sep)) = "Del Giro", True, False)
         
         'RUT Cliente
         lRUTEntidad = Trim(NextField2(Buf, p, Sep))
         If InStr(ES_DOC_EXPORT, "," & lCodDocDTESII & ",") <= 0 Then             'FCA - 12/10/2021
            If Not ValidRut(lRUTEntidad) Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "RUT inválido.")
               DocErr = True
            End If
         End If
            
         'Razón Social
         lRazonSocialEntidad = Trim(NextField2(Buf, p, Sep))
         If lRazonSocialEntidad = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta razón social.")
            DocErr = True
         End If
         
         lRazonSocialEntidad = Utf82Ansi(Utf82Ansi(lRazonSocialEntidad))  ' se hace dos veces porque al leerlo como texto duplica los caracteres (raro)
                 
         lIdEntidad = GetIdEntidad(lRUTEntidad, RazonSocial, NotValidRut)
         
         'Folio
         lNumDoc = Trim(NextField2(Buf, p, Sep))
         If lNumDoc = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta el número de documento.")
            DocErr = True
         End If
   
         'Fecha Docto (Fecha Emision)
         lFechaEmision = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         If lFechaEmision = 0 Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: Falta la fecha del Documento.")
            DocErr = True
         End If
         
         'Fecha Recepcion (esta fecha es de recepción en el SII y no nos sirve)
         lFechaRec = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         'no Aplica
         
         'Fecha Acuse  (esta será la fecha de recepción para nuestro sistema, de acuerdo a lo indicado por Thomson)
         'no Aplica
         lFechaAcuse = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         
         'Fecha Reclamo
         'No Aplica
         lFechaReclamo = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
              
         
         'Monto Exento;Monto Neto;Monto IVA ;Monto Total
         lMontoExento = vFmt(NextField2(Buf, p, Sep))
         lMontoAfecto = vFmt(NextField2(Buf, p, Sep))
         lMontoIVA = vFmt(NextField2(Buf, p, Sep))
         lMontoTotal = vFmt(NextField2(Buf, p, Sep))
         
         'Monto Exento;Monto Neto;Monto IVA Retenido Total ;Monto IVA Retenido Parcial; IVA no Retenido;
         lMontoIVARetTotal = vFmt(NextField2(Buf, p, Sep))
         lMontoIVARetParcial = vFmt(NextField2(Buf, p, Sep))
         
         'IVA Propio; IVA Terceros
         'No aplica por ahora  (A18, A19)
         lMontoIVANoRet = vFmt(NextField2(Buf, p, Sep))
         lMontoIVANoRet = 0    'por cálculos en GenMovDocumento
         lMontoIVAPropio = vFmt(NextField2(Buf, p, Sep))
         lMontoIVATerceros = vFmt(NextField2(Buf, p, Sep))
         
         'Rut emisor Liq. Factura; Neto Comisión Liq. Factura; Exento Comisión Liq. Factura; IVA Comisión Liq. Factura;
         'No aplican por ahora (A20 -> A23)
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         
         'IVA fuera de plazo;
         'No aplica por ahora (A25 -> A28)
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         
         'TipoDoc Referencia; Folio Doc Referencia;
         AuxStr = Trim(NextField2(Buf, p, Sep))
         If Val(AuxStr) <> 0 Then
            lTipoDocRef = GetTipoDocFromCodDocDTESII(lTipoLib, AuxStr)
         End If

         lNumDocRef = vFmt(NextField2(Buf, p, Sep))
         
         If lTipoDocRef <> 0 And lNumDocRef <> "" Then
         
         'pipe Tema 1 2738156
         
         
            Q1 = "SELECT IdDoc, DTE FROM Documento "
            Q1 = Q1 & " WHERE NumDoc = '" & lNumDocRef & "'"
            Q1 = Q1 & " AND TipoLib = " & LIB_VENTAS & " AND TipoDoc = " & lTipoDocRef
            '2800560
            Q1 = Q1 & " AND idempresa =" & gEmpresa.id
            'fin 2800560

            Set Rs = OpenRs(DbMain, Q1)

            If Not Rs.EOF Then
               lIdDocRef = vFld(Rs("IdDoc"))
               lDTEDocRef = vFld(Rs("DTE"))

            Else
                 'pipe Tema 1 2738156
                             
                If vError = 0 Then
                             
                        If MsgBox1("Existen Documentos que no tienen documento de referencia ¿Desea agregar los documentos de igual forma?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                           vError = 1
                        Else
                           vError = 2
                           Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "El documento de referencia no ha sido ingresado al sistema.")
                           DocErr = True
                           
                        End If
                 ElseIf vError = 2 Then
                   Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "El documento de referencia no ha sido ingresado al sistema.")
                   DocErr = True
                  End If
            End If

            Call CloseRs(Rs)
         'fin
         End If
            
         
         'Num. Ident. Receptor Extranjero; Nacionalidad Receptor Extranjero
         'No aplica por ahora (A28)
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))

         'Credito empresa constructora (A29)
         lCredEmpContructora = vFmt(NextField2(Buf, p, Sep))
         
         'Impto. Zona Franca (Ley 18211); Garantia Dep. Envases; Indicador Venta sin Costo; Indicador Servicio Periodico; Monto No facturable
         'No aplican por ahora (A30 -> A34)
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         
         'Total Monto Periodo (A35)
         lMontoTotalPeriodo = vFmt(NextField2(Buf, p, Sep))
         
         'Venta Pasajes Transporte Nacional; Venta Pasajes Transporte Internacional; Numero Interno
         'No aplican por ahora (A36 -> A38)
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
                 
         'Codigo Sucursal
         CodSucursal = Trim(NextField2(Buf, p, Sep))
         lIdSucursal = GetIdSucursal(CodSucursal)
         
         
         'NCE o NDE sobre Fact. de Compra;
         lNCE_NDE_FactCompra = Trim(NextField2(Buf, p, Sep))
            
         'Codigo Otro Impuesto; Valor Otro Impuesto; Tasa Otro Impuesto
'         lCodOtroImp = Trim(NextField2(Buf, p, Sep))
'         lMontoOtroImp = vFmt(NextField2(Buf, p, Sep))
'         lTasaOtroImp = vFmt(NextField2(Buf, p, Sep))
         
         
         'Codigo Otro Impuesto; Valor Otro Impuesto; Tasa Otro Impuesto (r = 0)
         lOtrosImp(r).CodSIIDTE = Trim(NextField2(Buf, p, Sep))
         lOtrosImp(r).Valor = vFmt(NextField2(Buf, p, Sep))
         lOtrosImp(r).Tasa = vFmt(NextField2(Buf, p, Sep))
         
         If lCodOtroImp <> "" Then
            IdTipoValLib = GetTipoTipoValLibFromCodSIIDTE(lTipoLib, lCodOtroImp)
            If IdTipoValLib <= 0 Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Código otro impuesto inválido.")
               DocErr = True
            End If
         End If
          
      Else  'es continuación del registro anterior, nuevo documento

         If lCodDocDTESII <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza tipo de documento con datos documento actual.")
            DocErr = True
         End If

         'tipo de venta: Del Giro
         Txt = NextField2(Buf, p, Sep)

         'RUT Cliente
         If lRUTEntidad <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza RUT con datos documento actual.")
            DocErr = True
         End If

         'Razón Social
         If lRazonSocialEntidad <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Razón Social con datos documento actual.")
            DocErr = True
         End If

         'Folio
         If lNumDoc <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Nro. Documento con datos documento actual.")
            DocErr = True
         End If

         'Fecha Docto
         If lFechaEmision <> Int(GetDate(Trim(NextField2(Buf, p, Sep)))) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Fecha Documento con datos documento actual.")
            DocErr = True
         End If

         For F = 8 To 40
            Txt = NextField2(Buf, p, Sep)   'saltamos los campos hasta llegar a los otros impuestos codificados
         Next F
           
         lOtrosImp(r).CodSIIDTE = Trim(NextField2(Buf, p, Sep))
         lOtrosImp(r).Valor = vFmt(NextField2(Buf, p, Sep))
         lOtrosImp(r).Tasa = vFmt(NextField2(Buf, p, Sep))
   
         If lOtrosImp(r).CodSIIDTE <> "" And lOtrosImp(r).CodSIIDTE <> "0" Then
            If Not ValidaOtroImp(lOtrosImp(r).CodSIIDTE) Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Código de impuesto inválido.")
               DocErr = True
            End If
         End If
         
      End If
      
      
'NextRec:
    Next
'   Loop
   
   Close #Fd

   
   'agregamos el último documento, si hay
   If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
      If Not DocErr Then
          If ValidaTotalesVentas(MsgTot) Then
            If GenDocumento Then       'puede que no se agregue si ya existe
               NDocsOK = NDocsOK + 1
            Else
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Documento ya existe.")
            End If
         Else
            DocErr = True
            NDocsConError = NDocsConError + 1
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, MsgTot)
         End If
      Else
         NDocsConError = NDocsConError + 1
      End If
   
   End If
   
EndFnc:

   

   If NDocsOK > 1 Then
      MsgDocsOK = "Se importaron " & NDocsOK & " documentos nuevos."
    
   ElseIf NDocsOK = 1 Then
      MsgDocsOK = "Se importó un documento nuevo."
      
   Else
      MsgDocsOK = "No se importaron documentos nuevos."
   End If
      
      
   If NDocsConError = 0 Then
      MsgBox1 "Proceso de importación finalizado." & vbNewLine & vbNewLine & MsgDocsOK, vbInformation + vbOKOnly
      
      '2907989
      If NDocsOK = 0 Then
        If MsgBox1("¿Desea revisar el log de importación " & FNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
           Call ShellExecute(Frm.hWnd, "open", FNameLogImp, "", "", SW_SHOW)
        End If
      End If
      
      '2907989
      
   Else
      If NDocsConError = 1 Then
         MsgBox1 "Se encontró un documento con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
      Else
         MsgBox1 "Se encontraron " & NDocsConError & " documentos con errores en la integración con el SII." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
      End If
      
      If MsgBox1("¿Desea revisar el log de importación " & FNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Frm.hWnd, "open", FNameLogImp, "", "", SW_SHOW)
      End If
      
   End If
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")   'para que no trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
   
   If FNameTmp <> "" Then   'generamos uno temporal
      Kill FNameTmp            'lo borramos
   End If
    
   

End Function

Public Function Import_LibroVentasSII(Frm As Form, ByVal FName As String, Ano As Integer, Mes As Integer) As Boolean
   Dim i As Integer, j As Integer, k As Integer, n As Integer, r As Integer, l As Integer, F As Integer
   Dim p As Long
   Dim FNameLogImp As String
   Dim FNameTmp As String
   Dim MaxRegLibComp As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Buf As String
   Dim CorrelativoDoc As Integer
   Dim NewDoc As Boolean
   Dim Txt As String
   Dim DocErr As Integer
   Dim NDocsConError As Integer, NDocsOK As Integer
   Dim MsgDocsOK As String
   Dim RazonSocial As String, NotValidRut As Boolean
   Dim FechaAcuse8Dias As Long
   Dim MsgAcuse As Boolean
   Dim MsgTot As String
   Dim Rc As Integer
   Dim AuxStr As String, AuxMonto As Double
   Dim CodSucursal As String
   Dim IdTipoValLib As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Lineas() As String
   Dim Resp As String
   
   'pipe 2738156 tema 1
   vError = 0
   
   lTipoLib = LIB_VENTAS
       
   Import_LibroVentasSII = False   'error
   On Error Resume Next
   
   FNameLogImp = gImportPath & "\Log\ImpLibVentSII-" & Format(Now, "yyyymmdd") & ".log"
   
   MaxRegLibComp = 3501
   On Error Resume Next
   
   
      '******************** ADO 2678537 Validación RC y RV (Victor Morales) 15-11-2021 *******************
   
   Lineas = Split(FName, "_")

   If UBound(Lineas) = 3 Then
      For i = LBound(Lineas) To UBound(Lineas)
        Select Case i

            Case 1:
                    If Mid(UCase(Lineas(i)), 1, 3) <> "VEN" Then
                        If MsgBox1("El archivo no es de VENTA es de " & Lineas(1) & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Exit Function
                        End If
                    End If
            Case 2:
                    If Mid(Lineas(i), 1, Len(Lineas(i)) - 2) <> gEmpresa.Rut Or Mid(Lineas(i), Len(Lineas(i)), 1) <> DV_Rut(gEmpresa.Rut) Then
                        If MsgBox1("El RUT del archivo capturado no coincide con la empresa que se está trabajando" & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Exit Function
                        End If
                    End If
            Case 3:
                    If Mid(Lineas(i), 1, 4) <> gEmpresa.Ano Or Mid(Lineas(i), 5, 2) <> Mes Then
                        If MsgBox1("La fecha " & Mid(Lineas(i), 1, 4) & Mid(Lineas(i), 5, 2) & " del archivo no corresponde al periodo que se está capturando " & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Exit Function
                        End If
                    End If
            Case Else:

        End Select
       Next
   Else
        If MsgBox1("Archivo de captura no cumple con requisitos de formato " & vbCrLf & "Rut " & gEmpresa.Rut & vbCrLf & "Mes " & Mes & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
        End If
   End If
   
   '****************************************************************************************************
   
   Rc = LineCount(FName, MaxRegLibComp)
   If ERR Then
      MsgErr FName
      Import_LibroVentasSII = -ERR
      Exit Function
   End If
   
   If gDbType = SQL_ACCESS Then
    If Rc < 0 Then
       MsgBox1 "El archivo tiene más registros de los que soporta el sistema para el libro de compras o ventas (máx. " & MaxRegLibComp - 1 & " Filas)." & vbCrLf & vbCrLf & "No es posible realizar el proceso de importación.", vbExclamation
       Exit Function
    End If
   End If
   
   If Not LoadCuentasDef Then
      Exit Function
   End If
   
   
   'convertimos el archivo de Unix a DOS porque así lo entrega el SII
   Rc = ConvUnix2DosFile(FName, FName & "_")
   
   If Rc < 0 Then    'hubo un error al leer el archivo
      Exit Function
   ElseIf Rc = 0 Then   ' no lo convirtió porque ya está en formato DOS, se usa archivo original
      FNameTmp = ""
   ElseIf Rc > 0 Then   'lo convirtió y generó arcgivo temporal con guión bajo al final. Leemos del archivo temporal y luego lo borramos
      FName = FName & "_"
      FNameTmp = FName
   End If
      
   'abrimos el archivo
   Fd = FreeFile
   Open FName For Input As #Fd
   If ERR Then
      MsgErr FName
      Import_LibroVentasSII = -ERR
      If FNameTmp <> "" Then   'generamos uno temporal
         Kill FNameTmp            'lo borramos
      End If
      Exit Function
   End If
      
   r = 0
   Sep = ";"
   NewDoc = False
   NDocsOK = 0
   CorrelativoDoc = 0
   MsgAcuse = False
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
         
   'Campos: Nro;Tipo Doc;Tipo Compra;RUT Cliente;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;Impto. Sin Derecho a Credito;IVA No Retenido;Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         
   Do Until EOF(Fd)
         
      Line Input #Fd, Buf
      l = l + 1
               
      Buf = Trim(Buf)
      
      '1er registro con nombres de campos
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 Then
         GoTo NextRec
      End If
      
      p = 1
   
      'ahora leemos los documentos y los insertamos uno por uno
      n = vFmt(NextField2(Buf, p, Sep))         'número correlativo de registro. Viene en blanco si continúa documento anterior
      If n = 0 Then
         If CorrelativoDoc = 0 Then
            Call AddLogImp(FNameLogImp, FName, l, "Falta número de registro.")
            DocErr = True
         Else   'es continuación del documento anterior
            NewDoc = False
            r = r + 1         'registro i-esimo del documento NReg
         End If
         
      Else                    'documento nuevo
      
         If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
            If Not DocErr Then
               If ValidaTotalesVentas(MsgTot) Then
                  If GenDocumento Then       'puede que no se agregue si ya existe
                     NDocsOK = NDocsOK + 1
                     If MsgTot <> "" Then
                        Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: " & MsgTot)
                     End If
                  Else
                     Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: Documento ya existe.")
                  End If
               Else
                  DocErr = True
                  NDocsConError = NDocsConError + 1
                  Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, MsgTot)
                                                         
               End If
            Else
               NDocsConError = NDocsConError + 1
            End If
      
         End If
         
         DocErr = False
         CorrelativoDoc = n
         r = 0                'primer registro del documento NReg
         
         'Inicializaciones
               
         'Nro;Tipo Doc;Tipo Compra;RUT Cliente;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;
         lCodDocDTESII = ""
         lTipoDoc = 0
         lDTE = False
         lDelGiro = 0
         lRUTEntidad = ""
         lRazonSocialEntidad = ""
         lIdEntidad = 0
         RazonSocial = ""
         NotValidRut = False
         lNumDoc = ""
         lFechaEmision = 0
         lFechaRec = 0
         lFechaAcuse = 0
         lFechaReclamo = 0
         FechaAcuse8Dias = 0
         lIdSucursal = 0
         
         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
         lMontoExento = 0
         lMontoAfecto = 0
         lMontoIVA = 0
         
         lMontoIVARetTotal = 0    'IVA Retenido total
         lMontoIVARetParcial = 0  'IVA Retenido parciaL
         lMontoIVANoRet = 0       'IVA NO Retenido
         lMontoIVAPropio = 0      'IVA Propio
         lMontoIVATerceros = 0    'IVA Terceros
         
         lCredEmpContructora = 0
         lMontoTotalPeriodo = 0
         
         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
         lMontoTotal = 0
         
         lTipoDocRef = 0
         lNumDocRef = ""
         lIdDocRef = 0
         lDTEDocRef = 0
   
                     
         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         lNCE_NDE_FactCompra = ""
         
         For k = 0 To UBound(lOtrosImp)
            lOtrosImp(k).CodSIIDTE = ""
            lOtrosImp(k).Tasa = 0
            lOtrosImp(k).Valor = 0
         Next k
         
      End If
        
        
      If r = 0 Then  'nuevo documento
      
         lCodDocDTESII = Trim(NextField2(Buf, p, Sep))
         If lCodDocDTESII = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta el tipo de documento.")
            DocErr = True
         End If
         
         lDTE = True
         lTipoDoc = GetTipoDocFromCodDocDTESII(lTipoLib, lCodDocDTESII)
         If lTipoDoc = 0 Then
            lDTE = False
            lTipoDoc = GetTipoDocFromCodDocSII(lTipoLib, lCodDocDTESII)
            If lTipoDoc = 0 Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Tipo de documento inválido.")
               DocErr = True
            End If
         End If
         
         
         'tipo de compra: Del Giro
         lDelGiro = IIf(Trim(NextField2(Buf, p, Sep)) = "Del Giro", True, False)
         
         'RUT Cliente
         lRUTEntidad = Trim(NextField2(Buf, p, Sep))
         If InStr(ES_DOC_EXPORT, "," & lCodDocDTESII & ",") <= 0 Then             'FCA - 12/10/2021
            If Not ValidRut(lRUTEntidad) Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "RUT inválido.")
               DocErr = True
            End If
         End If
            
         'Razón Social
         lRazonSocialEntidad = Trim(NextField2(Buf, p, Sep))
         If lRazonSocialEntidad = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta razón social.")
            DocErr = True
         End If
         
         lRazonSocialEntidad = Utf82Ansi(Utf82Ansi(lRazonSocialEntidad))  ' se hace dos veces porque al leerlo como texto duplica los caracteres (raro)
                 
         lIdEntidad = GetIdEntidad(lRUTEntidad, RazonSocial, NotValidRut)
         
         'Folio
         lNumDoc = Trim(NextField2(Buf, p, Sep))
         If lNumDoc = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta el número de documento.")
            DocErr = True
         End If
   
         'Fecha Docto (Fecha Emision)
         lFechaEmision = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         If lFechaEmision = 0 Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: Falta la fecha del Documento.")
            DocErr = True
         End If
         
         'Fecha Recepcion (esta fecha es de recepción en el SII y no nos sirve)
         lFechaRec = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         'no Aplica
         
         'Fecha Acuse  (esta será la fecha de recepción para nuestro sistema, de acuerdo a lo indicado por Thomson)
         'no Aplica
         lFechaAcuse = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         
         'Fecha Reclamo
         'No Aplica
         lFechaReclamo = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
              
         
         'Monto Exento;Monto Neto;Monto IVA ;Monto Total
         lMontoExento = vFmt(NextField2(Buf, p, Sep))
         lMontoAfecto = vFmt(NextField2(Buf, p, Sep))
         lMontoIVA = vFmt(NextField2(Buf, p, Sep))
         lMontoTotal = vFmt(NextField2(Buf, p, Sep))
         
         'Monto Exento;Monto Neto;Monto IVA Retenido Total ;Monto IVA Retenido Parcial; IVA no Retenido;
         lMontoIVARetTotal = vFmt(NextField2(Buf, p, Sep))
         lMontoIVARetParcial = vFmt(NextField2(Buf, p, Sep))
         
         'IVA Propio; IVA Terceros
         'No aplica por ahora  (A18, A19)
         lMontoIVANoRet = vFmt(NextField2(Buf, p, Sep))
         lMontoIVANoRet = 0    'por cálculos en GenMovDocumento
         lMontoIVAPropio = vFmt(NextField2(Buf, p, Sep))
         lMontoIVATerceros = vFmt(NextField2(Buf, p, Sep))
         
         'Rut emisor Liq. Factura; Neto Comisión Liq. Factura; Exento Comisión Liq. Factura; IVA Comisión Liq. Factura;
         'No aplican por ahora (A20 -> A23)
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         
         'IVA fuera de plazo;
         'No aplica por ahora (A25 -> A28)
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         
         'TipoDoc Referencia; Folio Doc Referencia;
         AuxStr = Trim(NextField2(Buf, p, Sep))
         If Val(AuxStr) <> 0 Then
            lTipoDocRef = GetTipoDocFromCodDocDTESII(lTipoLib, AuxStr)
         End If

         lNumDocRef = vFmt(NextField2(Buf, p, Sep))
         
         If lTipoDocRef <> 0 And lNumDocRef <> "" Then
         
         'pipe Tema 1 2738156
         
         
            Q1 = "SELECT IdDoc, DTE FROM Documento "
            Q1 = Q1 & " WHERE NumDoc = '" & lNumDocRef & "'"
            Q1 = Q1 & " AND TipoLib = " & LIB_VENTAS & " AND TipoDoc = " & lTipoDocRef
            '2800560
            Q1 = Q1 & " AND idempresa =" & gEmpresa.id
            'fin 2800560

            Set Rs = OpenRs(DbMain, Q1)

            If Not Rs.EOF Then
               lIdDocRef = vFld(Rs("IdDoc"))
               lDTEDocRef = vFld(Rs("DTE"))

            Else
                 'pipe Tema 1 2738156
                             
                If vError = 0 Then
                             
                        If MsgBox1("Existen Documentos en el archivo csv que no tienen documento de referencia ¿Desea agregar los documentos de igual forma?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                           vError = 1
                        Else
                           vError = 2
                           Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "El documento de referencia no ha sido ingresado al sistema.")
                           DocErr = True
                           
                        End If
                 ElseIf vError = 2 Then
                   Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "El documento de referencia no ha sido ingresado al sistema.")
                   DocErr = True
                  End If
            End If

            Call CloseRs(Rs)
         'fin
         End If
            
         
         'Num. Ident. Receptor Extranjero; Nacionalidad Receptor Extranjero
         'No aplica por ahora (A28)
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))

         'Credito empresa constructora (A29)
         lCredEmpContructora = vFmt(NextField2(Buf, p, Sep))
         
         'Impto. Zona Franca (Ley 18211); Garantia Dep. Envases; Indicador Venta sin Costo; Indicador Servicio Periodico; Monto No facturable
         'No aplican por ahora (A30 -> A34)
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         
         'Total Monto Periodo (A35)
         lMontoTotalPeriodo = vFmt(NextField2(Buf, p, Sep))
         
         'Venta Pasajes Transporte Nacional; Venta Pasajes Transporte Internacional; Numero Interno
         'No aplican por ahora (A36 -> A38)
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
                 
         'Codigo Sucursal
         CodSucursal = Trim(NextField2(Buf, p, Sep))
         lIdSucursal = GetIdSucursal(CodSucursal)
         
         
         'NCE o NDE sobre Fact. de Compra;
         lNCE_NDE_FactCompra = Trim(NextField2(Buf, p, Sep))
            
         'Codigo Otro Impuesto; Valor Otro Impuesto; Tasa Otro Impuesto
'         lCodOtroImp = Trim(NextField2(Buf, p, Sep))
'         lMontoOtroImp = vFmt(NextField2(Buf, p, Sep))
'         lTasaOtroImp = vFmt(NextField2(Buf, p, Sep))
         
         
         'Codigo Otro Impuesto; Valor Otro Impuesto; Tasa Otro Impuesto (r = 0)
         lOtrosImp(r).CodSIIDTE = Trim(NextField2(Buf, p, Sep))
         lOtrosImp(r).Valor = vFmt(NextField2(Buf, p, Sep))
         lOtrosImp(r).Tasa = vFmt(NextField2(Buf, p, Sep))
         
         If lCodOtroImp <> "" Then
            IdTipoValLib = GetTipoTipoValLibFromCodSIIDTE(lTipoLib, lCodOtroImp)
            If IdTipoValLib <= 0 Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Código otro impuesto inválido.")
               DocErr = True
            End If
         End If
          
      Else  'es continuación del registro anterior, nuevo documento

         If lCodDocDTESII <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza tipo de documento con datos documento actual.")
            DocErr = True
         End If

         'tipo de venta: Del Giro
         Txt = NextField2(Buf, p, Sep)

         'RUT Cliente
         If lRUTEntidad <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza RUT con datos documento actual.")
            DocErr = True
         End If

         'Razón Social
         If lRazonSocialEntidad <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Razón Social con datos documento actual.")
            DocErr = True
         End If

         'Folio
         If lNumDoc <> Trim(NextField2(Buf, p, Sep)) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Nro. Documento con datos documento actual.")
            DocErr = True
         End If

         'Fecha Docto
         If lFechaEmision <> Int(GetDate(Trim(NextField2(Buf, p, Sep)))) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "No calza Fecha Documento con datos documento actual.")
            DocErr = True
         End If

         For F = 8 To 40
            Txt = NextField2(Buf, p, Sep)   'saltamos los campos hasta llegar a los otros impuestos codificados
         Next F
           
         lOtrosImp(r).CodSIIDTE = Trim(NextField2(Buf, p, Sep))
         lOtrosImp(r).Valor = vFmt(NextField2(Buf, p, Sep))
         lOtrosImp(r).Tasa = vFmt(NextField2(Buf, p, Sep))
   
         If lOtrosImp(r).CodSIIDTE <> "" And lOtrosImp(r).CodSIIDTE <> "0" Then
            If Not ValidaOtroImp(lOtrosImp(r).CodSIIDTE) Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Código de impuesto inválido.")
               DocErr = True
            End If
         End If
         
      End If
      
      
NextRec:

   Loop
   
   Close #Fd

   
   'agregamos el último documento, si hay
   If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
      If Not DocErr Then
          If ValidaTotalesVentas(MsgTot) Then
            If GenDocumento Then       'puede que no se agregue si ya existe
               NDocsOK = NDocsOK + 1
            Else
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Documento ya existe.")
            End If
         Else
            DocErr = True
            NDocsConError = NDocsConError + 1
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, MsgTot)
         End If
      Else
         NDocsConError = NDocsConError + 1
      End If
   
   End If
   
EndFnc:

   

   If NDocsOK > 1 Then
      MsgDocsOK = "Se importaron " & NDocsOK & " documentos nuevos."
    
   ElseIf NDocsOK = 1 Then
      MsgDocsOK = "Se importó un documento nuevo."
      
   Else
      MsgDocsOK = "No se importaron documentos nuevos."
   End If
      
      
   If NDocsConError = 0 Then
      MsgBox1 "Proceso de importación finalizado." & vbNewLine & vbNewLine & MsgDocsOK, vbInformation + vbOKOnly
      
      '2907989
      If NDocsOK = 0 Then
        If MsgBox1("¿Desea revisar el log de importación " & FNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
           Call ShellExecute(Frm.hWnd, "open", FNameLogImp, "", "", SW_SHOW)
        End If
      End If
      
      '2907989
      
   Else
      If NDocsConError = 1 Then
         MsgBox1 "Se encontró un documento con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
      Else
         MsgBox1 "Se encontraron " & NDocsConError & " documentos con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
      End If
      
      If MsgBox1("¿Desea revisar el log de importación " & FNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Frm.hWnd, "open", FNameLogImp, "", "", SW_SHOW)
      End If
      
   End If
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")   'para que no trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
   
   If FNameTmp <> "" Then   'generamos uno temporal
      Kill FNameTmp            'lo borramos
   End If
    
   

End Function


'2862611
'CON DLL Descomentar y agrear DLL a los componentes
Public Function Import_LibroVentasSIIAuto1(Frm As Form, ByVal FName As String, Ano As Integer, Mes As Integer, Rut As String, Clave As String) As Boolean
'   Dim i As Integer, j As Integer, k As Integer, n As Integer, r As Integer, l As Integer, F As Integer
'   Dim p As Long
'   Dim FNameLogImp As String
'   Dim FNameTmp As String
'   Dim MaxRegLibComp As Integer
'   Dim Fd As Long
'   Dim Sep As String
'   Dim Buf As String
'   Dim CorrelativoDoc As Integer
'   Dim NewDoc As Boolean
'   Dim Txt As String
'   Dim DocErr As Integer
'   Dim NDocsConError As Integer, NDocsOK As Integer
'   Dim MsgDocsOK As String
'   Dim RazonSocial As String, NotValidRut As Boolean
'   Dim FechaAcuse8Dias As Long
'   Dim MsgAcuse As Boolean
'   Dim MsgTot As String
'   Dim Rc As Integer
'   Dim AuxStr As String, AuxMonto As Double
'   Dim CodSucursal As String
'   Dim IdTipoValLib As Integer
'   Dim Q1 As String
'   Dim Rs As Recordset
'   Dim Lineas() As String
'   Dim Resp As String
'   Dim CSII As TR_CCONECTSII.TR_CCONECTSII
'   Dim Info() As String
'
'   Set CSII = New TR_CCONECTSII.TR_CCONECTSII
'   'Info = CSII.GetVentas("11.108.309", "6", "3tres", "2022", "10")
'   Info = CSII.GetVentas(vFmtCID(Rut), DV_Rut(vFmtCID(Rut)), Clave, Ano, Mes)
'   'pipe 2738156 tema 1
'   vError = 0
'
'   lTipoLib = LIB_VENTAS
'
'   Import_LibroVentasSIIAuto1 = False   'error
'   On Error Resume Next
'
'   FNameLogImp = gImportPath & "\Log\ImpLibCompSII-" & Format(Now, "yyyymmdd") & ".log"
'
'   MaxRegLibComp = 3501
'   On Error Resume Next
'
'
'      '******************** ADO 2678537 Validación RC y RV (Victor Morales) 15-11-2021 *******************
'
'   Lineas = Split(fname, "_")
'
''   If UBound(Lineas) = 3 Then
''      For i = LBound(Lineas) To UBound(Lineas)
''        Select Case i
''
''            Case 1:
''                    If Mid(UCase(Lineas(i)), 1, 3) <> "VEN" Then
''                        If MsgBox1("El archivo no es de VENTA es de " & Lineas(1) & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
''                            Exit Function
''                        End If
''                    End If
''            Case 2:
''                    If Mid(Lineas(i), 1, Len(Lineas(i)) - 2) <> gEmpresa.Rut Or Mid(Lineas(i), Len(Lineas(i)), 1) <> DV_Rut(gEmpresa.Rut) Then
''                        If MsgBox1("El RUT del archivo capturado no coincide con la empresa que se está trabajando" & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
''                            Exit Function
''                        End If
''                    End If
''            Case 3:
''                    If Mid(Lineas(i), 1, 4) <> gEmpresa.Ano Or Mid(Lineas(i), 5, 2) <> Mes Then
''                        If MsgBox1("La fecha " & Mid(Lineas(i), 1, 4) & Mid(Lineas(i), 5, 2) & " del archivo no corresponde al periodo que se está capturando " & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
''                            Exit Function
''                        End If
''                    End If
''            Case Else:
''
''        End Select
''       Next
''   Else
''        If MsgBox1("Archivo de captura no cumple con requisitos de formato " & vbCrLf & "Rut " & gEmpresa.Rut & vbCrLf & "Mes " & Mes & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
''            Exit Function
''        End If
''   End If
''
''   '****************************************************************************************************
''
''   Rc = LineCount(fname, MaxRegLibComp)
''   If ERR Then
''      MsgErr fname
''      Import_LibroVentasSIIAuto1 = -ERR
''      Exit Function
''   End If
'
'   If gDbType = SQL_ACCESS Then
'    If MaxRegLibComp < UBound(Info) Then
'       MsgBox1 "El archivo tiene más registros de los que soporta el sistema para el libro de compras o ventas (máx. " & MaxRegLibComp - 1 & " Filas)." & vbCrLf & vbCrLf & "No es posible realizar el proceso de importación.", vbExclamation
'       Exit Function
'    End If
'   End If
'
'   If Not LoadCuentasDef Then
'      Exit Function
'   End If
'
'
''   'convertimos el archivo de Unix a DOS porque así lo entrega el SII
''   Rc = ConvUnix2DosFile(fname, fname & "_")
''
''   If Rc < 0 Then    'hubo un error al leer el archivo
''      Exit Function
''   ElseIf Rc = 0 Then   ' no lo convirtió porque ya está en formato DOS, se usa archivo original
''      FNameTmp = ""
''   ElseIf Rc > 0 Then   'lo convirtió y generó arcgivo temporal con guión bajo al final. Leemos del archivo temporal y luego lo borramos
''      fname = fname & "_"
''      FNameTmp = fname
''   End If
''
''   'abrimos el archivo
''   Fd = FreeFile
''   Open fname For Input As #Fd
''   If ERR Then
''      MsgErr fname
''      Import_LibroVentasSIIAuto1 = -ERR
''      If FNameTmp <> "" Then   'generamos uno temporal
''         Kill FNameTmp            'lo borramos
''      End If
''      Exit Function
''   End If
'
'   r = 0
'   Sep = ";"
'   NewDoc = False
'   NDocsOK = 0
'   CorrelativoDoc = 0
'   MsgAcuse = False
'
'   If gDbType = SQL_SERVER Then
'      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
'   End If
'
'   'Campos: Nro;Tipo Doc;Tipo Compra;RUT Cliente;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;Impto. Sin Derecho a Credito;IVA No Retenido;Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
'
'   For i = LBound(Info) To UBound(Info)
'   'Do Until EOF(Fd)
'
'      'Line Input #Fd, Buf
'      l = LBound(Info)
'
'      Buf = Trim(Info(i))
'      'Buf = Trim(Buf)
'
''      '1er registro con nombres de campos
''      If Buf = "" Then
''         GoTo NextRec
''      ElseIf l = 1 Then
''         GoTo NextRec
''      End If
'
'      p = 1
'
'      'ahora leemos los documentos y los insertamos uno por uno
'      n = vFmt(NextField2(Buf, p, Sep))         'número correlativo de registro. Viene en blanco si continúa documento anterior
'      If n = 0 Then
'         If CorrelativoDoc = 0 Then
'            Call AddLogImp(FNameLogImp, fname, l, "Falta número de registro.")
'            DocErr = True
'         Else   'es continuación del documento anterior
'            NewDoc = False
'            r = r + 1         'registro i-esimo del documento NReg
'         End If
'
'      Else                    'documento nuevo
'
'         If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
'            If Not DocErr Then
'               If ValidaTotalesVentas(MsgTot) Then
'                  If GenDocumento Then       'puede que no se agregue si ya existe
'                     NDocsOK = NDocsOK + 1
'                     If MsgTot <> "" Then
'                        Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Advertencia: " & MsgTot)
'                     End If
'                  Else
'                     Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Advertencia: Documento ya existe.")
'                  End If
'               Else
'                  DocErr = True
'                  NDocsConError = NDocsConError + 1
'                  Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, MsgTot)
'
'               End If
'            Else
'               NDocsConError = NDocsConError + 1
'            End If
'
'         End If
'
'         DocErr = False
'         CorrelativoDoc = n
'         r = 0                'primer registro del documento NReg
'
'         'Inicializaciones
'
'         'Nro;Tipo Doc;Tipo Compra;RUT Cliente;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;
'         lCodDocDTESII = ""
'         lTipoDoc = 0
'         lDTE = False
'         lDelGiro = 0
'         lRUTEntidad = ""
'         lRazonSocialEntidad = ""
'         lIdEntidad = 0
'         RazonSocial = ""
'         NotValidRut = False
'         lNumDoc = ""
'         lFechaEmision = 0
'         lFechaRec = 0
'         lFechaAcuse = 0
'         lFechaReclamo = 0
'         FechaAcuse8Dias = 0
'         lIdSucursal = 0
'
'         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
'         lMontoExento = 0
'         lMontoAfecto = 0
'         lMontoIVA = 0
'
'         lMontoIVARetTotal = 0    'IVA Retenido total
'         lMontoIVARetParcial = 0  'IVA Retenido parciaL
'         lMontoIVANoRet = 0       'IVA NO Retenido
'         lMontoIVAPropio = 0      'IVA Propio
'         lMontoIVATerceros = 0    'IVA Terceros
'
'         lCredEmpContructora = 0
'         lMontoTotalPeriodo = 0
'
'         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
'         lMontoTotal = 0
'
'         lTipoDocRef = 0
'         lNumDocRef = ""
'         lIdDocRef = 0
'         lDTEDocRef = 0
'
'
'         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
'         lNCE_NDE_FactCompra = ""
'
'         For k = 0 To UBound(lOtrosImp)
'            lOtrosImp(k).CodSIIDTE = ""
'            lOtrosImp(k).Tasa = 0
'            lOtrosImp(k).Valor = 0
'         Next k
'
'      End If
'
'
'      If r = 0 Then  'nuevo documento
'
'         lCodDocDTESII = Trim(NextField2(Buf, p, Sep))
'         If lCodDocDTESII = "" Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Falta el tipo de documento.")
'            DocErr = True
'         End If
'
'         lDTE = True
'         lTipoDoc = GetTipoDocFromCodDocDTESII(lTipoLib, lCodDocDTESII)
'         If lTipoDoc = 0 Then
'            lDTE = False
'            lTipoDoc = GetTipoDocFromCodDocSII(lTipoLib, lCodDocDTESII)
'            If lTipoDoc = 0 Then
'               Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Tipo de documento inválido.")
'               DocErr = True
'            End If
'         End If
'
'
'         'tipo de compra: Del Giro
'         lDelGiro = IIf(Trim(NextField2(Buf, p, Sep)) = "Del Giro", True, False)
'
'         'RUT Cliente
'         lRUTEntidad = Trim(NextField2(Buf, p, Sep))
'         If InStr(ES_DOC_EXPORT, "," & lCodDocDTESII & ",") <= 0 Then             'FCA - 12/10/2021
'            If Not ValidRut(lRUTEntidad) Then
'               Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "RUT inválido.")
'               DocErr = True
'            End If
'         End If
'
'         'Razón Social
'         lRazonSocialEntidad = Trim(NextField2(Buf, p, Sep))
'         If lRazonSocialEntidad = "" Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Falta razón social.")
'            DocErr = True
'         End If
'
'         lRazonSocialEntidad = Utf82Ansi(Utf82Ansi(lRazonSocialEntidad))  ' se hace dos veces porque al leerlo como texto duplica los caracteres (raro)
'
'         lIdEntidad = GetIdEntidad(lRUTEntidad, RazonSocial, NotValidRut)
'
'         'Folio
'         lNumDoc = Trim(NextField2(Buf, p, Sep))
'         If lNumDoc = "" Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Falta el número de documento.")
'            DocErr = True
'         End If
'
'         'Fecha Docto (Fecha Emision)
'         lFechaEmision = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
'         If lFechaEmision = 0 Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Advertencia: Falta la fecha del Documento.")
'            DocErr = True
'         End If
'
'         'Fecha Recepcion (esta fecha es de recepción en el SII y no nos sirve)
'         lFechaRec = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
'         'no Aplica
'
'         'Fecha Acuse  (esta será la fecha de recepción para nuestro sistema, de acuerdo a lo indicado por Thomson)
'         'no Aplica
'         lFechaAcuse = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
'
'         'Fecha Reclamo
'         'No Aplica
'         lFechaReclamo = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
'
'
'         'Monto Exento;Monto Neto;Monto IVA ;Monto Total
'         lMontoExento = vFmt(NextField2(Buf, p, Sep))
'         lMontoAfecto = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVA = vFmt(NextField2(Buf, p, Sep))
'         lMontoTotal = vFmt(NextField2(Buf, p, Sep))
'
'         'Monto Exento;Monto Neto;Monto IVA Retenido Total ;Monto IVA Retenido Parcial; IVA no Retenido;
'         lMontoIVARetTotal = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVARetParcial = vFmt(NextField2(Buf, p, Sep))
'
'         'IVA Propio; IVA Terceros
'         'No aplica por ahora  (A18, A19)
'         lMontoIVANoRet = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVANoRet = 0    'por cálculos en GenMovDocumento
'         lMontoIVAPropio = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVATerceros = vFmt(NextField2(Buf, p, Sep))
'
'         'Rut emisor Liq. Factura; Neto Comisión Liq. Factura; Exento Comisión Liq. Factura; IVA Comisión Liq. Factura;
'         'No aplican por ahora (A20 -> A23)
'         AuxStr = Trim(NextField2(Buf, p, Sep))
'         AuxMonto = vFmt(NextField2(Buf, p, Sep))
'         AuxMonto = vFmt(NextField2(Buf, p, Sep))
'         AuxMonto = vFmt(NextField2(Buf, p, Sep))
'
'         'IVA fuera de plazo;
'         'No aplica por ahora (A25 -> A28)
'         AuxMonto = vFmt(NextField2(Buf, p, Sep))
'
'         'TipoDoc Referencia; Folio Doc Referencia;
'         AuxStr = Trim(NextField2(Buf, p, Sep))
'         If Val(AuxStr) <> 0 Then
'            lTipoDocRef = GetTipoDocFromCodDocDTESII(lTipoLib, AuxStr)
'         End If
'
'         lNumDocRef = vFmt(NextField2(Buf, p, Sep))
'
'         If lTipoDocRef <> 0 And lNumDocRef <> "" Then
'
'         'pipe Tema 1 2738156
'
'
'            Q1 = "SELECT IdDoc, DTE FROM Documento "
'            Q1 = Q1 & " WHERE NumDoc = '" & lNumDocRef & "'"
'            Q1 = Q1 & " AND TipoLib = " & LIB_VENTAS & " AND TipoDoc = " & lTipoDocRef
'            '2800560
'            Q1 = Q1 & " AND idempresa =" & gEmpresa.id
'            'fin 2800560
'
'            Set Rs = OpenRs(DbMain, Q1)
'
'            If Not Rs.EOF Then
'               lIdDocRef = vFld(Rs("IdDoc"))
'               lDTEDocRef = vFld(Rs("DTE"))
'
'            Else
'                 'pipe Tema 1 2738156
'
'                If vError = 0 Then
'
'                        If MsgBox1("Existen Documentos en el archivo csv que no tienen documento de referencia ¿Desea agregar los documentos de igual forma?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                           vError = 1
'                        Else
'                           vError = 2
'                           Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "El documento de referencia no ha sido ingresado al sistema.")
'                           DocErr = True
'
'                        End If
'                 ElseIf vError = 2 Then
'                   Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "El documento de referencia no ha sido ingresado al sistema.")
'                   DocErr = True
'                  End If
'            End If
'
'            Call CloseRs(Rs)
'         'fin
'         End If
'
'
'         'Num. Ident. Receptor Extranjero; Nacionalidad Receptor Extranjero
'         'No aplica por ahora (A28)
'         AuxMonto = vFmt(NextField2(Buf, p, Sep))
'         AuxStr = Trim(NextField2(Buf, p, Sep))
'
'         'Credito empresa constructora (A29)
'         lCredEmpContructora = vFmt(NextField2(Buf, p, Sep))
'
'         'Impto. Zona Franca (Ley 18211); Garantia Dep. Envases; Indicador Venta sin Costo; Indicador Servicio Periodico; Monto No facturable
'         'No aplican por ahora (A30 -> A34)
'         AuxMonto = vFmt(NextField2(Buf, p, Sep))
'         AuxMonto = vFmt(NextField2(Buf, p, Sep))
'         AuxStr = Trim(NextField2(Buf, p, Sep))
'         AuxStr = Trim(NextField2(Buf, p, Sep))
'         AuxMonto = vFmt(NextField2(Buf, p, Sep))
'
'         'Total Monto Periodo (A35)
'         lMontoTotalPeriodo = vFmt(NextField2(Buf, p, Sep))
'
'         'Venta Pasajes Transporte Nacional; Venta Pasajes Transporte Internacional; Numero Interno
'         'No aplican por ahora (A36 -> A38)
'         AuxStr = Trim(NextField2(Buf, p, Sep))
'         AuxStr = Trim(NextField2(Buf, p, Sep))
'         AuxStr = Trim(NextField2(Buf, p, Sep))
'
'         'Codigo Sucursal
'         CodSucursal = Trim(NextField2(Buf, p, Sep))
'         lIdSucursal = GetIdSucursal(CodSucursal)
'
'
'         'NCE o NDE sobre Fact. de Compra;
'         lNCE_NDE_FactCompra = Trim(NextField2(Buf, p, Sep))
'
'         'Codigo Otro Impuesto; Valor Otro Impuesto; Tasa Otro Impuesto
''         lCodOtroImp = Trim(NextField2(Buf, p, Sep))
''         lMontoOtroImp = vFmt(NextField2(Buf, p, Sep))
''         lTasaOtroImp = vFmt(NextField2(Buf, p, Sep))
'
'
'         'Codigo Otro Impuesto; Valor Otro Impuesto; Tasa Otro Impuesto (r = 0)
'         lOtrosImp(r).CodSIIDTE = Trim(NextField2(Buf, p, Sep))
'         lOtrosImp(r).Valor = vFmt(NextField2(Buf, p, Sep))
'         lOtrosImp(r).Tasa = vFmt(NextField2(Buf, p, Sep))
'
'         If lCodOtroImp <> "" Then
'            IdTipoValLib = GetTipoTipoValLibFromCodSIIDTE(lTipoLib, lCodOtroImp)
'            If IdTipoValLib <= 0 Then
'               Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Código otro impuesto inválido.")
'               DocErr = True
'            End If
'         End If
'
'      Else  'es continuación del registro anterior, nuevo documento
'
'         If lCodDocDTESII <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza tipo de documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         'tipo de venta: Del Giro
'         Txt = NextField2(Buf, p, Sep)
'
'         'RUT Cliente
'         If lRUTEntidad <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza RUT con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Razón Social
'         If lRazonSocialEntidad <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Razón Social con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Folio
'         If lNumDoc <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Nro. Documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Fecha Docto
'         If lFechaEmision <> Int(GetDate(Trim(NextField2(Buf, p, Sep)))) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Fecha Documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         For F = 8 To 40
'            Txt = NextField2(Buf, p, Sep)   'saltamos los campos hasta llegar a los otros impuestos codificados
'         Next F
'
'         lOtrosImp(r).CodSIIDTE = Trim(NextField2(Buf, p, Sep))
'         lOtrosImp(r).Valor = vFmt(NextField2(Buf, p, Sep))
'         lOtrosImp(r).Tasa = vFmt(NextField2(Buf, p, Sep))
'
'         If lOtrosImp(r).CodSIIDTE <> "" And lOtrosImp(r).CodSIIDTE <> "0" Then
'            If Not ValidaOtroImp(lOtrosImp(r).CodSIIDTE) Then
'               Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Código de impuesto inválido.")
'               DocErr = True
'            End If
'         End If
'
'      End If
'
'
''NextRec:
'    Next
''   Loop
'
'   Close #Fd
'
'
'   'agregamos el último documento, si hay
'   If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
'      If Not DocErr Then
'          If ValidaTotalesVentas(MsgTot) Then
'            If GenDocumento Then       'puede que no se agregue si ya existe
'               NDocsOK = NDocsOK + 1
'            Else
'               Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Documento ya existe.")
'            End If
'         Else
'            DocErr = True
'            NDocsConError = NDocsConError + 1
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, MsgTot)
'         End If
'      Else
'         NDocsConError = NDocsConError + 1
'      End If
'
'   End If
'
'EndFnc:
'
'
'
'   If NDocsOK > 1 Then
'      MsgDocsOK = "Se importaron " & NDocsOK & " documentos nuevos."
'
'   ElseIf NDocsOK = 1 Then
'      MsgDocsOK = "Se importó un documento nuevo."
'
'   Else
'      MsgDocsOK = "No se importaron documentos nuevos."
'   End If
'
'
'   If NDocsConError = 0 Then
'      MsgBox1 "Proceso de importación finalizado." & vbNewLine & vbNewLine & MsgDocsOK, vbInformation + vbOKOnly
'
'      '2907989
'      If NDocsOK = 0 Then
'        If MsgBox1("¿Desea revisar el log de importación " & FNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'           Call ShellExecute(Frm.hWnd, "open", FNameLogImp, "", "", SW_SHOW)
'        End If
'      End If
'
'      '2907989
'
'   Else
'      If NDocsConError = 1 Then
'         MsgBox1 "Se encontró un documento con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
'      Else
'         MsgBox1 "Se encontraron " & NDocsConError & " documentos con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
'      End If
'
'      If MsgBox1("¿Desea revisar el log de importación " & FNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'         Call ShellExecute(Frm.hWnd, "open", FNameLogImp, "", "", SW_SHOW)
'      End If
'
'   End If
'
'   If gDbType = SQL_SERVER Then
'      Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")   'para que no trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
'   End If
'
'   If FNameTmp <> "" Then   'generamos uno temporal
'      Kill FNameTmp            'lo borramos
'   End If
'
'

End Function
'Fin 2862611
'Sin DLL
Public Function Import_LibroRetencionesSIIAuto(lAno As Integer, lMes As Integer, lCtaHonSinRet As Cuenta_t, lCtaBruto As Cuenta_t, Frm As Form, Info() As String) As Boolean
    Dim lFNameLogImp As String
    Dim Dt1 As Long, Dt2 As Long
    Dim ClasifEnt As Integer
    Dim i As Integer, l As Integer
    Dim Row As Integer, r As Integer
    Dim Sep As String
    Dim Buf As String
    Dim Q1 As String
    Dim Rs As Recordset
    Dim p As Long
    Dim CampoInvalido As String
    Dim NumDoc As String
    Dim FName As String
    Dim DtRec As Long, DtEmi As Long
    Dim TipoDoc As String
    Dim IdTipoDoc As Integer
    Dim DTE As Integer
    Dim Aux As String
    Dim IdEnt As Long
    Dim NotValidRut As Boolean
    Dim RutEnt As String, CodEnt As String, NombEnt As String
    Dim Descrip As String
    Dim Bruto As Double, Impuesto As Double, Neto As Double
    Dim AuxCodCta As String, AuxIdCta As Long, AuxDescCta As String
    Dim FldArray(6) As AdvTbAddNew_t
    Dim FldArray1(3) As AdvTbAddNew_t
    Dim IdDoc As Long
    Dim NRecErroneos As Long, StrNRecErroneos As String
    Dim Fd As Long
    Dim impBruto As Long
    Dim es3Por As Boolean

   vError = 0


   Import_LibroRetencionesSIIAuto = False

   On Error Resume Next

   lFNameLogImp = gImportPath & "\Log\ImpLibCompSII-" & Format(Now, "yyyymmdd") & ".log"

   On Error Resume Next


   Call FirstLastMonthDay(DateSerial(lAno, lMes, 1), Dt1, Dt2)


   On Error Resume Next
'   FrmMain.Cm_ComDlg.ShowOpen
'
'   If ERR = cdlCancel Then
'      Exit Function
'   ElseIf ERR Then
'      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
'      Exit Function
'   End If

'   If FrmMain.Cm_ComDlg.Filename = "" Then
'      Exit Function
'   End If
   ERR.Clear

   On Error GoTo 0

'   fname = FrmMain.Cm_ComDlg.Filename

'   MousePointer = vbHourglass
   DoEvents

'   Rc = MsgBox1("Atención:" & vbNewLine & vbNewLine & "Se importará el archivo:" & vbNewLine & vbNewLine & fname & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
'   If Rc = vbNo Then
'      Exit Function
'   End If

   'abrimos el archivo
'   Fd = FreeFile
'   Open fname For Input As #Fd
'   If ERR Then
'      MsgErr fname
'      Import_LibroRetencionesSIIAuto = -ERR
'      Exit Function
'   End If

   ClasifEnt = ENT_PROVEEDOR

'   For i = Grid.FixedRows To Grid.rows - 1
'      If Grid.TextMatrix(i, C_FECHA) = "" Then    'ya terminó la lista
'         Exit For
'      End If
'   Next i

   Row = i
   r = 0

   'Grid.FlxGrid.Redraw = False
    r = 0
   Sep = ";"
   'NewDoc = False
   'NDocsOK = 0
   'CorrelativoDoc = 0
   'MsgAcuse = False

   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If

   'Campos: Nro;Tipo Doc;Tipo Compra;RUT Proveedor;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;Impto. Sin Derecho a Credito;IVA No Retenido;Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto


   For i = LBound(Info) To UBound(Info)
   'Do Until EOF(Fd)

      'Line Input #Fd, Buf
      l = LBound(Info)

      Buf = Trim(Info(i))

'      '1er registro con nombres de campos
'      If Buf = "" Then
'         GoTo NextRec
'      ElseIf l = 1 Then
'         GoTo NextRec
'      End If

      p = 1
      Buf = Trim(Buf)



      '1er registro con nombres de campos
'      If Buf = "" Then
'         GoTo NextRec
'      ElseIf l = 1 And InStr(1, Buf, "RUT", vbTextCompare) Then
'         GoTo NextRec
'      End If

      CampoInvalido = ""

      'NumDoc
      NumDoc = Trim(Trim(NextField2(Buf, p, Sep)))
      If NumDoc = "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, FName, l, "N° de documento inválido.")
      End If

      'Fecha recepción/emisión
'      Aux = Trim(NextField2(Buf, p, Sep))
'      DtRec = ValFmtDate(Aux, False)
       DtRec = ValFmtDate(Trim(NextField2(Buf, p, Sep)), False)
      If DtRec = 0 Or DtRec < Dt1 Or DtRec > Dt2 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, FName, l, "Fecha recepción inválida o fuera del mes en edición.")
      End If

      'Tipo Doc
      TipoDoc = "BOH"
      IdTipoDoc = FindTipoDoc(LIB_RETEN, TipoDoc)
      If IdTipoDoc = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, FName, l, "Tipo de documento inválido o no corresponde al libro en edición. Valores perimtidos ""BOH"", ""BRT"".")
      End If

      'DTE

      'TxtDTE = "X"
      DTE = 1 'IIf(Val(TxtDTE) = 0 Or Trim(TxtDTE) = "", 0, 1)


      'Fecha emisión
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      'Aux = Trim(NextField2(Buf, p, Sep))
      'Aux = Trim(NextField2(Buf, p, Sep))
      DtEmi = DtRec 'ValFmtDate(Aux, False)
      If DtEmi = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, FName, l, "Fecha emisión inválida.")
      End If

      'Entidad
      IdEnt = 0
      NotValidRut = False
      Aux = Trim(NextField2(Buf, p, Sep))
      'Aux = Trim(NextField2(Buf, p, Sep))
      If Aux = "0-0" Or Aux = "" Then
         RutEnt = ""
      ElseIf Aux = "NULO" Then
         RutEnt = "NULO"
      Else
         RutEnt = vFmtCID(Aux)
         If RutEnt = "0" Or RutEnt = "" Then    'es inválido
            NotValidRut = True
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "RUT inválido")
         End If
      End If

      CodEnt = RutEnt

      NombEnt = Trim(NextField2(Buf, p, Sep))
      If NombEnt = "" And RutEnt <> "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, FName, l, "Falta ingresar nombre o razón social entidad.")
      End If

      'Descrpción
      'Descrip = Trim(NextField2(Buf, p, Sep))
      Descrip = "BH " & NumDoc

'      'sucursal
'      CodSuc = ""
'      IdSucursal = 0
'      Sucursal = ""
'
'      If CodSuc <> "" Then
'         Q1 = "SELECT IdSucursal, Descripcion FROM Sucursales WHERE Codigo ='" & CodSuc & "'"
'         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
'         Set Rs = OpenRs(DbMain, Q1)
'
'         If Rs.EOF Then
'            CampoInvalido = CampoInvalido & "," & p
'            Call AddLogImp(lFNameLogImp, fname, l, "Código de sucursal inválido")
'         Else
'            IdSucursal = vFld(Rs("IdSucursal"))
'            Sucursal = vFld(Rs("Descripcion"))
'         End If
'
'         Call CloseRs(Rs)
'      End If

      'Valores
      'HonSinRet = vFmt(Trim(NextField2(Buf, p, Sep)))
      Aux = Trim(NextField2(Buf, p, Sep)) 'vFmt(Trim(NextField2(Buf, p, Sep)))
      Neto = vFmt(Trim(NextField2(Buf, p, Sep)))
      Impuesto = vFmt(Trim(NextField2(Buf, p, Sep)))
      es3Por = False
      'Impuesto = Neto * 0.16
      impBruto = Neto * gImpRet(Val(1))
      If Impuesto > 0 And Abs(Impuesto - impBruto) > 2 Then
        es3Por = True
      End If
      'Impuesto = vFmt(Trim(NextField2(Buf, p, Sep)))
      Bruto = vFmt(Trim(NextField2(Buf, p, Sep)))
      'StrPImp = Trim(NextField2(Buf, p, Sep))
      'PImp = vFmt(StrPImp)
      
      'Ret3Porc = vFmt(Trim(NextField2(Buf, p, Sep)))
      

      'Cuentas Contables

      If Bruto < 0 Or Neto < 0 Or Impuesto < 0 Then
         CampoInvalido = CampoInvalido & "," & "Bruto, Impuesto, Neto"
         Call AddLogImp(lFNameLogImp, FName, l, "Valor de Bruto, Impuesto y/o Neto inválido.")
      End If

'      If HonSinRet > 0 And Bruto > 0 Then
'         CampoInvalido = CampoInvalido & "," & "Honorarios, Bruto"
'         Call AddLogImp(lFNameLogImp, fname, l, "No es posible ingresar Honorarios y valor Bruto para un mismo documento.")
'      End If

      If Bruto > 0 And Neto = 0 Then
         CampoInvalido = CampoInvalido & "," & "Neto"
         Call AddLogImp(lFNameLogImp, FName, l, "El valor Neto está en cero o es inválido.")
      End If

'      Select Case LCase(StrPImp)
'         Case gImpRet(IMPRET_NAC) * 100
'            IdPImp = IMPRET_NAC
'            StrPImp = StrPImp & "%"
'
'         Case gImpRet(IMPRET_EXT) * 100
'            IdPImp = IMPRET_EXT
'            StrPImp = StrPImp & "%"
'
'         Case "otro"
'            IdPImp = IMPRET_OTRO
'            StrPImp = "Otro"
'
'         Case Else
'            CampoInvalido = CampoInvalido & "," & "% Imp"
'            Call AddLogImp(lFNameLogImp, fname, l, "Porcentaje de impuesto inválido. Valores perimitidos """ & gImpRet(IMPRET_NAC) * 100 & ", """ & gImpRet(IMPRET_EXT) * 100 & ", ""Otro"".")
'
'      End Select

      'TipoReten = Trim(NextField2(Buf, p, Sep))
'      TipoReten = "honorarios"
'
'      Select Case LCase(TipoReten)
'         Case "honorarios"
'            IdTipoReten = TR_HONORARIOS
'            TipoReten = "Honorarios"
'
'         Case "dieta"
'            IdTipoReten = TR_DIETA
'            TipoReten = "Dieta"
'
'         Case "otro"
'            IdTipoReten = TR_OTRO
'            TipoReten = "Otro"
'
'         Case Else
'            CampoInvalido = CampoInvalido & "," & "% Imp"
'            Call AddLogImp(lFNameLogImp, fname, l, "Tipo retención inválido. Valores permitidos ""Honorarios"", ""Dieta"", ""Otro"".")
'
'      End Select

'      If Ret3Porc > 0 Then
'         If gCtasBas.IdCtaRet3Porc = 0 Then
'            CampoInvalido = CampoInvalido & "," & p
'            Call AddLogImp(lFNameLogImp, fname, l, "Falta definir la cuenta básica para Retención 3%'.")
'         End If
'         If IdTipoReten <> TR_HONORARIOS Then
'            CampoInvalido = CampoInvalido & "," & p
'            Call AddLogImp(lFNameLogImp, fname, l, "Tipo retención no permite 'Retención 3% Préstamo Solidario'.")
'         End If
'      End If

      'código cuenta
'      AuxCodCta = VFmtCodigoCta(Trim(NextField2(Buf, p, Sep)))
'
'      If AuxCodCta <> "" Then
'         AuxIdCta = GetIdCuenta(NomCta, AuxCodCta, AuxDescCta, UltNivel)
'         If AuxIdCta <= 0 Or Not UltNivel Then
'            CampoInvalido = CampoInvalido & "," & p
'            Call AddLogImp(lFNameLogImp, fname, l, "Código de cuenta inválido")
'         End If
'      Else
         AuxIdCta = 0
         AuxDescCta = ""
'      End If

'      NomCta = ""

      'Cuenta Contable Default

'      If HonSinRet <> 0 Then
'         If AuxIdCta > 0 Then
'            IdCtaHonSinRet = AuxIdCta
'            CodCtaHonSinRet = FmtCodCuenta(AuxCodCta)
'            DescCtaHonSinRet = AuxDescCta
'         Else
'            IdCtaHonSinRet = lCtaHonSinRet.id
'            CodCtaHonSinRet = FmtCodCuenta(lCtaHonSinRet.Codigo)
'            DescCtaHonSinRet = lCtaHonSinRet.Descripcion
'         End If
'      Else
'         IdCtaHonSinRet = 0
'         CodCtaHonSinRet = ""
'         DescCtaHonSinRet = ""
'      End If

'      If Bruto <> 0 Then
'         If AuxIdCta > 0 Then
'            IdCtaBruto = AuxIdCta
'            CodCtaBruto = FmtCodCuenta(AuxCodCta)
'            DescCtaBruto = AuxDescCta
'         Else
'            IdCtaBruto = lCtaBruto.id
'            CodCtaBruto = FmtCodCuenta(lCtaBruto.Codigo)
'            DescCtaBruto = lCtaBruto.Descripcion
'         End If
'      Else
'         IdCtaBruto = 0
'         CodCtaBruto = ""
'         DescCtaBruto = ""
'      End If

'      If HonSinRet <> 0 Then
'         IdCtaHonSinRet = lCtaHonSinRet.id
'         CodCtaHonSinRet = FmtCodCuenta(lCtaHonSinRet.Codigo)
'         DescCtaHonSinRet = lCtaHonSinRet.Descripcion
'      Else
'         IdCtaHonSinRet = 0
'         CodCtaHonSinRet = ""
'         DescCtaHonSinRet = ""
'      End If
'
'
'      If Bruto <> 0 Then
'         IdCtaBruto = lCtaBruto.id
'         CodCtaBruto = FmtCodCuenta(lCtaBruto.Codigo)
'         DescCtaBruto = lCtaBruto.Descripcion
'      Else
'         IdCtaBruto = 0
'         CodCtaBruto = ""
'         DescCtaBruto = ""
'      End If


'      'Fecha Vencim
'      Aux = Trim(NextField2(Buf, p, Sep))
'      DtVenc = 0
'      If Aux <> "" Then
'         DtVenc = ValFmtDate(Aux, False)
'         If DtVenc = 0 Then
'            CampoInvalido = CampoInvalido & "," & p
'            Call AddLogImp(lFNameLogImp, fname, l, "Fecha vencimiento inválida.")
'         End If
'      Else
'         CampoInvalido = CampoInvalido & "," & p
'         Call AddLogImp(lFNameLogImp, fname, l, "Falta ingresar Fecha Vencimiento.")
'      End If

      'si no hay errores y la entidad no existe, la insertamos

      If CampoInvalido = "" Then

         If RutEnt <> "" And RutEnt <> "NULO" Then

            Q1 = "SELECT IdEntidad, Nombre FROM Entidades WHERE Rut = '" & RutEnt & "'"
            Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
            Set Rs = OpenRs(DbMain, Q1)
            If Not Rs.EOF Then
               IdEnt = vFld(Rs("IdEntidad"))
               NombEnt = vFld(Rs("Nombre"))
            End If
            Call CloseRs(Rs)

            If IdEnt = 0 Then  'no existe

               'insertamos la nueva entidad

'               Set Rs = DbMain.OpenRecordset("Entidades", dbOpenTable)
'               Rs.AddNew
'
'               IdEnt = Rs("IdEntidad")
'               Rs("RUT") = RutEnt
'               Rs("Codigo") = CodEnt
'               Rs("Nombre") = NombEnt
'               Rs("Clasif" & ClasifEnt) = 1
'
'               Rs.Update
'               Rs.Close

               FldArray(0).FldName = "NotValidRut"
               FldArray(0).FldValue = 0
               FldArray(0).FldIsNum = True

               FldArray(1).FldName = "RUT"
               FldArray(1).FldValue = RutEnt
               FldArray(1).FldIsNum = False

               FldArray(2).FldName = "IdEmpresa"
               FldArray(2).FldValue = gEmpresa.id
               FldArray(2).FldIsNum = True

               FldArray(3).FldName = "Codigo"
               FldArray(3).FldValue = CodEnt
               FldArray(3).FldIsNum = False

               FldArray(4).FldName = "Nombre"
               FldArray(4).FldValue = NombEnt
               FldArray(4).FldIsNum = False

               FldArray(5).FldName = "Clasif" & ClasifEnt
               FldArray(5).FldValue = 1
               FldArray(5).FldIsNum = True
               
               If es3Por Then
                    FldArray(6).FldName = "Ret3Porc"
                    FldArray(6).FldValue = -1
                    FldArray(6).FldIsNum = True
               End If

               IdEnt = AdvTbAddNewMult(DbMain, "Entidades", "IdEntidad", FldArray)
            Else
                If es3Por Then
                    Q1 = "UPDATE Entidades SET "
                    Q1 = Q1 & "  Ret3Porc = -1 "
                    Q1 = Q1 & " WHERE IdEntidad = " & IdEnt
                    Call ExecSQL(DbMain, Q1)
                End If
            End If

         End If

    Q1 = "SELECT IdDoc FROM Documento "
    Q1 = Q1 & " WHERE TipoLib = " & LIB_RETEN & " AND TipoDoc = " & IdTipoDoc
    Q1 = Q1 & " AND NumDoc = '" & NumDoc & "'"
    Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
    '625206
    Q1 = Q1 & " AND idEntidad = " & IdEnt
    '625206
   Set Rs = OpenRs(DbMain, Q1)

   If Rs.EOF = True Then      'documento no existe

        FldArray1(0).FldName = "IdUsuario"
        FldArray1(0).FldValue = gUsuario.IdUsuario
        FldArray1(0).FldIsNum = True

        FldArray1(1).FldName = "FechaCreacion"
        FldArray1(1).FldValue = CLng(Int(Now))
        FldArray1(1).FldIsNum = True

        FldArray1(2).FldName = "IdEmpresa"
        FldArray1(2).FldValue = gEmpresa.id
        FldArray1(2).FldIsNum = True

        FldArray1(3).FldName = "Ano"
        FldArray1(3).FldValue = gEmpresa.Ano
        FldArray1(3).FldIsNum = True

        IdDoc = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray1)
        r = r + 1
   '625206
   Else
        IdDoc = vFld(Rs("IdDoc"))
   '625206
   End If

   Call CloseRs(Rs)

   
   Q1 = "UPDATE Documento SET "
         Q1 = Q1 & "  TipoLib = " & LIB_RETEN
         Q1 = Q1 & ", TipoDoc = " & IdTipoDoc
         Q1 = Q1 & ", NumDoc = '" & NumDoc & "'"
         Q1 = Q1 & ", DTE = -1"
         Q1 = Q1 & ", IdEntidad = " & IdEnt
         Q1 = Q1 & ", EntRelacionada = " & Abs(GetEntRelacionada(IdEnt))
         Q1 = Q1 & ", IdSucursal = 0"

         Q1 = Q1 & ", TipoEntidad = " & ENT_PROVEEDOR

         'por si acaso, ponemos la clasificación de la entidad
         Call ExecSQL(DbMain, "UPDATE Entidades SET Clasif" & ENT_PROVEEDOR & " = 1 WHERE IdEntidad = " & Val(IdEnt))

         
         Q1 = Q1 & ", FEmision = " & CLng(DtRec)
         If Impuesto > 0 Then
            
            If es3Por Then
                Impuesto = impBruto
                Q1 = Q1 & ", ValRet3Porc = " & Fix(Neto * 0.03)
                Q1 = Q1 & ", IdCuentaRet3Porc = " & gCtasBas.IdCtaRet3Porc
            End If
            
            Q1 = Q1 & ", Afecto = " & Neto 'Bruto
            Q1 = Q1 & ", IdCuentaAfecto = " & LoadDefCuentasRet(LIBRETEN_BRUTO, LIB_RETEN)
            
         Else
            Q1 = Q1 & ", Exento = " & Neto 'Bruto
            Q1 = Q1 & ", IdCuentaExento = " & LoadDefCuentasRet(LIBRETEN_HONORSINRET, LIB_RETEN)
         End If
         Q1 = Q1 & ", PorcentRetencion = 1 "
         Q1 = Q1 & ", TipoRetencion = 1 "
         Q1 = Q1 & ", OtroImp = " & Fix(Impuesto)
         Q1 = Q1 & ", IdCuentaOtroImp = " & gCtasBas.IdCtaImpRet
         Q1 = Q1 & ", Total = " & Bruto + Impuesto 'Neto
         Q1 = Q1 & ", IdCuentaTotal = " & gCtasBas.IdCtaNetoHon
         Q1 = Q1 & ", FEmisionOri = " & CLng(DtRec)
         Q1 = Q1 & ", FVenc = 0"
         Q1 = Q1 & ", Estado = " & ED_PENDIENTE
         Q1 = Q1 & ", SaldoDoc = NULL "
         Q1 = Q1 & " WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)


      Else
         NRecErroneos = NRecErroneos + 1
      End If

'NextRec:
'   Loop
Next
   Close #Fd


   If NRecErroneos = 0 Then
      If r = 1 Then
         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregó " & r & " documento.", vbInformation + vbOKOnly
         'MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
         'MsgBox1 "Favor para finalizar la importacion dirijase al libro de retenciones para agregar las cuentas asociadas", vbInformation
      ElseIf r > 1 Then
         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos.", vbInformation + vbOKOnly
         'MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
         'MsgBox1 "Favor para finalizar la importacion dirijase al libro de retenciones para agregar las cuentas asociadas", vbInformation
      Else  ' r=0
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & "- No se agregaron documentos.", vbInformation + vbOKOnly
      End If

   Else
      If NRecErroneos > 1 Then
         StrNRecErroneos = "- Se encontraron " & NRecErroneos & " registros con errores en el archivo."
      Else
         StrNRecErroneos = "- Se encontró " & NRecErroneos & " registro con errores en el archivo."
      End If

      If r = 1 Then
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregó " & r & " documento.", vbInformation + vbOKOnly
         'MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
         'MsgBox1 "Favor para finalizar la importacion dirijase al libro de retenciones para agregar las cuentas asociadas", vbInformation
      ElseIf r > 1 Then
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos.", vbInformation + vbOKOnly
         'MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
         'MsgBox1 "Favor para finalizar la importacion dirijase al libro de retenciones para agregar las cuentas asociadas", vbInformation
      Else  ' r=0
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- No se agregaron documentos.", vbInformation + vbOKOnly
      End If

      If MsgBox1("¿Desea revisar el log de importación " & lFNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Frm.hWnd, "open", lFNameLogImp, "", "", SW_SHOW)
      End If
   End If

   Import_LibroRetencionesSIIAuto = True

End Function
'CON DLL Descomentar y agregar DLL a los componentes
Private Function Import_LibroRetencionesSIIAuto1() As Boolean
'   Dim fname As String
'   Dim Buf As String
'   Dim Q1 As String
'   Dim Rs As Recordset
'   Dim ImpEnable As Boolean
'   Dim IdEnt As Long
'   Dim NotValidRut As Boolean
'   Dim i As Integer, l As Integer
'   Dim j As Integer, p As Long
'   Dim NRecErroneos As Long, StrNRecErroneos As String
'   Dim Rc As Integer
'   Dim Fd As Long
'   Dim Aux As String
'   Dim DtRec As Long, DtEmi As Long, DtVenc As Long
'   Dim CampoInvalido As String
'   Dim IdTipoDoc As Integer
'   Dim DTE As Integer, DelGiro As Integer, TxtDTE As String
'   Dim NumDoc As String, NumDocHasta As String
'   Dim RutEnt As String, CodEnt As String, NombEnt As String
'   Dim ClasifEnt As Integer
'   Dim HonSinRet As Double, Bruto As Double, PImp As Double, Impuesto As Double, Neto As Double
'   Dim StrPImp As String, IdPImp As Integer
'   Dim Row As Integer, r As Integer
'   Dim TipoDoc As String
'   Dim CodSuc As String, IdSucursal As Long, Sucursal As String
'   Dim Descrip As String
'   Dim Dt1 As Long, Dt2 As Long
'   Dim IdCtaHonSinRet As Long, CodCtaHonSinRet As String, DescCtaHonSinRet As String
'   Dim IdCtaBruto As Long, CodCtaBruto As String, DescCtaBruto As String
'   Dim IdCtaImp As Long, IdCtaNeto As Long
'   Dim Estado As Integer
'   Dim IdTipoReten As Integer, TipoReten As String
'   Dim AuxCodCta As String, AuxIdCta As Long, AuxDescCta As String
'   Dim NomCta As String, UltNivel As Boolean
'   Dim FldArray(5) As AdvTbAddNew_t
'   Dim Ret3Porc As Double, ValidaRet3Porc As Boolean
'   Dim MaxRegLibComp As Integer
'   Dim FNameLogImp As String
'   Dim CSII As TR_CCONECTSII.TR_CCONECTSII
'   Dim Info() As String
'
'   Set CSII = New TR_CCONECTSII.TR_CCONECTSII
'   'Info = CSII.GetBoletasEmitidas("17.533.256", "1", "Fpriet4512", "2020", "07")
'   Info = CSII.GetBoletasRecibidas(vFmtCID(Rut), DV_Rut(vFmtCID(Rut)), Clave, Ano, Mes)
'   'pipe 2738156 tema 1
'   vError = 0
'
'
'   Import_LibroRetencionesSIIAuto1 = False
'
'   On Error Resume Next
'
'   FNameLogImp = gImportPath & "\Log\ImpLibCompSII-" & Format(Now, "yyyymmdd") & ".log"
'
'   MaxRegLibComp = 3501
'   On Error Resume Next
'
'   Estado = ED_PENDIENTE
'
'   Call FirstLastMonthDay(DateSerial(lAno, lMes, 1), Dt1, Dt2)
'
''   FrmMain.Cm_ComDlg.CancelError = True
''   FrmMain.Cm_ComDlg.Filename = ""
''   FrmMain.Cm_ComDlg.InitDir = gImportPath
''   FrmMain.Cm_ComDlg.Filter = "Archivos de Texto (*.txt)|*.txt"
''   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
''   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
'
'   On Error Resume Next
''   FrmMain.Cm_ComDlg.ShowOpen
''
''   If ERR = cdlCancel Then
''      Exit Function
''   ElseIf ERR Then
''      MsgBox1 "Error " & ERR & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
''      Exit Function
''   End If
'
''   If FrmMain.Cm_ComDlg.Filename = "" Then
''      Exit Function
''   End If
'   ERR.Clear
'
'   On Error GoTo 0
'
''   fname = FrmMain.Cm_ComDlg.Filename
'
'   MousePointer = vbHourglass
'   DoEvents
'
''   Rc = MsgBox1("Atención:" & vbNewLine & vbNewLine & "Se importará el archivo:" & vbNewLine & vbNewLine & fname & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
''   If Rc = vbNo Then
''      Exit Function
''   End If
'
'   'abrimos el archivo
''   Fd = FreeFile
''   Open fname For Input As #Fd
''   If ERR Then
''      MsgErr fname
''      Import_LibroRetencionesSIIAuto1 = -ERR
''      Exit Function
''   End If
'
'   ClasifEnt = ENT_PROVEEDOR
'
''   For i = Grid.FixedRows To Grid.rows - 1
''      If Grid.TextMatrix(i, C_FECHA) = "" Then    'ya terminó la lista
''         Exit For
''      End If
''   Next i
'
'   Row = i
'   r = 0
'
'   'Grid.FlxGrid.Redraw = False
'    r = 0
'   Sep = ";"
'   NewDoc = False
'   NDocsOK = 0
'   CorrelativoDoc = 0
'   MsgAcuse = False
'
'   If gDbType = SQL_SERVER Then
'      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
'   End If
'
'   'Campos: Nro;Tipo Doc;Tipo Compra;RUT Proveedor;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;Impto. Sin Derecho a Credito;IVA No Retenido;Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
'
'
'   For i = LBound(Info) To UBound(Info)
'   'Do Until EOF(Fd)
'
'      'Line Input #Fd, Buf
'      l = LBound(Info)
'
'      Buf = Trim(Info(i))
'
''      '1er registro con nombres de campos
''      If Buf = "" Then
''         GoTo NextRec
''      ElseIf l = 1 Then
''         GoTo NextRec
''      End If
'
'      p = 1
'
'      'ahora leemos los documentos y los insertamos uno por uno
'      n = vFmt(NextField2(Buf, p, Sep))         'número correlativo de registro. Viene en blanco si continúa documento anterior
'      If n = 33 Then
'         n = 33
'      End If
'      If n = 0 Then
'         If CorrelativoDoc = 0 Then
'            Call AddLogImp(FNameLogImp, fname, l, "Falta número de registro.")
'            DocErr = True
'         Else   'es continuación del documento anterior
'            NewDoc = False
'            r = r + 1         'registro i-esimo del documento NReg
'         End If
'
'      Else                    'documento nuevo
'
'         If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
'            If Not DocErr Then
'               If ValidaTotalesCompras(MsgTot) Then
'                  If GenDocumento Then       'puede que no se agregue si ya existe
'                     NDocsOK = NDocsOK + 1
'                  Else
'                     Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Advertencia: Documento ya existe.")
'                  End If
'               Else
'                  DocErr = True
'                  NDocsConError = NDocsConError + 1
'                  Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, MsgTot)
'
'               End If
'            Else
'               NDocsConError = NDocsConError + 1
'            End If
'
'         End If
'
'         DocErr = False
'         CorrelativoDoc = n
'         r = 0                'primer registro del documento NReg
'
'         'Inicializaciones
'
'         'Nro;Tipo Doc;Tipo Compra;RUT Proveedor;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;
'         lCodDocDTESII = ""
'         lTipoDoc = 0
'         lDTE = False
'         lDelGiro = 0
'         lRUTEntidad = ""
'         lRazonSocialEntidad = ""
'         lIdEntidad = 0
'         RazonSocial = ""
'         NotValidRut = False
'         lNumDoc = ""
'         lFechaEmision = 0
'         lFechaRec = 0
'         lFechaAcuse = 0
'         FechaAcuse8Dias = 0
'
'         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
'         lMontoExento = 0
'         lMontoAfecto = 0
'         lMontoIVARec = 0
'         lMontoIVANoRec = 0
'         lCodIVANoRec = 0
'
'         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
'         lMontoTotal = 0
'         lMontoNetoActFijo = 0
'         lMontoIVAActFijo = 0
'         lMontoIVAUsoComun = 0
'
'         'Impto. Sin Derecho a Credito;IVA No Retenido;
'         lMontoImpSinDerechoCred = 0
'         lMontoIVANoRet = 0
'
'         lMontoOtroImp = 0   'totalizado para encabezado del doc
'
'         'Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;
'         lTabacosPuros = 0
'         lTabacosCigarrillos = 0
'         lTabacosElaborados = 0
'
'         lIdDocRef = 0
'         lTipoDocRef = 0
'         lNumDocRef = ""
'         lDTEDocRef = 0
'
'
'         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
'         lNCE_NDE_FactCompra = ""
'
'         For k = 0 To UBound(lOtrosImp)
'            lOtrosImp(k).CodSIIDTE = ""
'            lOtrosImp(k).Tasa = 0
'            lOtrosImp(k).Valor = 0
'         Next k
'
''         lIVANoRec = False
''         For k = 0 To UBound(lCodIVANoRec)
''            lCodIVANoRec(k) = 0
''            lMontoIVANoRec(k) = 0
''         Next k
'
'      End If
'
'
'      If r = 0 Then  'nuevo documento
'
'         lCodDocDTESII = Trim(NextField2(Buf, p, Sep))
'         If lCodDocDTESII = "" Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Falta el tipo de documento.")
'            DocErr = True
'         End If
'
'         lDTE = True
'         lTipoDoc = GetTipoDocFromCodDocDTESII(lTipoLib, lCodDocDTESII)
'         If lTipoDoc = 0 Then
'            lDTE = False
'            If lCodDocDTESII = "30" Then
'               lTipoDoc = FindTipoDoc(lTipoLib, "FAC")
'            ElseIf lCodDocDTESII = "32" Then
'               lTipoDoc = FindTipoDoc(lTipoLib, "FCE")
'            Else
'               Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Tipo de documento inválido.")
'               DocErr = True
'            End If
'         End If
'
'
'         'tipo de compra: Del Giro
'         lDelGiro = IIf(Trim(NextField2(Buf, p, Sep)) = "Del Giro", True, False)
'
'         'RUT Proveedor
'         lRUTEntidad = Trim(NextField2(Buf, p, Sep))
'         If Not ValidRut(lRUTEntidad) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "RUT inválido.")
'            DocErr = True
'         End If
'
'         'Razón Social
'         lRazonSocialEntidad = Trim(NextField2(Buf, p, Sep))
'         If lRazonSocialEntidad = "" Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Falta razón social.")
'            DocErr = True
'         End If
'
'         lRazonSocialEntidad = Utf82Ansi(Utf82Ansi(lRazonSocialEntidad))  ' se hace dos veces porque al leerlo como texto duplica los caracteres (raro)
'
'         lIdEntidad = GetIdEntidad(lRUTEntidad, RazonSocial, NotValidRut)
'
'         'Folio
'         lNumDoc = Trim(NextField2(Buf, p, Sep))
'         If lNumDoc = "" Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Falta el número de documento.")
'            DocErr = True
'         End If
'
'         'Fecha Docto (Fecha Emision)
'         lFechaEmision = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
'         If lFechaEmision = 0 Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Advertencia: Falta la fecha del Documento.")
'            DocErr = True
'         End If
'
'
'         'Fecha Recepcion (esta fecha es de recepción en el SII y no nos sirve)
'         lFechaRec = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
''         If Month(lFechaRec) <> Mes Or Year(lFechaRec) <> Ano Then
''            Call AddLogImp(FNameLogImp, fname, l, "Fecha de recepción no corresponde al mes-año seleccionado.")
''            DocErr = True
''         End If
'
'         'Fecha Acuse  (esta será la fecha de recepción para nuestro sistema, de acuerdo a lo indicado por Thomson)
'         lFechaAcuse = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
'         If lFechaAcuse = 0 Then
'
'            If Not MsgAcuse Then
'               If MsgBox1("Se han detectado que uno o más documentos no poseen fecha de acuse de recibo, por lo que podrían existir documentos que deberían ser contabilizados en otro mes." & vbCrLf & vbCrLf & "¿Esta seguro de continuar con el proceso?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                  GoTo EndFnc
'               End If
'               MsgAcuse = True
'            End If
'
'            '8 días para el acuse
'            FechaAcuse8Dias = DateAdd("d", 8, lFechaRec)
'            If month(FechaAcuse8Dias) <> Mes Or Year(FechaAcuse8Dias) <> Ano Then
'               Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "El documento no puede ser contabilizado en el mes-año seleccionado, ya que la fecha de acuse de recibo debería corresponder a otro mes.")
'               DocErr = True
'            End If
'
'            lFechaAcuse = FechaAcuse8Dias   'dejamos esta como fecha de acuse
'
'
'         ElseIf month(lFechaAcuse) <> Mes Or Year(lFechaAcuse) <> Ano Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "La fecha del acuse de recibo no corresponde al mes-año seleccionado.")
'            DocErr = True
'         End If
'
'
'         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
'         lMontoExento = vFmt(NextField2(Buf, p, Sep))
'         lMontoAfecto = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVARec = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVANoRec = vFmt(NextField2(Buf, p, Sep))
'         lCodIVANoRec = vFmt(NextField2(Buf, p, Sep))
'
'         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
'         lMontoTotal = vFmt(NextField2(Buf, p, Sep))
'         lMontoNetoActFijo = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVAActFijo = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVAUsoComun = vFmt(NextField2(Buf, p, Sep))
'
'         lMontoAfecto = lMontoAfecto - lMontoNetoActFijo    'Victor Morales 14 nov 2019
'         lMontoIVARec = lMontoIVARec - lMontoIVAActFijo
'
'         'Impto. Sin Derecho a Credito;IVA No Retenido;
'         lMontoImpSinDerechoCred = vFmt(NextField2(Buf, p, Sep))
'         lMontoIVANoRet = vFmt(NextField2(Buf, p, Sep))
'
'         'Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;
'         lTabacosPuros = vFmt(NextField2(Buf, p, Sep))
'         lTabacosCigarrillos = vFmt(NextField2(Buf, p, Sep))
'         lTabacosElaborados = vFmt(NextField2(Buf, p, Sep))
'
'         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
'         lNCE_NDE_FactCompra = Trim(NextField2(Buf, p, Sep))
'
'      Else
'
'         If lCodDocDTESII <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza tipo de documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         'tipo de compra: Del Giro
'         Txt = NextField2(Buf, p, Sep)
'
'         'RUT Proveedor
'         If lRUTEntidad <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza RUT con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Razón Social
'         If lRazonSocialEntidad <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Razón Social con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Folio
'         If lNumDoc <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Nro. Documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Fecha Docto
'         If lFechaEmision <> Int(GetDate(Trim(NextField2(Buf, p, Sep)))) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Fecha Documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         For F = 8 To 24
'            Txt = NextField2(Buf, p, Sep)   'saltamos los campos hasta llegar a los otros impuestos codificados
'         Next F
'
'      End If
'
'      lOtrosImp(r).CodSIIDTE = Trim(NextField2(Buf, p, Sep))
'      lOtrosImp(r).Valor = vFmt(NextField2(Buf, p, Sep))
'      lOtrosImp(r).Tasa = vFmt(NextField2(Buf, p, Sep))
'
'      If lOtrosImp(r).CodSIIDTE <> "" And lOtrosImp(r).CodSIIDTE <> "0" Then
'         IdTipoValLib = GetTipoTipoValLibFromCodSIIDTE(lTipoLib, lOtrosImp(r).CodSIIDTE)
'         If IdTipoValLib <= 0 Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Código otro impuesto inválido.")
'            DocErr = True
'
'         ElseIf Not ValidaOtroImp(lOtrosImp(r).CodSIIDTE) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Código de impuesto inválido.")
'            DocErr = True
'         End If
'      End If
'
''NextRec:
'
'   Next
'   'Loop
'
'
'   Close #Fd
'
'   Grid.FlxGrid.Redraw = True
'
'   Me.MousePointer = vbDefault
'
'   If NRecErroneos = 0 Then
'      If r = 1 Then
'         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregó " & r & " documento.", vbInformation + vbOKOnly
'         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
'      ElseIf r > 1 Then
'         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos.", vbInformation + vbOKOnly
'         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
'      Else  ' r=0
'         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & "- No se agregaron documentos.", vbInformation + vbOKOnly
'      End If
'
'   Else
'      If NRecErroneos > 1 Then
'         StrNRecErroneos = "- Se encontraron " & NRecErroneos & " registros con errores en el archivo."
'      Else
'         StrNRecErroneos = "- Se encontró " & NRecErroneos & " registro con errores en el archivo."
'      End If
'
'      If r = 1 Then
'         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregó " & r & " documento.", vbInformation + vbOKOnly
'         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
'      ElseIf r > 1 Then
'         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos.", vbInformation + vbOKOnly
'         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
'      Else  ' r=0
'         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- No se agregaron documentos.", vbInformation + vbOKOnly
'      End If
'
'      If MsgBox1("¿Desea revisar el log de importación " & lFNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'         Call ShellExecute(Me.hWnd, "open", lFNameLogImp, "", "", SW_SHOW)
'      End If
'   End If
'
'   Import_LibroRetencionesSIIAuto1 = True

End Function


Public Function Import_LibroVentasSII_OLD(Frm As Form, ByVal FName As String, Ano As Integer, Mes As Integer) As Boolean
   Dim i As Integer, j As Integer, k As Integer, n As Integer, r As Integer, l As Integer, F As Integer
   Dim p As Long
   Dim FNameLogImp As String
   Dim FNameTmp As String
   Dim MaxRegLibComp As Integer
   Dim Fd As Long
   Dim Sep As String
   Dim Buf As String
   Dim CorrelativoDoc As Integer
   Dim NewDoc As Boolean
   Dim Txt As String
   Dim DocErr As Integer
   Dim NDocsConError As Integer, NDocsOK As Integer
   Dim MsgDocsOK As String
   Dim RazonSocial As String, NotValidRut As Boolean
   Dim FechaAcuse8Dias As Long
   Dim MsgAcuse As Boolean
   Dim MsgTot As String
   Dim Rc As Integer
   Dim AuxStr As String, AuxMonto As Double
   Dim CodSucursal As String
   Dim IdTipoValLib As Integer
   
   
   lTipoLib = LIB_VENTAS
       
   Import_LibroVentasSII_OLD = False   'error
   On Error Resume Next
   
   FNameLogImp = gImportPath & "\Log\ImpLibCompSII-" & Format(Now, "yyyymmdd") & ".log"
   
   MaxRegLibComp = 3000
   On Error Resume Next
   
   Rc = LineCount(FName, MaxRegLibComp)
   If ERR Then
      MsgErr FName
      Import_LibroVentasSII_OLD = -ERR
      Exit Function
   End If
   
   If gDbType = SQL_ACCESS Then
    If Rc < 0 Then
       MsgBox1 "El archivo tiene más registros de los que soporta el sistema para el libro de compras o ventas (máx. " & MaxRegLibComp & " documentos)." & vbCrLf & vbCrLf & "No es posible realizar el proceso de importación.", vbExclamation
       Exit Function
    End If
   End If
   
   If Not LoadCuentasDef Then
      Exit Function
   End If
   
   
   'convertimos el archivo de Unix a DOS porque así lo entrega el SII
   Rc = ConvUnix2DosFile(FName, FName & "_")
   
   If Rc < 0 Then    'hubo un error al leer el archivo
      Exit Function
   ElseIf Rc = 0 Then   ' no lo convirtió porque ya está en formato DOS, se usa archivo original
      FNameTmp = ""
   ElseIf Rc > 0 Then   'lo convirtió y generó arcgivo temporal con guión bajo al final. Leemos del archivo temporal y luego lo borramos
      FName = FName & "_"
      FNameTmp = FName
   End If
      
   'abrimos el archivo
   Fd = FreeFile
   Open FName For Input As #Fd
   If ERR Then
      MsgErr FName
      Import_LibroVentasSII_OLD = -ERR
      If FNameTmp <> "" Then   'generamos uno temporal
         Kill FNameTmp            'lo borramos
      End If
      Exit Function
   End If
      
   r = 0
   Sep = ";"
   NewDoc = False
   NDocsOK = 0
   CorrelativoDoc = 0
   MsgAcuse = False
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS OFF")   'para que trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
         
   'Campos: Nro;Tipo Doc;Tipo Compra;RUT Cliente;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;Impto. Sin Derecho a Credito;IVA No Retenido;Tabacos Puros;Tabacos Cigarrillos;Tabacos Elaborados;NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         
   Do Until EOF(Fd)
         
      Line Input #Fd, Buf
      l = l + 1
               
      Buf = Trim(Buf)
      
      '1er registro con nombres de campos
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 Then
         GoTo NextRec
      End If
      
      p = 1
   
      'ahora leemos los documentos y los insertamos uno por uno
      n = vFmt(NextField2(Buf, p, Sep))         'número correlativo de registro. Viene en blanco si continúa documento anterior
      If n = 0 Then
         If CorrelativoDoc = 0 Then
            Call AddLogImp(FNameLogImp, FName, l, "Falta número de registro.")
            DocErr = True
         Else   'es continuación del documento anterior
            NewDoc = False
            r = r + 1         'registro i-esimo del documento NReg
         End If
         
      Else                    'documento nuevo
      
         If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
            If Not DocErr Then
               If ValidaTotalesVentas(MsgTot) Then
                  If GenDocumento Then       'puede que no se agregue si ya existe
                     NDocsOK = NDocsOK + 1
                     If MsgTot <> "" Then
                        Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: " & MsgTot)
                     End If
                  Else
                     Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Advertencia: Documento ya existe.")
                  End If
               Else
                  DocErr = True
                  NDocsConError = NDocsConError + 1
                  Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, MsgTot)
                                                         
               End If
            Else
               NDocsConError = NDocsConError + 1
            End If
      
         End If
         
         DocErr = False
         CorrelativoDoc = n
         r = 0                'primer registro del documento NReg
         
         'Inicializaciones
               
         'Nro;Tipo Doc;Tipo Compra;RUT Cliente;Razon Social;Folio;Fecha Docto;Fecha Recepcion;Fecha Acuse;
         lCodDocDTESII = ""
         lTipoDoc = 0
         lDTE = False
         lDelGiro = 0
         lRUTEntidad = ""
         lRazonSocialEntidad = ""
         lIdEntidad = 0
         RazonSocial = ""
         NotValidRut = False
         lNumDoc = ""
         lFechaEmision = 0
         lFechaRec = 0
         lFechaAcuse = 0
         lFechaReclamo = 0
         FechaAcuse8Dias = 0
         lIdSucursal = 0
         
         'Monto Exento;Monto Neto;Monto IVA Recuperable;Monto Iva No Recuperable;Codigo IVA No Rec.;
         lMontoExento = 0
         lMontoAfecto = 0
         lMontoIVA = 0
         
         lMontoIVARetTotal = 0    'IVA Retenido total
         lMontoIVARetParcial = 0  'IVA Retenido parciaL
         lMontoIVANoRet = 0       'IVA NO Retenido
         lMontoIVAPropio = 0      'IVA Propio
         lMontoIVATerceros = 0    'IVA Terceros
         
         lCredEmpContructora = 0
         lMontoTotalPeriodo = 0
         
         'Monto Total;Monto Neto Activo Fijo;IVA Activo Fijo;IVA uso Comun;
         lMontoTotal = 0
                     
         'NCE o NDE sobre Fact. de Compra;Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         lNCE_NDE_FactCompra = ""
         
         For k = 0 To UBound(lOtrosImp)
            lOtrosImp(k).CodSIIDTE = ""
            lOtrosImp(k).Tasa = 0
            lOtrosImp(k).Valor = 0
         Next k
         
      End If
        
        
      If r = 0 Then  'nuevo documento
      
         lCodDocDTESII = Trim(NextField2(Buf, p, Sep))
         If lCodDocDTESII = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta el tipo de documento.")
            DocErr = True
         End If
         
         lDTE = True
         lTipoDoc = GetTipoDocFromCodDocDTESII(lTipoLib, lCodDocDTESII)
         If lTipoDoc = 0 Then
            lDTE = False
            lTipoDoc = GetTipoDocFromCodDocSII(lTipoLib, lCodDocDTESII)
            If lTipoDoc = 0 Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Tipo de documento inválido.")
               DocErr = True
            End If
         End If
         
         
         'tipo de compra: Del Giro
         lDelGiro = IIf(Trim(NextField2(Buf, p, Sep)) = "Del Giro", True, False)
         
         'RUT Cliente
         lRUTEntidad = Trim(NextField2(Buf, p, Sep))
         If Not ValidRut(lRUTEntidad) Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "RUT inválido.")
            DocErr = True
         End If
         
            
         'Razón Social
         lRazonSocialEntidad = Trim(NextField2(Buf, p, Sep))
         If lRazonSocialEntidad = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta razón social.")
            DocErr = True
         End If
         
         lRazonSocialEntidad = Utf82Ansi(Utf82Ansi(lRazonSocialEntidad))  ' se hace dos veces porque al leerlo como texto duplica los caracteres (raro)
                 
         lIdEntidad = GetIdEntidad(lRUTEntidad, RazonSocial, NotValidRut)
         
         'Folio
         lNumDoc = Trim(NextField2(Buf, p, Sep))
         If lNumDoc = "" Then
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Falta el número de documento.")
            DocErr = True
         End If
   
         'Fecha Docto (Fecha Emision)
         lFechaEmision = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         
         'Fecha Recepcion (esta fecha es de recepción en el SII y no nos sirve)
         lFechaRec = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         'no Aplica
         
         'Fecha Acuse  (esta será la fecha de recepción para nuestro sistema, de acuerdo a lo indicado por Thomson)
         'no Aplica
         lFechaAcuse = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
         
         'Fecha Reclamo
         'No Aplica
         lFechaReclamo = Int(GetDate(Trim(NextField2(Buf, p, Sep))))
              
         
         'Monto Exento;Monto Neto;Monto IVA ;Monto Total
         lMontoExento = vFmt(NextField2(Buf, p, Sep))
         lMontoAfecto = vFmt(NextField2(Buf, p, Sep))
         lMontoIVA = vFmt(NextField2(Buf, p, Sep))
         lMontoTotal = vFmt(NextField2(Buf, p, Sep))
         
         'Monto Exento;Monto Neto;Monto IVA Retenido Total ;Monto IVA Retenido Parcial; IVA no Retenido;
         lMontoIVARetTotal = vFmt(NextField2(Buf, p, Sep))
         lMontoIVARetParcial = vFmt(NextField2(Buf, p, Sep))
         
         'IVA Propio; IVA Terceros
         'No aplica por ahora  (A18, A19)
         lMontoIVANoRet = vFmt(NextField2(Buf, p, Sep))
         lMontoIVANoRet = 0    'por cálculos en GenMovDocumento
         lMontoIVAPropio = vFmt(NextField2(Buf, p, Sep))
         lMontoIVATerceros = vFmt(NextField2(Buf, p, Sep))
         
         'Rut emisor Liq. Factura; Neto Comisión Liq. Factura; Exento Comisión Liq. Factura; IVA Comisión Liq. Factura;
         'No aplican por ahora (A20 -> A23)
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         
         'IVA fuera de plazo; TipoDoc Referencia; Folio Doc Referencia; Num. Ident. Receptor Extranjero; Nacionalidad Receptor Extranjero
         'No aplican por ahora (A25 -> A28)
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))

         'Credito empresa constructora (A29)
         lCredEmpContructora = vFmt(NextField2(Buf, p, Sep))
         
         'Impto. Zona Franca (Ley 18211); Garantia Dep. Envases; Indicador Venta sin Costo; Indicador Servicio Periodico; Monto No facturable
         'No aplican por ahora (A30 -> A34)
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxMonto = vFmt(NextField2(Buf, p, Sep))
         
         'Total Monto Periodo (A35)
         lMontoTotalPeriodo = vFmt(NextField2(Buf, p, Sep))
         
         'Venta Pasajes Transporte Nacional; Venta Pasajes Transporte Internacional; Numero Interno
         'No aplican por ahora (A36 -> A38)
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
         AuxStr = Trim(NextField2(Buf, p, Sep))
                 
         'Codigo Sucursal
         CodSucursal = Trim(NextField2(Buf, p, Sep))
         lIdSucursal = GetIdSucursal(CodSucursal)
         
         
         'NCE o NDE sobre Fact. de Compra; Codigo Otro Impuesto;Valor Otro Impuesto;Tasa Otro Impuesto
         lNCE_NDE_FactCompra = Trim(NextField2(Buf, p, Sep))
            
         'Codigo Otro Impuesto; Valor Otro Impuesto; Tasa Otro Impuesto
         lCodOtroImp = Trim(NextField2(Buf, p, Sep))
         lMontoOtroImp = vFmt(NextField2(Buf, p, Sep))
         lTasaOtroImp = vFmt(NextField2(Buf, p, Sep))
         
         If lCodOtroImp <> "" Then
            IdTipoValLib = GetTipoTipoValLibFromCodSIIDTE(lTipoLib, lCodOtroImp)
            If IdTipoValLib <= 0 Then
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Código otro impuesto inválido.")
               DocErr = True
            End If
         End If
          
'      Else
'
'         If lCodDocDTESII <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza tipo de documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         'tipo de compra: Del Giro
'         Txt = NextField2(Buf, p, Sep)
'
'         'RUT Cliente
'         If lRUTEntidad <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza RUT con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Razón Social
'         If lRazonSocialEntidad <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Razón Social con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Folio
'         If lNumDoc <> Trim(NextField2(Buf, p, Sep)) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Nro. Documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         'Fecha Docto
'         If lFechaEmision <> Int(GetDate(Trim(NextField2(Buf, p, Sep)))) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "No calza Fecha Documento con datos documento actual.")
'            DocErr = True
'         End If
'
'         For F = 8 To 24
'            Txt = NextField2(Buf, p, Sep)   'saltamos los campos hasta llegar a los otros impuestos codificados
'         Next F
     
      End If
      
'      lOtrosImp(r).CodSIIDTE = Trim(NextField2(Buf, p, Sep))
'      lOtrosImp(r).Valor = vFmt(NextField2(Buf, p, Sep))
'      lOtrosImp(r).Tasa = vFmt(NextField2(Buf, p, Sep))
'
'      If lOtrosImp(r).CodSIIDTE <> "" And lOtrosImp(r).CodSIIDTE <> "0" Then
'         If Not ValidaOtroImp(lOtrosImp(r).CodSIIDTE) Then
'            Call AddLogImp(FNameLogImp, fname, CorrelativoDoc, "Código de impuesto inválido.")
'            DocErr = True
'         End If
'      End If
      
      
NextRec:

   Loop
   
   Close #Fd

   
   'agregamos el último documento, si hay
   If CorrelativoDoc > 0 Then       'hay un documento antes, lo insertamos
      If Not DocErr Then
          If ValidaTotalesVentas(MsgTot) Then
            If GenDocumento Then       'puede que no se agregue si ya existe
               NDocsOK = NDocsOK + 1
            Else
               Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, "Documento ya existe.")
            End If
         Else
            DocErr = True
            NDocsConError = NDocsConError + 1
            Call AddLogImp(FNameLogImp, FName, CorrelativoDoc, MsgTot)
         End If
      Else
         NDocsConError = NDocsConError + 1
      End If
   
   End If
   
EndFnc:

   If NDocsOK > 1 Then
      MsgDocsOK = "Se importaron " & NDocsOK & " documentos nuevos."
   ElseIf NDocsOK = 1 Then
      MsgDocsOK = "Se importó un documento nuevo."
   Else
      MsgDocsOK = "No se importaron documentos nuevos."
   End If
      
   If NDocsConError = 0 Then
      MsgBox1 "Proceso de importación finalizado." & vbNewLine & vbNewLine & MsgDocsOK, vbInformation + vbOKOnly
      
   Else
      If NDocsConError = 1 Then
         MsgBox1 "Se encontró un documento con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
      Else
         MsgBox1 "Se encontraron " & NDocsConError & " documentos con errores en el archivo de importación." & vbNewLine & vbNewLine & MsgDocsOK, vbExclamation + vbOKOnly
      End If
      
      If MsgBox1("¿Desea revisar el log de importación " & FNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Frm.hWnd, "open", FNameLogImp, "", "", SW_SHOW)
      End If
      
   End If
   
   If gDbType = SQL_SERVER Then
      Call ExecSQL(DbMain, "SET ANSI_WARNINGS ON")   'para que no trunque la glosa en forma automática, sin error (truncar glosa en silencio) FCA 26 sept 2019
   End If
   
   If FNameTmp <> "" Then   'generamos uno temporal
      Kill FNameTmp            'lo borramos
   End If

End Function
Private Function GenDocumento() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim IdDoc As Long
   Dim i As Integer
   Dim EsRebaja As Boolean
   Dim IdCtaAfectoEntidad As Long, IdCtaExentoEntidad As Long, IdCtaTotalEntidad As Long, IdPropIVAEntidad As Integer, NombreEntidad As String
   Dim Giro As Boolean
   Dim FldArray(5) As AdvTbAddNew_t
   
   GenDocumento = False
   
   'primero validamos que exista el documento, si ya existe no se agrega ni se actualiza, dado que se importa desde el SII y este documento no debería cambiar
   Q1 = "SELECT IdDoc FROM Documento "
   Q1 = Q1 & " WHERE TipoLib = " & lTipoLib & " AND TipoDoc = " & lTipoDoc
   Q1 = Q1 & " AND (IdEntidad = " & lIdEntidad & " OR RutEntidad = '" & lRUTEntidad & "')"    'en rigor si la entidad no existe, tampoco existe el documento
   Q1 = Q1 & " AND NumDoc = '" & lNumDoc & "'"
'   If lCodDocDTESII = "30" Or lCodDocDTESII = "32" Then
      Q1 = Q1 & " AND iif(DTE <> 0, 1, 0) = " & Abs(lDTE)
'   Else
'      Q1 = Q1 & " AND DTE <> 0 "                                                              'deberían ser sólo DTE
'   End If
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
    
   If Not Rs.EOF Then      'documento ya existe
      
      '3008736
      If lTipoLib = LIB_COMPRAS Then
        Q1 = "UPDATE Documento SET FEmision = " & lFechaAcuse
      ElseIf lTipoLib = LIB_VENTAS Then
        Q1 = "UPDATE Documento SET FEmision = " & lFechaEmision
      End If
      '3008736
      
      'Q1 = "UPDATE Documento SET FEmision = " & lFechaEmision
      Q1 = Q1 & "  WHERE NumDoc = '" & lNumDoc & "'"
      Q1 = Q1 & " AND TipoLib = " & lTipoLib & " AND TipoDoc = " & lTipoDoc
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Q1 = Q1 & " AND IdDoc = " & vFld(Rs("IdDoc"))
      Q1 = Q1 & " AND (IdEntidad = " & lIdEntidad & " OR RutEntidad = '" & lRUTEntidad & "')"
      
      Call ExecSQL(DbMain, Q1)
      Call CloseRs(Rs)
      Exit Function
   End If
   
   Call CloseRs(Rs)
   
   If lIdEntidad = 0 Then
      If Not AddEntidad(lRUTEntidad, lRazonSocialEntidad, lIdEntidad) Then
         Exit Function
      End If
   End If
         
   'Insert
'   Set Rs = DbMain.OpenRecordset("Documento")
'   Rs.AddNew
'
'   IdDoc = vFld(Rs("IdDoc"))
'   Rs.Fields("IdUsuario") = gUsuario.IdUsuario
'   Rs.Fields("FechaCreacion") = CLng(Int(Now))
''   Rs.Fields("FEmision") = lFechaRec           'Fecha recepción
'   Rs.Fields("FEmision") = lFechaAcuse          'Fecha recepción utilizamos la fecha de Acuse del cliente
'   Rs.Fields("FEmisionOri") = lFechaEmision     'Fecha Docto
'
'   On Error Resume Next
'
'   For i = 1 To 10
'      Rs.Fields("NumDoc") = CLng(Rnd * 321654356#)
'
'      Rs.Update
'
'      If Err = 0 Then
'         Exit For
'      End If
'      Err.Clear
'   Next i
'
'   Rs.Close
'
'   Set Rs = Nothing
   
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
   
   FldArray(4).FldName = "FEmision"       'Fecha recepción utilizamos la fecha de Acuse del cliente  (no se usa la FechaRec)
   If lTipoLib = LIB_COMPRAS Then
      FldArray(4).FldValue = lFechaAcuse
   ElseIf lTipoLib = LIB_VENTAS Then
      FldArray(4).FldValue = lFechaEmision
   End If
   FldArray(4).FldIsNum = True
         
   FldArray(5).FldName = "FEmisionOri"    'Fecha Docto
   FldArray(5).FldValue = lFechaEmision
   FldArray(5).FldIsNum = True
         
   IdDoc = AdvTbAddNewMult(DbMain, "Documento", "IdDoc", FldArray)
            
   
   lIdDoc = IdDoc
   

   'y ahora llenamos el resto de los campos
   Q1 = "UPDATE Documento SET "
   Q1 = Q1 & "  TipoLib = " & lTipoLib
   Q1 = Q1 & ", TipoDoc = " & lTipoDoc
   Q1 = Q1 & ", CorrInterno = 0"
'   If lCodDocDTESII = "30" Or lCodDocDTESII = "32" Then
'      Q1 = Q1 & ", DTE = 0"            'caso especial (parche) por FAC y FCE no DTE
'   Else
      Q1 = Q1 & ", DTE = " & Abs(lDTE)           'debería ser sólo DTE
'   End If
   Q1 = Q1 & ", Giro = " & CInt(lDelGiro)

   'NumFiscImpr y NumInformeZ no se usan porque no se incluye este tipo de documentos en el reporte del SII
   'Tampoco se incluyen las boletas y las ventas sin documento   (Joshua Nicolás Catrin 3 may 2018)
   
   
   Q1 = Q1 & ", NumDoc = '" & lNumDoc & "'"
   Q1 = Q1 & ", NumDocHasta = '0'"
   Q1 = Q1 & ", CantBoletas = 0"

   Q1 = Q1 & ", IdEntidad = " & lIdEntidad
   Q1 = Q1 & ", EntRelacionada = 0"
   Q1 = Q1 & ", IdSucursal = " & lIdSucursal

   If lTipoLib = LIB_COMPRAS Then
      Q1 = Q1 & ", TipoEntidad = " & ENT_PROVEEDOR
      'por si acaso, ponemos la clasificación de la entidad
      Call ExecSQL(DbMain, "UPDATE Entidades SET Clasif" & ENT_PROVEEDOR & " = 1 WHERE IdEntidad = " & lIdEntidad & " AND IdEmpresa = " & gEmpresa.id)
   Else
      Q1 = Q1 & ", TipoEntidad = " & ENT_CLIENTE
      'por si acaso, ponemos la clasificación de la entidad
      Call ExecSQL(DbMain, "UPDATE Entidades SET Clasif" & ENT_CLIENTE & " = 1 WHERE IdEntidad = " & lIdEntidad & " AND IdEmpresa = " & gEmpresa.id)
   End If

   
   Call GetCuentasEntidad(lTipoLib, lIdEntidad, IdCtaAfectoEntidad, IdCtaExentoEntidad, IdCtaTotalEntidad, IdPropIVAEntidad, Giro, NombreEntidad)  'Giro es sólo para libro de ventas
   
   If lTipoLib = LIB_COMPRAS Then
'      Q1 = Q1 & ", PropIVA = " & PIVA_SINPROP    'OJO Verificar con Thomson
      Q1 = Q1 & ", PropIVA = " & IdPropIVAEntidad
   End If

'   Q1 = Q1 & ", FEmision = " & lFechaAcuse            'Fecha recepción utilizamos la fecha de Acuse del cliente
'   Q1 = Q1 & ", FEmisionOri = " & lFechaEmision
   Q1 = Q1 & ", FVenc = " & CLng(DateAdd("d", 30, lFechaEmision))
   Q1 = Q1 & ", Exento = " & Abs(lMontoExento)

      
   lDefCtasProveedor = False
   If IdCtaAfectoEntidad <> 0 And IdCtaAfectoEntidad <> lIdCtaAfecto Then
      lIdCtaAfectoEntidad = IdCtaAfectoEntidad
      lDefCtasProveedor = True
   Else
      lIdCtaAfectoEntidad = lIdCtaAfecto
   End If
   
   If IdCtaExentoEntidad <> 0 And IdCtaExentoEntidad <> lIdCtaExento Then
      lIdCtaExentoEntidad = IdCtaExentoEntidad
      lDefCtasProveedor = True
   Else
      lIdCtaExentoEntidad = lIdCtaExento
   End If
   
   If IdCtaTotalEntidad <> 0 And IdCtaTotalEntidad <> lIdCtaTotal Then
      lIdCtaTotalEntidad = IdCtaTotalEntidad
      lDefCtasProveedor = True
   Else
      lIdCtaTotalEntidad = lIdCtaTotal
   End If
   
   
   If lMontoExento <> 0 Then
      Q1 = Q1 & ", IdCuentaExento = " & lIdCtaExentoEntidad
   Else
      Q1 = Q1 & ", IdCuentaExento = 0"
   End If

   Q1 = Q1 & ", Afecto = " & Abs(lMontoAfecto)
   If lMontoAfecto <> 0 Then
      Q1 = Q1 & ", IdCuentaAfecto = " & lIdCtaAfectoEntidad
   Else
      Q1 = Q1 & ", IdCuentaAfecto = 0 "
   End If

   EsRebaja = gTipoDoc(GetTipoDoc(lTipoLib, lTipoDoc)).EsRebaja
   lDescrip = gTipoDoc(GetTipoDoc(lTipoLib, lTipoDoc)).Diminutivo & " " & lNumDoc & " - " & lRazonSocialEntidad

   Q1 = Q1 & ", IVA = " & Abs(lMontoIVARec)
   Q1 = Q1 & ", IdCuentaIVA = " & lIdCtaIVA
   If EsRebaja Then
      Q1 = Q1 & ", OtroImp = " & lMontoOtroImp * -1
   Else
      Q1 = Q1 & ", OtroImp = " & lMontoOtroImp     'valor 0 dado que no lo tenemos acumulado
   End If
   If lMontoOtroImp <> 0 Then
      Q1 = Q1 & ", IdCuentaOtroImp = " & lIdCtaOtrosImp
   End If
   
   Q1 = Q1 & ", Total = " & Abs(lMontoTotal)
   Q1 = Q1 & ", VentasAcumInfZ = 0"
   Q1 = Q1 & ", IdCuentaTotal = " & lIdCtaTotalEntidad
   Q1 = Q1 & ", Descrip = '" & ParaSQL(Left(lDescrip, 100)) & "'"
   Q1 = Q1 & ", IdANegCCosto = ' '"
   Q1 = Q1 & ", Estado = " & ED_PENDIENTE
   Q1 = Q1 & ", SaldoDoc = NULL"
   
   
   Q1 = Q1 & " WHERE IdDoc = " & lIdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Call ExecSQL(DbMain, Q1)
   
   'Documento asociado
   If lIdDocRef <> 0 Then
      
      Q1 = "UPDATE Documento SET IdDocAsoc = " & lIdDocRef & ", TipoDocAsoc = " & lTipoDocRef & ", NumDocAsoc = '" & lNumDocRef & "', DTEDocAsoc = " & lDTEDocRef
      Q1 = Q1 & "  WHERE IdDoc =" & lIdDoc
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

      Call ExecSQL(DbMain, Q1)
   
   End If
   

   Call GetANegCCostoEntidad(lTipoLib, lIdEntidad, lIdAreaNegAfectoEntidad, lIdAreaNegExentoEntidad, lIdAreaNegTotalEntidad, lIdCCostoAfectoEntidad, lIdCCostoExentoEntidad, lIdCCostoTotalEntidad)
   
   Call GenMovDocumento
   
   GenDocumento = True

End Function

Public Function LoadDefCuentasRet(TipoValor As Integer, TipoLib As Integer) As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Item As String
   
   LoadDefCuentasRet = 0
      
   If TipoLib > 0 Then
   
      Q1 = "SELECT CuentasBasicas.IdCuenta, Cuentas.Codigo, Cuentas.Nombre, Cuentas.Descripcion, TipoValor "
      Q1 = Q1 & " FROM CuentasBasicas INNER JOIN Cuentas ON CuentasBasicas.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "CuentasBasicas", "Cuentas")
      Q1 = Q1 & " WHERE TipoLib = " & TipoLib
      Q1 = Q1 & " AND TipoValor = " & TipoValor
      Q1 = Q1 & " AND CuentasBasicas.IdEmpresa = " & gEmpresa.id & " AND CuentasBasicas.Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY TipoValor, CuentasBasicas.Id "
      
      Set Rs = OpenRs(DbMain, Q1)
   
      Do While Rs.EOF = False
                      
         LoadDefCuentasRet = vFld(Rs("IdCuenta"))
                          
         Rs.MoveNext
         
      Loop
      Call CloseRs(Rs)
      
   End If
   
End Function


Private Function LoadCuentasDef() As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Item As String
   
   LoadCuentasDef = False
   
   lIdCtaAfecto = 0
   lIdCtaExento = 0
   lIdCtaTotal = 0
   
   lIdCtaIVA = 0
   lIdCtaIVAIrrec = 0
   lIdCtaOtrosImp = 0
   lIdCtaOtrosImpFacCompra = 0
   
   'cuentas Entidad
   lIdCtaAfectoEntidad = 0
   lIdCtaExentoEntidad = 0
   lIdCtaTotalEntidad = 0
   lDefCtasProveedor = False  'indoca si el proveedor tiene una cuenta definida distinta que la de omisión
   
   'area de negocio Entidad
   lIdAreaNegAfectoEntidad = 0
   lIdAreaNegExentoEntidad = 0
   lIdAreaNegTotalEntidad = 0
   
   'centro de costo entidad
   lIdCCostoAfectoEntidad = 0
   lIdCCostoExentoEntidad = 0
   lIdCCostoTotalEntidad = 0

            
   If lTipoLib > 0 Then
   
      Q1 = "SELECT IdCuenta, TipoValor "
      Q1 = Q1 & " FROM CuentasBasicas  "
      Q1 = Q1 & " WHERE TipoLib = " & lTipoLib
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY TipoValor, Id "
      
      Set Rs = OpenRs(DbMain, Q1)
   
      Do While Rs.EOF = False
                      
         Select Case vFld(Rs("TipoValor"))
         
            Case LIBVENTAS_AFECTO, LIBCOMPRAS_AFECTO
            
               If lIdCtaAfecto = 0 Then
                  lIdCtaAfecto = vFld(Rs("IdCuenta"))
               End If
               
            Case LIBVENTAS_EXENTO, LIBCOMPRAS_EXENTO
            
               If lIdCtaExento = 0 Then
                  lIdCtaExento = vFld(Rs("IdCuenta"))
               End If
            
            Case LIBVENTAS_TOTAL, LIBCOMPRAS_TOTAL
            
               If lIdCtaTotal = 0 Then
                  lIdCtaTotal = vFld(Rs("IdCuenta"))
               End If
               
         End Select
                           
         Rs.MoveNext
         
      Loop
      Call CloseRs(Rs)
      
   End If
   
   If lTipoLib = LIB_COMPRAS Then
      lIdCtaIVA = gCtasBas.IdCtaIVACred
      lIdCtaIVAIrrec = gCtasBas.IdCtaIVAIrrec
      lIdCtaOtrosImp = gCtasBas.IdCtaOtrosImpCred
      lIdCtaOtrosImpFacCompra = gCtasBas.IdCtaOtrosImpDeb
   Else
      lIdCtaIVA = gCtasBas.IdCtaIVADeb               'LIB_VENTAS
      lIdCtaIVAIrrec = 0
      lIdCtaOtrosImp = gCtasBas.IdCtaOtrosImpDeb
      lIdCtaOtrosImpFacCompra = gCtasBas.IdCtaOtrosImpCred
   End If
   
   If lIdCtaAfecto = 0 Or lIdCtaExento = 0 Or lIdCtaTotal = 0 Then
      MsgBox1 "Falta definir las cuentas básicas para los libros de compras, ventas y/o retenciones." & vbCrLf & vbCrLf & "Utilice el menú Configuración >> Configuración inicial >> Definir cuentas básicas", vbExclamation
      Exit Function
   End If
      
   If lIdCtaIVA = 0 Then
      MsgBox1 "Falta definir la cuenta de IVA." & vbCrLf & vbCrLf & "Utilice el menú Configuración >> Configuración inicial >> Definir cuentas básicas", vbExclamation
      Exit Function
   End If
   
   If lTipoLib = LIB_COMPRAS Then
      If lIdCtaIVAIrrec = 0 Then
         MsgBox1 "Falta definir las cuentas de IVA e IVA Irrecuperable." & vbCrLf & vbCrLf & "Utilice el menú Configuración >> Configuración inicial >> Definir cuentas básicas", vbExclamation
         Exit Function
      End If
   End If
   
   If lIdCtaOtrosImp = 0 Or lIdCtaOtrosImpFacCompra = 0 Then
      MsgBox1 "Falta definir las cuentas de Otros Impuestos." & vbCrLf & vbCrLf & "Utilice el menú Configuración >> Configuración inicial >> Definir cuentas básicas", vbExclamation
      Exit Function
   End If
   
   LoadCuentasDef = True
End Function

'EsDelGiro es sólo para libro de ventas
'PropIVA es sólo para compras
Public Sub GetCuentasEntidad(ByVal TipoLib As Integer, ByVal IdEntidad As Long, IdCtaAfecto As Long, IdCtaExento As Long, IdCtaTotal As Long, IdPropIVA As Integer, EsDelGiro As Boolean, NombreEntidad As String)
   Dim Rs As Recordset
   Dim Q1 As String
   
   If IdEntidad = 0 Then
      Exit Sub
   End If
   
   IdCtaAfecto = 0
   IdCtaExento = 0
   IdCtaTotal = 0
   EsDelGiro = 0
   IdPropIVA = 0
   NombreEntidad = ""
   
   'PropIVA, DelGiro y Nombre
   Q1 = "SELECT PropIVA, EsDelGiro, Entidades.Nombre FROM Entidades WHERE IdEntidad = " & IdEntidad
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      IdPropIVA = vFld(Rs("PropIVA"))
      EsDelGiro = vFld(Rs("EsDelGiro"))
      NombreEntidad = vFld(Rs("Nombre"))
   End If
   
   Call CloseRs(Rs)
   
   'Afecto
   If TipoLib = LIB_COMPRAS Then
      Q1 = "SELECT IdCuenta FROM Entidades INNER JOIN Cuentas ON Entidades.CodCtaAfecto = Cuentas.Codigo "
   Else
      Q1 = "SELECT IdCuenta FROM Entidades INNER JOIN Cuentas ON Entidades.CodCtaAfectoVta = Cuentas.Codigo "
   End If
   
   Q1 = Q1 & " AND Entidades.IdEmpresa = Cuentas.IdEmpresa "
   Q1 = Q1 & " WHERE IdEntidad = " & IdEntidad
   Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      IdCtaAfecto = vFld(Rs("IdCuenta"))
   End If

   Call CloseRs(Rs)
   
   'Exento
   If TipoLib = LIB_COMPRAS Then
      Q1 = "SELECT IdCuenta FROM Entidades INNER JOIN Cuentas ON Entidades.CodCtaExento = Cuentas.Codigo "
   Else
      Q1 = "SELECT IdCuenta FROM Entidades INNER JOIN Cuentas ON Entidades.CodCtaExentoVta = Cuentas.Codigo "
   End If
   
   Q1 = Q1 & " AND Entidades.IdEmpresa = Cuentas.IdEmpresa "
   Q1 = Q1 & " WHERE IdEntidad = " & IdEntidad
   Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      IdCtaExento = vFld(Rs("IdCuenta"))
   End If

   Call CloseRs(Rs)
   
   'Total
   If TipoLib = LIB_COMPRAS Then
      Q1 = "SELECT IdCuenta FROM Entidades INNER JOIN Cuentas ON Entidades.CodCtaTotal = Cuentas.Codigo "
   Else
      Q1 = "SELECT IdCuenta FROM Entidades INNER JOIN Cuentas ON Entidades.CodCtaTotalVta = Cuentas.Codigo "
   End If
   
   Q1 = Q1 & " AND Entidades.IdEmpresa = Cuentas.IdEmpresa "
   Q1 = Q1 & " WHERE IdEntidad = " & IdEntidad
   Q1 = Q1 & " AND Cuentas.IdEmpresa = " & gEmpresa.id & " AND Cuentas.Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      IdCtaTotal = vFld(Rs("IdCuenta"))
   End If

   Call CloseRs(Rs)
   
End Sub

Private Function GetIdCuentaImpAdic(ByVal TipoLib As Integer, ByVal TipoValLib As Integer) As Long
   Dim Rs As Recordset
   Dim Q1 As String
   
   If TipoValLib <= 0 Then
      GetIdCuentaImpAdic = 0
      Exit Function
   End If
   
   Q1 = "SELECT IdCuenta FROM ImpAdic WHERE TipoLib = " & TipoLib & " AND TipoValor = " & TipoValLib
   Q1 = Q1 & " AND ImpAdic.IdEmpresa = " & gEmpresa.id & " AND ImpAdic.Ano = " & gEmpresa.Ano
   Set Rs = OpenRs(DbMain, Q1)
   
   GetIdCuentaImpAdic = 0
   
   If Not Rs.EOF Then
      GetIdCuentaImpAdic = vFld(Rs("IdCuenta"))
   End If
   
   Call CloseRs(Rs)
   

End Function

Public Function GetANegCCostoEntidad(ByVal TipoLib As Integer, ByVal IdEntidad As Long, IdAreaNegAfectoEntidad As Long, IdAreaNegExentoEntidad As Long, IdAreaNegTotalEntidad As Long, IdCCostoAfectoEntidad As Long, IdCCostoExentoEntidad As Long, IdCCostoTotalEntidad As Long) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   
   IdAreaNegExentoEntidad = 0
   IdAreaNegAfectoEntidad = 0
   IdAreaNegTotalEntidad = 0
   IdCCostoExentoEntidad = 0
   IdCCostoAfectoEntidad = 0
   IdCCostoTotalEntidad = 0
   
   GetANegCCostoEntidad = False
   
   If IdEntidad = 0 Then
      Exit Function
   End If
   
   If TipoLib = LIB_COMPRAS Then
      Q1 = "SELECT CodAreaNegAfecto, CodAreaNegExento, CodAreaNegTotal, CodCCostoAfecto, CodCCostoExento, CodCCostoTotal "
      Q1 = Q1 & " FROM Entidades WHERE IdEntidad = " & IdEntidad
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
      
         IdAreaNegAfectoEntidad = GetAreaNegocio(vFld(Rs("CodAreaNegAfecto")))
         IdAreaNegExentoEntidad = GetAreaNegocio(vFld(Rs("CodAreaNegExento")))
         IdAreaNegTotalEntidad = GetAreaNegocio(vFld(Rs("CodAreaNegTotal")))
         
         IdCCostoAfectoEntidad = GetCentroCosto(vFld(Rs("CodCCostoAfecto")))
         IdCCostoExentoEntidad = GetCentroCosto(vFld(Rs("CodCCostoExento")))
         IdCCostoTotalEntidad = GetCentroCosto(vFld(Rs("CodCCostoTotal")))
         
         GetANegCCostoEntidad = True
      
      End If
   Else
      Q1 = "SELECT CodAreaNegAfectoVta, CodAreaNegExentoVta, CodAreaNegTotalVta, CodCCostoAfectoVta, CodCCostoExentoVta, CodCCostoTotalVta "
      Q1 = Q1 & " FROM Entidades WHERE IdEntidad = " & IdEntidad
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
      
         IdAreaNegAfectoEntidad = GetAreaNegocio(vFld(Rs("CodAreaNegAfectoVta")))
         IdAreaNegExentoEntidad = GetAreaNegocio(vFld(Rs("CodAreaNegExentoVta")))
         IdAreaNegTotalEntidad = GetAreaNegocio(vFld(Rs("CodAreaNegTotalVta")))
         
         IdCCostoAfectoEntidad = GetCentroCosto(vFld(Rs("CodCCostoAfectoVta")))
         IdCCostoExentoEntidad = GetCentroCosto(vFld(Rs("CodCCostoExentoVta")))
         IdCCostoTotalEntidad = GetCentroCosto(vFld(Rs("CodCCostoTotalVta")))
         
         GetANegCCostoEntidad = True
      
      End If
   End If
     
   Call CloseRs(Rs)
   
End Function

Private Function ValidaOtroImp(ByVal CodImpSII As String) As Boolean

   If CodImpSII = "" Or CodImpSII = "0" Then
      ValidaOtroImp = True
      Exit Function
   End If

   ValidaOtroImp = False
   
   Select Case lCodDocDTESII
      Case "33"
         If InStr(IMP_33, "," & CodImpSII & ",") <= 0 Then
            Exit Function
         End If
      
      Case "45", "46"
         If InStr(IMP_45_46, "," & CodImpSII & ",") <= 0 Then
            Exit Function
         End If
      
      Case "60", "61", "55", "56"
         If InStr(IMP_SUMA_60_61_55_56, "," & CodImpSII & ",") <= 0 And InStr(IMP_RESTA_60_61_55_56, "," & CodImpSII & ",") <= 0 Then
            Exit Function
         End If
   End Select
   
   ValidaOtroImp = True
   
End Function
Private Function EsIVARetenido(ByVal TipoLib As Integer, ByVal CodImpSII As String) As Boolean
   
   EsIVARetenido = False
   
   If CodImpSII = "" Or CodImpSII = "0" Then
      Exit Function
   End If

   If TipoLib = LIB_COMPRAS Then
      If lCodDocDTESII = "45" Or lCodDocDTESII = "46" Or lCodDocDTESII = "60" Or lCodDocDTESII = "61" Then
         If InStr(ES_IVARETENIDO, "," & CodImpSII & ",") > 0 Then
            EsIVARetenido = True
         End If
      End If
      
   End If
   
End Function

Private Function ValidaTotalesCompras(Msg As String) As Boolean
   Dim Total As Double
   Dim TotOtroImp As Double, TotOtros As Double
   Dim k As Integer
   Dim NotaCredDeb As Boolean

   ValidaTotalesCompras = False
   
   
   TotOtros = lMontoIVANoRec + lMontoImpSinDerechoCred + lMontoIVANoRet + lTabacosPuros + lTabacosCigarrillos + lTabacosElaborados
         
   If InStr(",60,61,55,56,", "," & lCodDocDTESII & ",") > 0 Then
      NotaCredDeb = True
   End If
   
   For k = 0 To UBound(lOtrosImp)
   
      If lOtrosImp(k).CodSIIDTE <> "" And lOtrosImp(k).CodSIIDTE <> "0" Then
      
         If NotaCredDeb Then
            If InStr(IMP_RESTA_60_61_55_56, "," & lOtrosImp(k).CodSIIDTE & ",") > 0 Then
               TotOtroImp = TotOtroImp - lOtrosImp(k).Valor
            Else
               TotOtroImp = TotOtroImp + lOtrosImp(k).Valor
            End If
         Else
            TotOtroImp = TotOtroImp + lOtrosImp(k).Valor
         End If
         
      Else
         Exit For
         
      End If
   Next k
   
   Select Case lCodDocDTESII
      Case "30", "33"
         'Total Documento = Monto Exento + Monto Neto + Monto IVA Recuperable + Monto No Recuperable + IVA Activo Fijo + IVA Uso Común + Impto. Sin Derecho a Credito + Tabacos Puros + Tabacos Cigarrillos + Tabacos Elaborados + Valor Otro Impuesto.
         Total = lMontoExento + lMontoAfecto + lMontoNetoActFijo + lMontoIVARec + lMontoIVAActFijo + lMontoIVAUsoComun + TotOtros + TotOtroImp
         Msg = "El Monto Total ($ " & Format(Total, NUMFMT) & ") debe cuadrar con la suma del Monto Neto, Exento, IVA, Monto Neto Act. Fijo, IVA Act. Fijo, IVA No Recup., IVA Uso Común, Otros Impuestos s/crédito e Imp. Adicionales."
      
      Case "32", "34"
         'Total Documento = Monto Exento
         Total = lMontoExento
         Msg = "El Monto Total($ " & Format(Total, NUMFMT) & ") debe cuadrar con el Monto Exento."
      
      Case "45", "46"
         'Total Documento = Monto Exento + Monto Neto + Monto IVA Recuperable + Monto No Recuperable + IVA Activo Fijo + IVA Uso Común + Impto. Sin Derecho a Credito + Tabacos Puros + Tabacos Cigarrillos + Tabacos Elaborados - Valor Otro Impuesto.
         Total = lMontoExento + lMontoAfecto + lMontoNetoActFijo + lMontoIVARec + lMontoIVAActFijo + lMontoIVAUsoComun + TotOtros - TotOtroImp
         Msg = "El Monto Total ($ " & Format(Total, NUMFMT) & ") debe cuadrar con la suma del Monto Neto, Exento, IVA, Monto Neto Act. Fijo, IVA Act. Fijo, IVA No Recup., IVA Uso Común, Otros Impuestos s/crédito menos Imp. que corresponden a retenciones."
               
      Case "60", "61", "55", "56"
         'Total Documento = Monto Exento + Monto Neto + Monto IVA Recuperable + Monto No Recuperable + IVA Activo Fijo + IVA Uso Común + Impto. Sin Derecho a Credito + Tabacos Puros + Tabacos Cigarrillos + Tabacos Elaborados ± Valor Otro Impuesto.         If InStr(IMP_SUMA_60_61_55_56, "," & CodImpSII & ",") <= 0 And InStr(IMP_RESTA_60_61_55_56, "," & CodImpSII & ",") <= 0 Then
         Total = lMontoExento + lMontoAfecto + lMontoNetoActFijo + lMontoIVARec + lMontoIVAActFijo + lMontoIVAUsoComun + TotOtros + TotOtroImp
         Msg = "El Monto Total ($ " & Format(Total, NUMFMT) & ") debe cuadrar con la suma del Monto Neto, Exento, IVA, Monto Neto Act. Fijo, IVA Act. Fijo, IVA No Recup., IVA Uso Común, Otros Impuestos s/crédito e Imp. Adicionales menos Imp. Que corresponden a retenciones."
         
      Case Else
         Total = lMontoTotal   'siempre lo damos por válido
         Msg = ""
         
   End Select
   
   
   If Total = lMontoTotal Then
      ValidaTotalesCompras = True
   End If

End Function

Private Function ValidaTotalesVentas(Msg As String) As Boolean
   Dim Total As Double
   Dim TotOtroImp As Double, TotOtros As Double
   Dim k As Integer
   Dim NotaCredDeb As Boolean

   ValidaTotalesVentas = False
   
   Msg = ""
      
'   Select Case lCodDocDTESII
'      Case "30", "33"
'         'Total Documento = Monto Exento + Monto Neto + Monto IVA Recuperable + Monto No Recuperable + IVA Activo Fijo + IVA Uso Común + Impto. Sin Derecho a Credito + Tabacos Puros + Tabacos Cigarrillos + Tabacos Elaborados + Valor Otro Impuesto.
'         Total = lMontoExento + lMontoAfecto + lMontoNetoActFijo + lMontoIVARec + lMontoIVAActFijo + lMontoIVAUsoComun + TotOtros + TotOtroImp
'         Msg = "El Monto Total ($ " & Format(Total, NUMFMT) & ") debe cuadrar con la suma del Monto Neto, Exento, IVA, Monto Neto Act. Fijo, IVA Act. Fijo, IVA No Recup., IVA Uso Común, Otros Impuestos s/crédito e Imp. Adicionales."
'
'      Case "32", "34"
'         'Total Documento = Monto Exento
'         Total = lMontoExento
'         Msg = "El Monto Total($ " & Format(Total, NUMFMT) & ") debe cuadrar con el Monto Exento."
'
'      Case "45", "46"
'         'Total Documento = Monto Exento + Monto Neto + Monto IVA Recuperable + Monto No Recuperable + IVA Activo Fijo + IVA Uso Común + Impto. Sin Derecho a Credito + Tabacos Puros + Tabacos Cigarrillos + Tabacos Elaborados - Valor Otro Impuesto.
'         Total = lMontoExento + lMontoAfecto + lMontoNetoActFijo + lMontoIVARec + lMontoIVAActFijo + lMontoIVAUsoComun + TotOtros - TotOtroImp
'         Msg = "El Monto Total ($ " & Format(Total, NUMFMT) & ") debe cuadrar con la suma del Monto Neto, Exento, IVA, Monto Neto Act. Fijo, IVA Act. Fijo, IVA No Recup., IVA Uso Común, Otros Impuestos s/crédito menos Imp. que corresponden a retenciones."
'
'      Case "60", "61", "55", "56"
'         'Total Documento = Monto Exento + Monto Neto + Monto IVA Recuperable + Monto No Recuperable + IVA Activo Fijo + IVA Uso Común + Impto. Sin Derecho a Credito + Tabacos Puros + Tabacos Cigarrillos + Tabacos Elaborados ± Valor Otro Impuesto.         If InStr(IMP_SUMA_60_61_55_56, "," & CodImpSII & ",") <= 0 And InStr(IMP_RESTA_60_61_55_56, "," & CodImpSII & ",") <= 0 Then
'         Total = lMontoExento + lMontoAfecto + lMontoNetoActFijo + lMontoIVARec + lMontoIVAActFijo + lMontoIVAUsoComun + TotOtros + TotOtroImp
'         Msg = "El Monto Total ($ " & Format(Total, NUMFMT) & ") debe cuadrar con la suma del Monto Neto, Exento, IVA, Monto Neto Act. Fijo, IVA Act. Fijo, IVA No Recup., IVA Uso Común, Otros Impuestos s/crédito e Imp. Adicionales menos Imp. Que corresponden a retenciones."
'
'      Case Else
'         Total = lMontoTotal   'siempre lo damos por válido
'         Msg = ""
'
'   End Select
   
   Total = lMontoAfecto + lMontoExento + lMontoIVA + lMontoOtroImp - lMontoIVARetTotal - lMontoIVARetParcial - lCredEmpContructora
   
   If Total <> lMontoTotal Then
      Msg = "El Monto Total ($ " & Format(Total, NUMFMT) & ") no cuadra con la suma del Monto Neto, Exento, IVA, Otros Impuestos mesos IVA Retenido Total, IVA Retenido Parcial y Crédito Empresa Constructora."
   End If
  
'   If Total = lMontoTotal Then
      ValidaTotalesVentas = True
'   End If

End Function


Private Sub GenMovDocumento()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim QBase As String
   Dim i As Integer, r As Integer
   Dim Glosa As String
   Dim TipoDocNC As Boolean
   Dim Idx As Integer
   Dim TipoIVAIrrec As Integer
   Dim ValOtroImp As Double
   Dim IdTipoValLib As Integer
   
   Dim Exento As Double
   Dim Afecto As Double
   Dim IVA As Double
   Dim ValOtros As Double
   Dim OtroImp As Double
   Dim RetParcial As Double
   
   Dim IVAIrrec As Integer
   Dim HayOtrosImp As Boolean, HayANegCcosto As Boolean
   Dim IdCuenta As Long
   Dim EsIVAIrrec As Boolean
   
   Dim HaySoloIVARetenido As Boolean
   Dim OImpEsIVARetenido As Boolean
   Dim OImpTipoIVARetenido As Integer
   Dim nIVARetenido As Integer
   Dim nIVANoRetenido As Integer
   Dim TipoDocNCF As Integer
   Dim Tasa As Single, EsRecuperable As Boolean


   QBase = "INSERT INTO MovDocumento"
   QBase = QBase & "(IdDoc, IdEmpresa, Ano, Orden, IdCuenta, Debe, Haber, Glosa, IdTipoValLib, EsTotalDoc, IdAreaNeg, IdCCosto, Tasa, EsRecuperable, CodSIIDTE) "
   QBase = QBase & " VALUES(" & lIdDoc & "," & gEmpresa.id & "," & gEmpresa.Ano & ","


   Glosa = ParaSQL(Left(lDescrip, 50))
            
   EsIVAIrrec = False

   Idx = GetTipoDoc(lTipoLib, lTipoDoc)
   If Idx >= 0 Then
      TipoDocNC = gTipoDoc(Idx).EsRebaja
   End If

   i = 1

   If lTipoLib = LIB_COMPRAS Then

      'Exento

      If lMontoExento <> 0 Then
         Q1 = QBase & i & ","                                     'Orden
         Q1 = Q1 & lIdCtaExentoEntidad & ","                      'IdCuenta

         If TipoDocNC Then
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(lMontoExento) & ","  'Haber
         Else
            Q1 = Q1 & Abs(lMontoExento) & ","  'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                             'Glosa
         Q1 = Q1 & LIBCOMPRAS_EXENTO & ","                        'IdTipoValLib
         Q1 = Q1 & "0,"                                           'EsTotalDoc
         Q1 = Q1 & lIdAreaNegExentoEntidad & ","                  'IdAreaNeg
         Q1 = Q1 & lIdCCostoExentoEntidad & ","                   'IdCCosto
         Q1 = Q1 & "0,0,' '" & ")"                                'Tasa, EsRecuperable, CodSIIDTE

         Call ExecSQL(DbMain, Q1)

         i = i + 1
         
         If lIdAreaNegExentoEntidad <> 0 Or lIdCCostoExentoEntidad <> 0 Then
            HayANegCcosto = True
         End If
         
         Exento = lMontoExento

      End If

      'Afecto (idem Exento)
      If lMontoAfecto <> 0 Then
         Q1 = QBase & i & ","                                     'Orden
         Q1 = Q1 & lIdCtaAfectoEntidad & ","                             'IdCuenta

         If TipoDocNC Then
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(lMontoAfecto) & ","                     'Haber
         Else
            Q1 = Q1 & Abs(lMontoAfecto) & ","                     'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                             'Glosa
         Q1 = Q1 & LIBCOMPRAS_AFECTO & ","                        'IdTipoValLib
         Q1 = Q1 & "0,"                                           'EsTotalDoc
         Q1 = Q1 & lIdAreaNegAfectoEntidad & ","                  'IdAreaNeg
         Q1 = Q1 & lIdCCostoAfectoEntidad & ","                   'IdCCosto
         Q1 = Q1 & "0,0,' '" & ")"                                'Tasa, EsRecuperable, CodSIIDTE

         Call ExecSQL(DbMain, Q1)

         i = i + 1
         
         Afecto = Afecto + lMontoAfecto
         
         If lIdAreaNegAfectoEntidad <> 0 Or lIdCCostoAfectoEntidad <> 0 Then
            HayANegCcosto = True
         End If

      End If

      'Neto Activo Fijo
      If lMontoNetoActFijo <> 0 Then
         Q1 = QBase & i & ","                                     'Orden
         Q1 = Q1 & lIdCtaAfectoEntidad & ","                             'IdCuenta

         If TipoDocNC Then
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(lMontoNetoActFijo) & ","  'Haber
         Else
            Q1 = Q1 & Abs(lMontoNetoActFijo) & ","  'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                             'Glosa
         Q1 = Q1 & LIBCOMPRAS_AFECTO & ","                        'IdTipoValLib
         Q1 = Q1 & "0,0,0,"                                       'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,0,' '" & ")"                                'Tasa, EsRecuperable, CodSIIDTE

         Call ExecSQL(DbMain, Q1)

         i = i + 1
         
         Afecto = Afecto + lMontoNetoActFijo
      End If

      'IVA Crédito Fiscal
      If lMontoIVARec <> 0 Or lMontoIVAUsoComun <> 0 Then
         Q1 = QBase & i & ","                                     'Orden
         
         If Abs(DateDiff("m", lFechaRec, lFechaEmision)) > 2 Then  'pasamos el IVA a IVA Irrecuperable N°2 de acuerdo a lo indicado por Thomson
            EsIVAIrrec = True
            Q1 = Q1 & lIdCtaIVAIrrec & ","                         'IdCuenta
         Else
            Q1 = Q1 & lIdCtaIVA & ","                              'IdCuenta
         End If

         If TipoDocNC Then
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(lMontoIVARec) + Abs(lMontoIVAUsoComun) & ","     'Haber
         Else
            Q1 = Q1 & Abs(lMontoIVARec) + Abs(lMontoIVAUsoComun) & ","     'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                             'Glosa

         If EsIVAIrrec Then
            Q1 = Q1 & LIBCOMPRAS_IVAIRREC2 & ","                  'IdTipoValLib
         Else
            Q1 = Q1 & LIBCOMPRAS_IVACREDFISC & ","                'IdTipoValLib
         End If
         
         Q1 = Q1 & "0,0,0,"                                       'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,0,' '" & ")"                                'Tasa, EsRecuperable, CodSIIDTE
         
         Call ExecSQL(DbMain, Q1)

         i = i + 1
         
         IVA = IVA + lMontoIVARec
      End If

      'IVA Activo Fijo
      If lMontoIVAActFijo <> 0 Then
         Q1 = QBase & i & ","                                     'Orden
         Q1 = Q1 & lIdCtaIVA & ","                                'IdCuenta

         If TipoDocNC Then
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(lMontoIVAActFijo) & ","                 'Haber
         Else
            Q1 = Q1 & Abs(lMontoIVAActFijo) & ","                 'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                             'Glosa

         'si es activo fijo, ponemos IVA Activo Fijo
         Q1 = Q1 & LIBCOMPRAS_IVAACTFIJO & ","                    'IdTipoValLib

         Q1 = Q1 & "0,0,0,"                                       'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,0,' '" & ")"                                'Tasa, EsRecuperable, CodSIIDTE
         
         Call ExecSQL(DbMain, Q1)

         i = i + 1
         
         IVA = IVA + lMontoIVAActFijo
         
         HayOtrosImp = True
         
      End If

      'IVA No recuperable
      If lMontoIVANoRec <> 0 Then
         Q1 = QBase & i & ","                                     'Orden
         Q1 = Q1 & lIdCtaIVAIrrec & ","                                'IdCuenta

         If TipoDocNC Then
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(lMontoIVANoRec) & ","                   'Haber
         Else
            Q1 = Q1 & Abs(lMontoIVANoRec) & ","                   'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                             'Glosa

         'Vemos el tipo de IVA Irrec
         TipoIVAIrrec = LIBCOMPRAS_IVAIRREC2                      'duda Ktaherine 13 ago 2019
         Select Case lCodIVANoRec
            Case 1
               TipoIVAIrrec = LIBCOMPRAS_IVAIRREC1
            Case 2
               TipoIVAIrrec = LIBCOMPRAS_IVAIRREC2
            Case 3
               TipoIVAIrrec = LIBCOMPRAS_IVAIRREC3
            Case 4
               TipoIVAIrrec = LIBCOMPRAS_IVAIRREC4
            Case 9
               TipoIVAIrrec = LIBCOMPRAS_IVAIRREC9
         End Select
         
         Q1 = Q1 & TipoIVAIrrec & ","                                'IdTipoValLib

         Q1 = Q1 & "0,0,0,"                                          'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,0,' '" & ")"                                   'Tasa, EsRecuperable, CodSIIDTE
         
         Call ExecSQL(DbMain, Q1)

         i = i + 1
         
         IVA = IVA + lMontoIVANoRec
         HayOtrosImp = True
      
      End If


      'Otros Impuestos Cigarrillos
      If lTabacosPuros <> 0 Then
         Q1 = QBase & i & ","                                        'Orden
         Q1 = Q1 & lIdCtaOtrosImp & ","                              'IdCuenta

         If TipoDocNC Then
            If lTabacosPuros * -1 > 0 Then    'como es NC, lo damos vuelta nuevamente, como lo hicimos al ingresarlo
               Q1 = Q1 & "0" & ","                                   'Debe
               Q1 = Q1 & Abs(lTabacosPuros) & ","                    'Haber
            Else
               Q1 = Q1 & Abs(lTabacosPuros) & ","                    'Debe
               Q1 = Q1 & "0" & ","                                   'Haber
            End If
         Else
            If lTabacosPuros > 0 Then
               Q1 = Q1 & Abs(lTabacosPuros) & ","                    'Debe
               Q1 = Q1 & "0" & ","                                   'Haber
            Else
               Q1 = Q1 & "0" & ","                                   'Debe
               Q1 = Q1 & Abs(lTabacosPuros) & ","                    'Haber
            End If
         End If

         Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
         Q1 = Q1 & LIBCOMPRAS_OTROSIMP & ","                         'IdTipoValLib
         Q1 = Q1 & "0,0,0,"                                          'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,0,' '" & ")"                                   'Tasa, EsRecuperable, CodSIIDTE

         Call ExecSQL(DbMain, Q1)

         i = i + 1
         
         OtroImp = OtroImp + lTabacosPuros
         HayOtrosImp = True
      End If
      
      If lTabacosCigarrillos <> 0 Then
         Q1 = QBase & i & ","                                        'Orden
         Q1 = Q1 & lIdCtaOtrosImp & ","                              'IdCuenta

         If TipoDocNC Then
            If lTabacosCigarrillos * -1 > 0 Then    'como es NC, lo damos vuelta nuevamente, como lo hicimos al ingresarlo
               Q1 = Q1 & "0" & ","                                   'Debe
               Q1 = Q1 & Abs(lTabacosCigarrillos) & ","              'Haber
            Else
               Q1 = Q1 & Abs(lTabacosCigarrillos) & ","              'Debe
               Q1 = Q1 & "0" & ","                                   'Haber
            End If
         Else
            If lTabacosCigarrillos > 0 Then
               Q1 = Q1 & Abs(lTabacosCigarrillos) & ","              'Debe
               Q1 = Q1 & "0" & ","                                   'Haber
            Else
               Q1 = Q1 & "0" & ","                                   'Debe
               Q1 = Q1 & Abs(lTabacosCigarrillos) & ","              'Haber
            End If
         End If

         Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
         Q1 = Q1 & LIBCOMPRAS_OTROSIMP & ","                         'IdTipoValLib
         Q1 = Q1 & "0,0,0,"                                          'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,0,' '" & ")"                                   'Tasa, EsRecuperable, CodSIIDTE

         Call ExecSQL(DbMain, Q1)

         i = i + 1
         
         OtroImp = OtroImp + lTabacosCigarrillos
         HayOtrosImp = True
      End If

      If lTabacosElaborados <> 0 Then
         Q1 = QBase & i & ","                                        'Orden
         Q1 = Q1 & lIdCtaOtrosImp & ","                              'IdCuenta

         If TipoDocNC Then
            If lTabacosElaborados * -1 > 0 Then    'como es NC, lo damos vuelta nuevamente, como lo hicimos al ingresarlo
               Q1 = Q1 & "0" & ","                                   'Debe
               Q1 = Q1 & Abs(lTabacosElaborados) & ","               'Haber
            Else
               Q1 = Q1 & Abs(lTabacosElaborados) & ","               'Debe
               Q1 = Q1 & "0" & ","                                   'Haber
            End If
         Else
            If lTabacosElaborados > 0 Then
               Q1 = Q1 & Abs(lTabacosElaborados) & ","               'Debe
               Q1 = Q1 & "0" & ","                                   'Haber
            Else
               Q1 = Q1 & "0" & ","                                   'Debe
               Q1 = Q1 & Abs(lTabacosElaborados) & ","               'Haber
            End If
         End If

         Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
         Q1 = Q1 & LIBCOMPRAS_OTROSIMP & ","                         'IdTipoValLib
         Q1 = Q1 & "0,0,0,"                                          'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,0,' '" & ")"                                   'Tasa, EsRecuperable, CodSIIDTE

         Call ExecSQL(DbMain, Q1)

         i = i + 1
      
         OtroImp = OtroImp + lTabacosElaborados
         HayOtrosImp = True
      End If
           
      'Otros Impuestos codificados Crédito Fiscal
      HaySoloIVARetenido = False
      OImpEsIVARetenido = False
      nIVARetenido = 0
      nIVANoRetenido = 0
      
      
      For r = 0 To UBound(lOtrosImp)
      
         If lOtrosImp(r).CodSIIDTE <> "" And lOtrosImp(r).CodSIIDTE <> "0" Then
         
            IdCuenta = 0
            ValOtroImp = lOtrosImp(r).Valor
            OImpEsIVARetenido = False
            
            IdTipoValLib = GetTipoTipoValLibFromCodSIIDTE(lTipoLib, lOtrosImp(r).CodSIIDTE)
            If IdTipoValLib > 0 Then
               IdCuenta = GetCtaImpAdic(lTipoLib, gTipoValLib(IdTipoValLib).TipoValLib, Tasa, EsRecuperable)
               
            End If
            
            OImpEsIVARetenido = EsIVARetenido(LIB_COMPRAS, lOtrosImp(r).CodSIIDTE)
            If OImpTipoIVARetenido > 0 Then
               nIVARetenido = nIVARetenido + 1
            Else
               nIVANoRetenido = nIVANoRetenido + 1
            End If
            
            'si el usuario no configuró sus propios impuestos adicionales, tomamos los por omisión
            
            If IdCuenta = 0 Then
               IdCuenta = lIdCtaOtrosImp
               Tasa = gTipoValLib(IdTipoValLib).Tasa
               EsRecuperable = gTipoValLib(IdTipoValLib).EsRecuperable
            End If
            
            If ValOtroImp <> 0 And IdTipoValLib > 0 And IdCuenta > 0 Then
               Q1 = QBase & i & ","                                        'Orden
               Q1 = Q1 & IdCuenta & ","                                 'IdCuenta
      
               'si es IVA retenido...
               'Si NO es IVA retenido nunca viene negativo
               If TipoDocNC Then
               
                  If OImpEsIVARetenido Then
                     If ValOtroImp * -1 > 0 Then    'como es NC, lo damos vuelta nuevamente, como lo hicimos al ingresarlo
                        Q1 = Q1 & "0" & ","                                'Debe
                        Q1 = Q1 & Abs(ValOtroImp) & ","                    'Haber
                     Else
                        Q1 = Q1 & Abs(ValOtroImp) & ","                    'Debe
                        Q1 = Q1 & "0" & ","                                'Haber
                     End If
                     
                  Else   'NO es IVA Retenido
                     If ValOtroImp * -1 > 0 Then    'como es NC, lo damos vuelta nuevamente, como lo hicimos al ingresarlo
                        Q1 = Q1 & Abs(ValOtroImp) & ","                    'Debe
                        Q1 = Q1 & "0" & ","                                'Haber
                     Else
                        Q1 = Q1 & "0" & ","                                'Debe             'Lo dimos vuelta por solicitud de Nicolas Catrin el 27 ago 2018
                        Q1 = Q1 & Abs(ValOtroImp) & ","                    'Haber
                     End If
                  End If
                  
               Else
                  If OImpEsIVARetenido Then
                     If ValOtroImp > 0 Then    'como es NC, lo damos vuelta nuevamente, como lo hicimos al ingresarlo
                        Q1 = Q1 & "0" & ","                                'Debe
                        Q1 = Q1 & Abs(ValOtroImp) & ","                    'Haber
                     Else
                        Q1 = Q1 & Abs(ValOtroImp) & ","                    'Debe
                        Q1 = Q1 & "0" & ","                                'Haber
                     End If
                  Else
                     If ValOtroImp > 0 Then
                        Q1 = Q1 & Abs(ValOtroImp) & ","                       'Debe
                        Q1 = Q1 & "0" & ","                                   'Haber
                     Else
                        Q1 = Q1 & "0" & ","                                   'Debe
                        Q1 = Q1 & Abs(ValOtroImp) & ","                       'Haber
                     End If
                  End If
               End If
      
               Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
               Q1 = Q1 & gTipoValLib(IdTipoValLib).TipoValLib & ","        'IdTipoValLib
               Q1 = Q1 & "0,0,0,"                                          'EsTotalDoc, IdCCosto, IdAreaNeg
               Q1 = Q1 & str(Tasa) & ","         'Tasa
               Q1 = Q1 & Abs(CInt(EsRecuperable)) & ","                       'EsRecuperable
               Q1 = Q1 & "'" & gTipoValLib(IdTipoValLib).CodSIIDTE & "' )"         'CodSIIDTE
               
               Call ExecSQL(DbMain, Q1)
      
               i = i + 1
            
               OtroImp = OtroImp + ValOtroImp
               HayOtrosImp = True
           
            End If
            
         End If
         
      Next r
      
      'Total
      If lMontoTotal <> 0 Then
         Q1 = QBase & i & ","                                        'Orden
         Q1 = Q1 & lIdCtaTotalEntidad & ","                          'IdCuenta

         If TipoDocNC Then
            Q1 = Q1 & Abs(lMontoTotal) & ","                         'Debe
            Q1 = Q1 & "0" & ","                                      'Haber
         Else
            Q1 = Q1 & "0" & ","                                      'Debe
            Q1 = Q1 & Abs(lMontoTotal) & ","                         'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
         Q1 = Q1 & LIBCOMPRAS_TOTAL & ","                            'IdTipoValLib
         Q1 = Q1 & "1,"                                              'EsTotalDoc
         Q1 = Q1 & lIdAreaNegTotalEntidad & ","                      'IdAreaNeg
         Q1 = Q1 & lIdCCostoTotalEntidad & ","                       'IdCCosto
         Q1 = Q1 & "0,0,' '" & ")"                                   'Tasa, EsRecuperable, CodSIIDTE

         Call ExecSQL(DbMain, Q1)
      
         If lIdAreaNegTotalEntidad <> 0 Or lIdCCostoTotalEntidad <> 0 Then
            HayANegCcosto = True
         End If

      End If
      
      
'-------------------------------------------------------------------------------------------------
   ElseIf lTipoLib = LIB_VENTAS Then
      
      'Exento

      If lMontoExento <> 0 Then
         Q1 = QBase & i & ","                                     'Orden
         Q1 = Q1 & lIdCtaExentoEntidad & ","                      'IdCuenta

         If TipoDocNC Then
            Q1 = Q1 & Abs(lMontoExento) & ","                     'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         Else
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(lMontoExento) & ","                     'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                             'Glosa
         Q1 = Q1 & LIBVENTAS_EXENTO & ","                         'IdTipoValLib
         Q1 = Q1 & "0,"                                           'EsTotalDoc
         Q1 = Q1 & lIdAreaNegExentoEntidad & ","                  'IdAreaNeg
         Q1 = Q1 & lIdCCostoExentoEntidad & ","                   'IdCCosto
         Q1 = Q1 & "0,0,' '" & ")"                                'Tasa, EsRecuperable, CodSIIDTE

         Call ExecSQL(DbMain, Q1)

         i = i + 1
         
         If lIdAreaNegExentoEntidad <> 0 Or lIdCCostoExentoEntidad <> 0 Then
            HayANegCcosto = True
         End If
         
         Exento = lMontoExento

      End If

      'Afecto (idem Exento)
      If lMontoAfecto <> 0 Then
         Q1 = QBase & i & ","                                     'Orden
         Q1 = Q1 & lIdCtaAfectoEntidad & ","                      'IdCuenta

         If TipoDocNC Then
            Q1 = Q1 & Abs(lMontoAfecto) & ","                     'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         Else
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(lMontoAfecto) & ","                     'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                             'Glosa
         Q1 = Q1 & LIBVENTAS_AFECTO & ","                        'IdTipoValLib
         Q1 = Q1 & "0,"                                           'EsTotalDoc
         Q1 = Q1 & lIdAreaNegAfectoEntidad & ","                  'IdAreaNeg
         Q1 = Q1 & lIdCCostoAfectoEntidad & ","                   'IdCCosto
         Q1 = Q1 & "0,0,' '" & ")"                                'Tasa, EsRecuperable, CodSIIDTE

         Call ExecSQL(DbMain, Q1)

         i = i + 1
         
         Afecto = Afecto + lMontoAfecto
         
         If lIdAreaNegAfectoEntidad <> 0 Or lIdCCostoAfectoEntidad <> 0 Then
            HayANegCcosto = True
         End If

      End If
      
      'IVA Débito Fiscal
      If lMontoIVA <> 0 Then
         Q1 = QBase & i & ","                                     'Orden
         
         Q1 = Q1 & lIdCtaIVA & ","                              'IdCuenta

         If TipoDocNC Then
            Q1 = Q1 & Abs(lMontoIVA) & ","                        'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         Else
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Abs(lMontoIVA) & ","                        'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                             'Glosa

         Q1 = Q1 & LIBVENTAS_IVADEBFISC & ","                  'IdTipoValLib
         Q1 = Q1 & "0,0,0,"                                       'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,0,' '" & ")"                                'Tasa, EsRecuperable, CodSIIDTE
         
         Call ExecSQL(DbMain, Q1)

         i = i + 1
         
         IVA = IVA + lMontoIVA
      End If
     
      'IVA Retenido Total
      '(usamos los campos que vienen al final de cada registro con los IVA retenidos codificados)
'      If lMontoIVARetTotal <> 0 Then
'         Q1 = QBase & i & ","                                        'Orden
'         Q1 = Q1 & lIdCtaOtrosImp & ","                              'IdCuenta
'
'         If TipoDocNC Then
'            If lMontoIVARetTotal * -1 > 0 Then    'como es NC, lo damos vuelta nuevamente, como lo hicimos al ingresarlo
'               Q1 = Q1 & Abs(lMontoIVARetTotal) & ","                                   'Debe
'               Q1 = Q1 & "0" & ","                'Haber
'            Else
'               Q1 = Q1 & "0" & ","                'Debe
'               Q1 = Q1 & Abs(lMontoIVARetTotal) & ","                                   'Haber
'            End If
'         Else
'            If lMontoIVARetTotal > 0 Then
'               Q1 = Q1 & Abs(lMontoIVARetTotal) & ","                'Debe
'               Q1 = Q1 & "0" & ","                                   'Haber
'            Else
'               Q1 = Q1 & "0" & ","                                   'Debe
'               Q1 = Q1 & Abs(lMontoIVARetTotal) & ","                'Haber
'            End If
'         End If
'
'         Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
'         Q1 = Q1 & LIBVENTAS_IVARETTOT & ","                         'IdTipoValLib
'         Q1 = Q1 & "0,0,0,"                                          'EsTotalDoc, IdCCosto, IdAreaNeg
'         Q1 = Q1 & "0,0,' '" & ")"                                   'Tasa, EsRecuperable, CodSIIDTE
'
'         Call ExecSQL(DbMain, Q1)
'
'         i = i + 1
'
'         OtroImp = OtroImp - lMontoIVARetTotal
'
'         HayOtrosImp = True
'      End If
     
      'IVA Retenido Parcial
      '(usamos los campos que vienen al final de cada registro con los IVA retenidos codificados)
'      If lMontoIVARetParcial <> 0 Then
'         Q1 = QBase & i & ","                                        'Orden
'         Q1 = Q1 & lIdCtaOtrosImp & ","                              'IdCuenta
'
'         If TipoDocNC Then
'            If lMontoIVARetParcial * -1 > 0 Then    'como es NC, lo damos vuelta nuevamente, como lo hicimos al ingresarlo
'               Q1 = Q1 & Abs(lMontoIVARetParcial) & ","                                  'Debe
'               Q1 = Q1 & "0" & ","                 'Haber
'            Else
'               Q1 = Q1 & "0" & ","                 'Debe
'               Q1 = Q1 & Abs(lMontoIVARetParcial) & ","                                   'Haber
'            End If
'         Else
'            If lMontoIVARetParcial > 0 Then
'               Q1 = Q1 & Abs(lMontoIVARetParcial) & ","                'Debe
'               Q1 = Q1 & "0" & ","                                   'Haber
'            Else
'               Q1 = Q1 & "0" & ","                                   'Debe
'               Q1 = Q1 & Abs(lMontoIVARetParcial) & ","                'Haber
'            End If
'         End If
'
'         Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
'         Q1 = Q1 & LIBVENTAS_IVARETPARC & ","                        'IdTipoValLib
'         Q1 = Q1 & "0,0,0,"                                          'EsTotalDoc, IdCCosto, IdAreaNeg
'         Q1 = Q1 & "0,0,' '" & ")"                                   'Tasa, EsRecuperable, CodSIIDTE
'
'         Call ExecSQL(DbMain, Q1)
'
'         OtroImp = OtroImp - lMontoIVARetParcial
'         i = i + 1
'
'         HayOtrosImp = True
'      End If
         
      'Crédito Empresa Constructora
      If lCredEmpContructora <> 0 Then
         Q1 = QBase & i & ","                                        'Orden
         Q1 = Q1 & lIdCtaOtrosImp & ","                         'IdCuenta

         If TipoDocNC Then
            Q1 = Q1 & Abs(lCredEmpContructora) & ","                'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         Else
            Q1 = Q1 & Abs(lCredEmpContructora) & ","                'Debe
            Q1 = Q1 & "0" & ","                                   'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
         Q1 = Q1 & LIBVENTAS_REBAJA65 & ","                          'IdTipoValLib
         Q1 = Q1 & "0,0,0,"                                          'EsTotalDoc, IdCCosto, IdAreaNeg
         Q1 = Q1 & "0,0,' '" & ")"                                   'Tasa, EsRecuperable, CodSIIDTE

         Call ExecSQL(DbMain, Q1)

         OtroImp = OtroImp - lCredEmpContructora
         i = i + 1
         
         HayOtrosImp = True
      End If
         
      'Otros Impuestos
      '(usamos los campos que vienen al final de cada registro con los IVA retenidos codificados, y eso está más abajo)
'      If lMontoOtroImp <> 0 Then
'
'         IdTipoValLib = GetTipoTipoValLibFromCodSIIDTE(lTipoLib, lCodOtroImp)
'         IdCuenta = 0
'         If IdTipoValLib > 0 Then
'            IdCuenta = GetIdCuentaImpAdic(lTipoLib, gTipoValLib(IdTipoValLib).TipoValLib)
'         End If
'
'         IdCuenta = IIf(IdCuenta = 0, lIdCtaOtrosImp, IdCuenta)
'         If IdCuenta > 0 And IdTipoValLib > 0 Then
'            Q1 = QBase & i & ","                                        'Orden
'   '         Q1 = Q1 & lIdCtaOtrosImp & ","                              'IdCuenta
'            Q1 = Q1 & IdCuenta & ","                                    'IdCuenta
'
'            If TipoDocNC Then
'               If lMontoOtroImp * -1 > 0 Then    'como es NC, lo damos vuelta nuevamente, como lo hicimos al ingresarlo
'                  Q1 = Q1 & Abs(lMontoOtroImp) & ","                    'Debe
'                  Q1 = Q1 & "0" & ","                                   'Haber
'               Else
'                  Q1 = Q1 & "0" & ","                                   'Debe
'                  Q1 = Q1 & Abs(lMontoOtroImp) & ","                    'Haber
'               End If
'            Else
'               If lMontoOtroImp > 0 Then
'                  Q1 = Q1 & "0" & ","                                   'Debe
'                  Q1 = Q1 & Abs(lMontoOtroImp) & ","                    'Haber
'               Else
'                  Q1 = Q1 & Abs(lMontoOtroImp) & ","                    'Debe
'                  Q1 = Q1 & "0" & ","                                   'Haber
'               End If
'            End If
'
'            Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
'   '         Q1 = Q1 & LIBVENTAS_OTROSIMP & ","                         'IdTipoValLib
'            Q1 = Q1 & gTipoValLib(IdTipoValLib).TipoValLib & ","        'IdTipoValLib
'            Q1 = Q1 & "0,0,0,"                                          'EsTotalDoc, IdCCosto, IdAreaNeg
'            Q1 = Q1 & Str(lTasaOtroImp) & ",0, '" & lCodOtroImp & "')"  'Tasa, EsRecuperable, CodSIIDTE
'
'            Call ExecSQL(DbMain, Q1)
'
'            i = i + 1
'
'            OtroImp = OtroImp + lMontoOtroImp
'            HayOtrosImp = True
'
'         End If
'
'      End If
      
      'Otros Impuestos codificados
      HaySoloIVARetenido = False
      OImpTipoIVARetenido = 0
      nIVARetenido = 0
      nIVANoRetenido = 0
      
      
      For r = 0 To UBound(lOtrosImp)
      
         If lOtrosImp(r).CodSIIDTE <> "" And lOtrosImp(r).CodSIIDTE <> "0" Then
         
            IdCuenta = 0
            ValOtroImp = lOtrosImp(r).Valor
            
            IdTipoValLib = GetTipoTipoValLibFromCodSIIDTE(lTipoLib, lOtrosImp(r).CodSIIDTE)
            If IdTipoValLib > 0 Then
               IdCuenta = GetCtaImpAdic(lTipoLib, gTipoValLib(IdTipoValLib).TipoValLib, Tasa, EsRecuperable)
               OImpTipoIVARetenido = gTipoValLib(IdTipoValLib).TipoIVARetenido
            End If
                     
            'si el usuario no configuró sus propios impuestos adicionales, tomamos los por omisión
            
            If IdCuenta = 0 And IdTipoValLib > 0 Then
               IdCuenta = lIdCtaOtrosImp
               Tasa = gTipoValLib(IdTipoValLib).Tasa
               EsRecuperable = gTipoValLib(IdTipoValLib).EsRecuperable
            End If
                      
            If OImpTipoIVARetenido > 0 Then
               nIVARetenido = nIVARetenido + 1
            Else
               nIVANoRetenido = nIVANoRetenido + 1
            End If
            
            If ValOtroImp <> 0 And IdTipoValLib > 0 And IdCuenta > 0 Then
               Q1 = QBase & i & ","                                        'Orden
               Q1 = Q1 & IdCuenta & ","                                 'IdCuenta
      
      
               'si es IVA retenido...
               'Si NO es IVA retenido nunca viene negativo
               If TipoDocNC Then
               
                  If OImpTipoIVARetenido > 0 Then
                     Q1 = Q1 & "0" & ","                    'Debe
                     Q1 = Q1 & Abs(ValOtroImp) & ","         'Haber
                     
                  Else   'NO es IVA Retenido
                     Q1 = Q1 & Abs(ValOtroImp) & ","        'Debe
                     Q1 = Q1 & "0" & ","                    'Haber             'Lo dimos vuelta por solicitud de Nicolas Catrin el 27 ago 2018
                  End If
                  
               Else
                  If OImpTipoIVARetenido > 0 Then
                     Q1 = Q1 & Abs(ValOtroImp) & ","         'Debe
                     Q1 = Q1 & "0" & ","                    'Haber
                  Else
                     Q1 = Q1 & "0" & ","                       'Debe
                     Q1 = Q1 & Abs(ValOtroImp) & ","           'Haber
                  End If
               End If
      
               Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
               Q1 = Q1 & gTipoValLib(IdTipoValLib).TipoValLib & ","        'IdTipoValLib
               Q1 = Q1 & "0,0,0,"                                          'EsTotalDoc, IdCCosto, IdAreaNeg
               Q1 = Q1 & str(gTipoValLib(IdTipoValLib).Tasa) & ","         'Tasa
               Q1 = Q1 & Abs(CInt(gTipoValLib(IdTipoValLib).EsRecuperable)) & ","      'EsRecuperable
               Q1 = Q1 & "'" & gTipoValLib(IdTipoValLib).CodSIIDTE & "' )"         'CodSIIDTE
               
               Call ExecSQL(DbMain, Q1)
      
               i = i + 1
            
               OtroImp = OtroImp + ValOtroImp
               HayOtrosImp = True
           
            End If
            
         End If
         
      Next r
      
      
      'Total
      If lMontoTotal <> 0 Then
         Q1 = QBase & i & ","                                        'Orden
         Q1 = Q1 & lIdCtaTotalEntidad & ","                          'IdCuenta

         If TipoDocNC Then
            Q1 = Q1 & "0" & ","                                      'Debe
            Q1 = Q1 & Abs(lMontoTotal) & ","                         'Haber
         Else
            Q1 = Q1 & Abs(lMontoTotal) & ","                         'Debe
            Q1 = Q1 & "0" & ","                                      'Haber
         End If

         Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
         Q1 = Q1 & LIBVENTAS_TOTAL & ","                            'IdTipoValLib
         Q1 = Q1 & "1,"                                              'EsTotalDoc
         Q1 = Q1 & lIdAreaNegTotalEntidad & ","                      'IdAreaNeg
         Q1 = Q1 & lIdCCostoTotalEntidad & ","                       'IdCCosto
         Q1 = Q1 & "0,0,' '" & ")"                                   'Tasa, EsRecuperable, CodSIIDTE

         Call ExecSQL(DbMain, Q1)
      
         If lIdAreaNegTotalEntidad <> 0 Or lIdCCostoTotalEntidad <> 0 Then
            HayANegCcosto = True
         End If

      End If
     
   End If
   
   'Ahora actualizamos los campos del encabezado del documento
   
   'limpiamos los campos
   Q1 = "UPDATE Documento SET Exento = " & Exento & ", Afecto=" & Afecto & ", IVA=" & IVA & ", OtroImp=" & OtroImp & ", OtrosVal=0 WHERE IdDoc=" & lIdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)
   
   'OJO: En el campo IVA va IVA Crédito/Débito + IVA Act. Fijo + IVA Irrecuperable en todos sus tipos
   
   RetParcial = 0
   IVAIrrec = 0
   
   If lTipoLib = LIB_COMPRAS Then
   
      If lMontoIVANoRec > 0 Then
                  
         If lMontoIVANoRec = lMontoIVARec Then
            IVAIrrec = IVAIRREC_TOTAL
         ElseIf lMontoIVANoRec < lMontoIVARec Then
            IVAIrrec = IVAIRREC_PARCIAL
         ElseIf lMontoIVARec = 0 Then
            IVAIrrec = IVAIRREC_TOTAL
         End If
         
      Else
         IVAIrrec = IVAIRREC_CERO
         
      End If

   ElseIf lTipoLib = LIB_VENTAS Then
   
      If lMontoIVARetParcial <> 0 Then
         RetParcial = 1
      End If
   End If
   
   
   'actualizamos los campos, incluyendo los que limpiamos
   Q1 = "UPDATE Documento SET MovEdited=" & CInt(HayOtrosImp Or HayANegCcosto Or lDefCtasProveedor) & ", FacCompraRetParcial = " & RetParcial & ", IVAIrrecuperable = " & IVAIrrec & ", ValIVAIrrec = " & lMontoIVANoRec & ", CodSIIDTEIVAIrrec = " & lCodIVANoRec & ", IVAInmueble = 0, IVAActFijo = " & lMontoIVAActFijo
   If Exento <> 0 Then
      Q1 = Q1 & ", IdCuentaExento = " & lIdCtaExentoEntidad
   End If
   If Afecto <> 0 Then
      Q1 = Q1 & ", IdCuentaAfecto = " & lIdCtaAfectoEntidad
   End If
   Q1 = Q1 & ", IdCuentaTotal = " & lIdCtaTotalEntidad
   
   Q1 = Q1 & " WHERE IdDoc=" & lIdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)

   'actualizamos OtrosImp calculándolo exactamente de la misma manera como se calcula en el libro de ComprasVentas
   If lTipoLib = LIB_COMPRAS Or lTipoLib = LIB_VENTAS Then
      
      Q1 = "SELECT Total, Afecto, Exento, IVA FROM Documento WHERE IdDoc=" & lIdDoc
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
         'exactamente igual como se calcula en el Libro de ComprasVentas
         ValOtros = Abs(vFld(Rs("Total")) - (vFld(Rs("Exento")) + vFld(Rs("Afecto")) + vFld(Rs("IVA"))))
         Q1 = "UPDATE Documento SET OtroImp = " & ValOtros
         If ValOtros <> 0 Then
            Q1 = Q1 & ", IdCuentaOtroImp = " & lIdCtaOtrosImp
         Else
            Q1 = Q1 & ", IdCuentaOtroImp = 0 "
         End If
         
         Q1 = Q1 & "  WHERE IdDoc=" & lIdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      End If
      Call CloseRs(Rs)
   End If

   'si es NCC y sólo vienen IVARetenido, hay que cambiar el tipo de doc a NCF
   If lTipoLib = LIB_COMPRAS And TipoDocNC Then
      If nIVARetenido > 0 And nIVANoRetenido = 0 Then
         TipoDocNCF = FindTipoDoc(LIB_COMPRAS, "NCF")
         
         Q1 = "UPDATE Documento SET "
         Q1 = Q1 & " TipoDoc = " & TipoDocNCF
         Q1 = Q1 & " WHERE IdDoc=" & lIdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      End If
   End If
   
End Sub

