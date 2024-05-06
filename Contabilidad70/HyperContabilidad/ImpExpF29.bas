Attribute VB_Name = "ImpExpF29"
Option Explicit

Public Const LAU_TRET_RENTASCAP = 30
Public Const LAU_TRET_HONORARIOS = 1
Public Const LAU_TRET_PARTDIR10 = 2
Public Const LAU_TRET_PARTDIR20 = 4
Public Const LAU_TRET_RETMINEROS = 35

'cuentas default
Private lIdCtaAfecto As Long
Private lIdCtaExento As Long
Private lIdCtaTotal As Long
Private lIdCtaActFijoAfecto As Long
Private lIdCtaActFijoExento As Long

Private lIdCtaIVA As Long
Private lIdCtaOtrosImp As Long

Private lIdCtaHonSinRet As Long
Private lIdCtaBruto As Long

Private lIdCtaImpRet As Long
Private lIdCtaNetoHon As Long
Private lIdCtaNetoDieta As Long

Private lIdCtaIVAIrrec As Long

'Esta función genera la estructura de datos para llenar el reporte Resumen de Libros y
'para realizar la exportación a IVA estándar, tomando los movimientos del documento,
'considerando que ahora se ingresa el detalle.
'
'La idea es:
'
' - distribuir correctamente el activo fijo de las facturas que tienen tanto activo fijo como otros no activos fijos
'     - para esto se requiere definir un TipoValor IVA Cred. Fisc. Activo Fijo con el fin de distribuir
'       no sólo el afecto y exento, sino también el IVA
'
'
Public Function GenResLibros(ByVal Where As String, ResLib() As ResLib_t, ResOImp() As ResOImp_t, ExpIVA As Boolean) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim CurReg As String
   Dim NewReg As String
   Dim Wh As String
   Dim Dif As Double
   Dim AuxDif As Double
   Dim EsIgual As Boolean
   Dim IifActFijo As String
   Dim CondTipoVal As String
   Dim IifTipoReten As String
   Dim IifNCVExenta As String
   Dim TipoDocFVE As Integer
   Dim TipoDoc As Integer
   Dim InicioAno As Long
   Dim IifSuper As String
   Dim InnerJoinSuper As String
   Dim IifSuperOrder As String, IifSuperAs As String, IifSuperGroup As String
   Dim TipoIVARetenido As Integer
   Dim Idx As Integer
   
   GenResLibros = False
      
   Wh = "Documento.TipoLib IN(" & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ") AND Documento.Estado <> " & ED_ANULADO
   If Where <> "" Then
      Wh = Wh & " AND " & Where
   End If
   
   TipoDocFVE = FindTipoDoc(LIB_VENTAS, "FVE")
   
   CondTipoVal = "(Documento.TipoLib = " & LIB_COMPRAS & " And Not(IdTipoValLib IS NULL) And(IdTipoValLib = " & LIBCOMPRAS_IVAACTFIJO & "))"
   IifActFijo = " iif((Cuentas.Atrib" & ATRIB_ACTIVOFIJO & " IS NULL OR Cuentas.Atrib" & ATRIB_ACTIVOFIJO & " = 0 ) AND NOT " & CondTipoVal & ", 0, 1 )"
      
   IifTipoReten = " iif(Documento.TipoRetencion IS NULL, 0, Documento.TipoRetencion) "
   IifNCVExenta = " IIf(TipoDocs.Diminutivo = 'NCV' AND Documento.IVA = 0, 1, 0) "    'marcamos las notas de crédito de venta que sólo tienen parte exenta
   
   InicioAno = DateSerial(gEmpresa.Ano, 1, 1)
   
   If InicioAno >= gFechaInicioSupermercados Then
      IifSuper = " IIf( TipoDocs.Diminutivo = 'FAC' AND Entidades.EsSupermercado <> 0, 1, 0) "
      IifSuperAs = IifSuper & " As EsSupermercado,"
      IifSuperOrder = ", " & IifSuper
      IifSuperGroup = IifSuper & ", "
      InnerJoinSuper = " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
      InnerJoinSuper = InnerJoinSuper & " AND Entidades.IdEmpresa = Documento.IdEmpresa ) "
   Else
     InnerJoinSuper = ")"
   End If
    
   Q1 = "SELECT " & SqlMonthLng("FEmision") & " As Mes, Documento.TipoLib, Documento.TipoDoc, DTE, Documento.Giro, Documento.FacCompraRetParcial, Documento.IVAIrrecuperable, MovDocumento.IdTipoValLib, "
   Q1 = Q1 & IifActFijo & " As ActFijo, "
   Q1 = Q1 & IifTipoReten & "  As TipoReten, "
   Q1 = Q1 & IifNCVExenta & " As NCV_Exenta, "
   Q1 = Q1 & IifSuperAs
   Q1 = Q1 & " CodF29Count, CodF29Neto, CodF29IVA, CodF29IVADTE, CodF29AFCount, CodF29AFIVA, CodF29ExCount, CodF29Exento, CodF29RetHon, CodF29RetDieta, "
   Q1 = Q1 & " CodF29IVARet3ro, CodF29CountNoGiro, CodF29NetoNoGiro, CodF29IVANoGiro, CodF29ExCountNoGiro, CodF29ExentoNoGiro, "
   Q1 = Q1 & " CodF29CountRetParcial, CodF29NetoRetParcial, CodF29DifIVARetParcial, CodF29CountDTE, CodF29NetoDTE, CodF29IVaIrrecDTE, "
   Q1 = Q1 & " CodF29CountIVAIrrec, CodF29NetoIVAIrrec, CodF29CountSuper, CodF29IVASuper, "
   Q1 = Q1 & " TipoDocs.EsRebaja, "
   Q1 = Q1 & " Sum(MovDocumento.Debe) as SumDebe, Sum(MovDocumento.Haber) as SumHaber "
   Q1 = Q1 & " FROM (((( Documento "
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib= TipoDocs.TipoLib AND Documento.TipoDoc= TipoDocs.TipoDoc) "
   Q1 = Q1 & " LEFT JOIN MovDocumento ON Documento.IdDoc=MovDocumento.IdDoc"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas  ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento") & " )"
   Q1 = Q1 & InnerJoinSuper
   Q1 = Q1 & " WHERE " & Wh & " AND (MovDocumento.EsTotalDoc = 0 OR MovDocumento.EsTotalDoc IS NULL)"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY " & SqlMonthLng("FEmision") & ", Documento.TipoLib, Documento.TipoDoc, DTE, Documento.Giro, Documento.FacCompraRetParcial, Documento.IVAIrrecuperable, MovDocumento.IdTipoValLib, "
   Q1 = Q1 & IifActFijo & ", "
   Q1 = Q1 & IifTipoReten & ", "
   Q1 = Q1 & IifNCVExenta & ", "
   Q1 = Q1 & IifSuperGroup
   Q1 = Q1 & " CodF29Count, CodF29Neto, CodF29IVA, CodF29IVADTE, CodF29AFCount, CodF29AFIVA, CodF29ExCount, CodF29Exento, CodF29RetHon, CodF29RetDieta, CodF29IVARet3ro, "
   Q1 = Q1 & " CodF29CountNoGiro, CodF29NetoNoGiro, CodF29IVANoGiro, CodF29ExCountNoGiro, CodF29ExentoNoGiro, "
   Q1 = Q1 & " CodF29CountRetParcial, CodF29NetoRetParcial, CodF29DifIVARetParcial, CodF29CountDTE, CodF29NetoDTE, CodF29IVaIrrecDTE, CodF29CountIVAIrrec, CodF29NetoIVAIrrec, "
   Q1 = Q1 & " CodF29CountSuper, CodF29IVASuper, EsRebaja "
   Q1 = Q1 & " ORDER BY " & SqlMonthLng("FEmision") & ", Documento.TipoLib, Documento.TipoDoc " & IifSuperOrder & ","
   'Q1 = Q1 & IifActFijo & " , iif(DTE <> 0, 1, 0) , " & IifTipoReten & " , iif(Giro <> 0, 1, 0) , iif(FacCompraRetParcial <> 0, 1, 0), IVAIrrecuperable"     'FCA 10 abr 2015
   Q1 = Q1 & IifActFijo & " , " & IifTipoReten & " , iif(Documento.Giro <> 0, 1, 0) , iif(FacCompraRetParcial <> 0, 1, 0), IVAIrrecuperable, iif(DTE <> 0, 1, 0) "

   Set Rs = OpenRs(DbMain, Q1)
      
   i = -1
   CurReg = ""
   
   ReDim ResLib(30)
   
   Do While Rs.EOF = False
   
      TipoDoc = vFld(Rs("TipoDoc"))
      
'      If vFld(Rs("NCV_Exenta")) <> 0 And ExpIVA Then     'si es una nota de crédito que sólo tiene monto exento se asimila a una Factura de Venta Exenta, desde el punto de vista del Form 29, ya que se traspasa a los mismos códigos
'         TipoDoc = TipoDocFVE                            '(solicitado por Victor el 16 abril 2013)
'      End If
      
      NewReg = vFld(Rs("Mes")) & "-" & vFld(Rs("TipoLib")) & "-" & TipoDoc
      
      If vFld(Rs("ActFijo")) = 0 Then
         NewReg = NewReg & "-0"
         
      Else
         NewReg = NewReg & "-1"
               
      End If
         
      If vFld(Rs("Giro")) = 0 Then
         NewReg = NewReg & "-0"
         
      Else
         NewReg = NewReg & "-1"
               
      End If
         
      If vFld(Rs("FacCompraRetParcial")) = 0 Then
         NewReg = NewReg & "-0"
         
      Else
         NewReg = NewReg & "-1"
               
      End If
      
      NewReg = NewReg & "-" & vFld(Rs("IVAIrrecuperable"))
               
      NewReg = NewReg & "-" & vFld(Rs("TipoReten"))
           
      If IifSuper <> "" Then
         NewReg = NewReg & "-" & vFld(Rs("EsSupermercado"))
      End If
                     
      If NewReg <> CurReg Then
                  
         i = i + 1
         
         If i > UBound(ResLib) Then
            ReDim Preserve ResLib(i + 10)
         End If
            
         ResLib(i).Mes = vFld(Rs("Mes"))
         ResLib(i).TipoLib = vFld(Rs("TipoLib"))
         ResLib(i).TipoDoc = TipoDoc
         ResLib(i).ActFijo = IIf(vFld(Rs("ActFijo")) = 0, 0, 1)
         ResLib(i).Giro = IIf(vFld(Rs("Giro")) = 0, 0, 1)
         ResLib(i).FacCompraRetParcial = IIf(vFld(Rs("FacCompraRetParcial")) = 0, 0, 1)
         ResLib(i).IVAIrrec = vFld(Rs("IVAIrrecuperable"))
         ResLib(i).TipoReten = vFld(Rs("TipoReten"))
         If IifSuper = "" Then
            ResLib(i).EsSupermercado = 0
         Else
            ResLib(i).EsSupermercado = IIf(vFld(Rs("EsSupermercado")) = 0, 0, 1)
         End If
         
         ResLib(i).CodF29Count = vFld(Rs("CodF29Count"))
         ResLib(i).CodF29Neto = vFld(Rs("CodF29Neto"))
         ResLib(i).CodF29IVA = vFld(Rs("CodF29IVA"))
         ResLib(i).CodF29CountNoGiro = vFld(Rs("CodF29CountNoGiro"))
         ResLib(i).CodF29NetoNoGiro = vFld(Rs("CodF29NetoNoGiro"))
         ResLib(i).CodF29IVANoGiro = vFld(Rs("CodF29IVANoGiro"))
         ResLib(i).CodF29IVADTE = vFld(Rs("CodF29IVADTE"))
         ResLib(i).CodF29CountDTE = vFld(Rs("CodF29CountDTE"))
         ResLib(i).CodF29NetoDTE = vFld(Rs("CodF29NetoDTE"))
         ResLib(i).CodF29IVAIrrecDTE = vFld(Rs("CodF29IVaIrrecDTE"))
         ResLib(i).CodF29AFCount = vFld(Rs("CodF29AFCount"))
         ResLib(i).CodF29AFIVA = vFld(Rs("CodF29AFIVA"))
         ResLib(i).CodF29ExCount = vFld(Rs("CodF29ExCount"))
         ResLib(i).CodF29Exento = vFld(Rs("CodF29Exento"))
         ResLib(i).CodF29RetHon = vFld(Rs("CodF29RetHon"))
         ResLib(i).CodF29RetDieta = vFld(Rs("CodF29RetDieta"))
         ResLib(i).CodF29IVARet3ro = vFld(Rs("CodF29IVARet3ro"))
         ResLib(i).CodF29ExCountNoGiro = vFld(Rs("CodF29ExCountNoGiro"))
         ResLib(i).CodF29ExentoNoGiro = vFld(Rs("CodF29ExentoNoGiro"))
         ResLib(i).CodF29CountRetParcial = vFld(Rs("CodF29CountRetParcial"))
         ResLib(i).CodF29NetoRetParcial = vFld(Rs("CodF29NetoRetParcial"))
         ResLib(i).CodF29DifIVARetParcial = vFld(Rs("CodF29DifIVARetParcial"))
         ResLib(i).CodF29CountIVAIrrec = vFld(Rs("CodF29CountIVAIrrec"))
         ResLib(i).CodF29NetoIVAIrrec = vFld(Rs("CodF29NetoIVAIrrec"))
         ResLib(i).CodF29CountSuper = vFld(Rs("CodF29CountSuper"))
         ResLib(i).CodF29IVASuper = vFld(Rs("CodF29IVASuper"))
                  
         CurReg = NewReg
         
      End If
            
      Dif = Abs(vFld(Rs("SumDebe")) - vFld(Rs("SumHaber")))
      
      Select Case ResLib(i).TipoLib
      
         Case LIB_VENTAS
         
            Select Case vFld(Rs("IdTipoValLib"))
               
               Case LIBVENTAS_AFECTO
                  
                  If vFld(Rs("Giro")) <> 0 Then
                     
                     If vFld(Rs("FacCompraRetParcial")) = 0 Then     'las facturas de compra son siempre del giro
                        ResLib(i).Afecto = ResLib(i).Afecto + Dif
                     Else
                        ResLib(i).NetoRetParcial = ResLib(i).NetoRetParcial + Dif
                     End If
                  
                  Else
                     ResLib(i).AfectoNoGiro = ResLib(i).AfectoNoGiro + Dif
                  
                  End If
                    
               Case LIBVENTAS_EXENTO
               
'                  If vFld(Rs("NCV_Exenta")) <> 0 And ExpIVA Then
'
'                     'si es una nota de crédito que sólo tiene monto exento se asimila a una Factura de Venta Exenta, desde el punto de vista del Form 29, ya que se traspasa a los mismos códigos, pero debe restar
'                     If vFld(Rs("Giro")) <> 0 Then
'                        ResLib(i).Exento = ResLib(i).Exento - Dif
'                     Else
'                        ResLib(i).ExentoNoGiro = ResLib(i).ExentoNoGiro - Dif
'                     End If
'
'                  Else
                     If vFld(Rs("Giro")) <> 0 Then
                        ResLib(i).Exento = ResLib(i).Exento + Dif
                     Else
                        ResLib(i).ExentoNoGiro = ResLib(i).ExentoNoGiro + Dif
                     End If
                     
'                  End If
                 
               Case LIBVENTAS_IVADEBFISC
                  If vFld(Rs("Giro")) <> 0 Then       'aquí no se hace diferencia con RetParcial como en el afecto, porque el IVA ret Parcial se calcula a partir del Afecto Ret Parcial
                     ResLib(i).IVA = ResLib(i).IVA + Dif
                  Else
                     ResLib(i).IVANoGiro = ResLib(i).IVANoGiro + IIf(vFld(Rs("EsRebaja")) <> 0, Dif * -1, Dif)
                  End If
                  
                  If vFld(Rs("DTE")) <> 0 Then
                     ResLib(i).IVADTE = ResLib(i).IVADTE + IIf(vFld(Rs("EsRebaja")) <> 0, Dif * -1, Dif)
                  End If
                              
                  If vFld(Rs("FacCompraRetParcial")) <> 0 Then
                     ResLib(i).DifIVARetParcial = ResLib(i).DifIVARetParcial + Dif    'vamos sumando el IVA para luego restar el ret parcial en la función ExportF29
                  End If
                  
               Case LIBVENTAS_TOTAL
               
               Case Else
                  AuxDif = vFld(Rs("SumHaber")) - vFld(Rs("SumDebe"))
                  'cambiar IIf por AuxDif y eliminar cambio de signo en LoadAll de FrmResLibAux ya que OtroImp viene listo, igual que IVADTE y OtroImpDTE
                  'ResLib(i).OtroImp = ResLib(i).OtroImp + IIf(vFld(Rs("EsRebaja")) <> 0, AuxDif * -1, AuxDif)
                  ResLib(i).OtroImp = ResLib(i).OtroImp + AuxDif
                  If vFld(Rs("DTE")) <> 0 Then
                     ResLib(i).OImpDTE = ResLib(i).OImpDTE + AuxDif
                  End If
                  
                  Idx = FindTipoValLib(ResLib(i).TipoLib, vFld(Rs("IdTipoValLib")))
                  TipoIVARetenido = gTipoValLib(Idx).TipoIVARetenido

                  If TipoIVARetenido = IVARET_PARCIAL Or TipoIVARetenido = IVARET_TOTAL Then
                     ResLib(i).IVARetenido = ResLib(i).IVARetenido + AuxDif
                  End If
                  
            End Select
         
         Case LIB_COMPRAS
         
             Select Case vFld(Rs("IdTipoValLib"))
             
               Case LIBCOMPRAS_AFECTO
                  ResLib(i).Afecto = ResLib(i).Afecto + Dif
                                                      
                  If vFld(Rs("DTE")) <> 0 Then
                     ResLib(i).NetoDTE = ResLib(i).NetoDTE + IIf(vFld(Rs("EsRebaja")) <> 0, Dif * -1, Dif)
                  End If
                                    
               Case LIBCOMPRAS_EXENTO
                  ResLib(i).Exento = ResLib(i).Exento + Dif
                  
                  If vFld(Rs("DTE")) <> 0 Then
                     ResLib(i).NetoDTE = ResLib(i).NetoDTE + IIf(vFld(Rs("EsRebaja")) <> 0, Dif * -1, Dif)
                  End If
                  
               Case LIBCOMPRAS_IVACREDFISC, LIBCOMPRAS_IVAACTFIJO
               
                  If vFld(Rs("IVAIrrecuperable")) <= IVAIRREC_PARCIAL Then   'no tiene IVA Irrec o es parcial
                     ResLib(i).IVA = ResLib(i).IVA + Dif
                  End If
                  
                  If vFld(Rs("DTE")) <> 0 Then
                     ResLib(i).IVADTE = ResLib(i).IVADTE + IIf(vFld(Rs("EsRebaja")) <> 0, Dif * -1, Dif)
                  End If
                                                
               Case LIBCOMPRAS_TOTAL
               
               Case Else
               
                  AuxDif = vFld(Rs("SumDebe")) - vFld(Rs("SumHaber"))
                  'cambiar IIf por AuxDif y eliminar cambio de signo en LoadAll de FrmResLibAux ya que OtroImp viene listo, igual que IVADTE y OtroImpDTE
                  'ResLib(i).OtroImp = ResLib(i).OtroImp + IIf(vFld(Rs("EsRebaja")) <> 0, AuxDif * -1, AuxDif)
                  
                  ResLib(i).OtroImp = ResLib(i).OtroImp + AuxDif
                  
                  If vFld(Rs("DTE")) <> 0 Then
                     ResLib(i).OImpDTE = ResLib(i).OImpDTE + AuxDif
                  End If
                  
                  If vFld(Rs("IVAIrrecuperable")) <> 0 And (vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC1 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC2 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC3 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC4 Or vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAIRREC9) Then
                     ResLib(i).NetoIVAIrrec = ResLib(i).NetoIVAIrrec + Round(AuxDif / gIVA)
                     
                     If vFld(Rs("DTE")) <> 0 Then
                        ResLib(i).IVAIrrecDTE = ResLib(i).IVAIrrecDTE + AuxDif
                     End If
                     
                  End If
                  
                  Idx = FindTipoValLib(ResLib(i).TipoLib, vFld(Rs("IdTipoValLib")))
                  TipoIVARetenido = gTipoValLib(Idx).TipoIVARetenido

                  If TipoIVARetenido = IVARET_PARCIAL Or TipoIVARetenido = IVARET_TOTAL Then
                     ResLib(i).IVARetenido = ResLib(i).IVARetenido + AuxDif
                  End If
                                    
            End Select
        
         Case LIB_RETEN
         
             Select Case vFld(Rs("IdTipoValLib"))
               Case LIBRETEN_BRUTO
                  ResLib(i).Afecto = ResLib(i).Afecto + Dif
                  
               Case LIBRETEN_HONORSINRET
                  ResLib(i).Exento = ResLib(i).Exento + Dif
                  
               Case LIBRETEN_IMPUESTO, LIBRETEN_RET3PORC
                  ResLib(i).OtroImp = ResLib(i).OtroImp + Dif
                  
               Case LIBRETEN_NETO
                                 
            End Select
       
      End Select
      
      If vFld(Rs("DTE")) <> 0 Then
         'ResLib(i).CountDTE = ResLib(i).CountDTE + vFld(Rs("N"))
      End If
         
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   If i >= 0 Then
      ReDim Preserve ResLib(i)
   Else
      ReDim ResLib(0)
      Exit Function
   End If

   'ahora contamos la cantidad de documentos
   Q1 = "SELECT DISTINCT " & SqlMonthLng("FEmision") & " As Mes, Documento.TipoLib, Documento.TipoDoc, DTE, Documento.Giro, Documento.FacCompraRetParcial, Documento.IVAIrrecuperable, Documento.Exento, Documento.Afecto, "
   Q1 = Q1 & IifActFijo & " As ActFijo, "
   'Q1 = Q1 & IifIVAIrrec & " As IVAIrrec, "
   Q1 = Q1 & IifNCVExenta & " As NCV_Exenta, "
   Q1 = Q1 & IifSuperAs
   Q1 = Q1 & " iif(Documento.TipoRetencion IS NULL, 0, Documento.TipoRetencion) As TipoReten, "
   Q1 = Q1 & " TipoDocs.EsRebaja, Documento.IdDoc,  Documento.NumDoc, NumDocHasta, CantBoletas "
   
   Q1 = Q1 & " FROM ((((Documento "
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib= TipoDocs.TipoLib AND Documento.TipoDoc= TipoDocs.TipoDoc) "
   Q1 = Q1 & " LEFT JOIN MovDocumento ON Documento.IdDoc = MovDocumento.IdDoc"
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " LEFT JOIN Cuentas  ON MovDocumento.IdCuenta = Cuentas.IdCuenta "
   Q1 = Q1 & JoinEmpAno(gDbType, "Cuentas", "MovDocumento") & " )"
   Q1 = Q1 & InnerJoinSuper
   
   Q1 = Q1 & " WHERE " & Wh & " AND (MovDocumento.EsTotalDoc = 0 OR MovDocumento.EsTotalDoc IS NULL)"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY Documento.TipoLib, Documento.TipoDoc, DTE, Documento.Giro, Documento.FacCompraRetParcial, Documento.IVAIrrecuperable, Exento, Afecto, " & SqlMonthLng("FEmision") & ", "
   Q1 = Q1 & IifActFijo & ", "
   'Q1 = Q1 & IifIVAIrrec & ", "
   Q1 = Q1 & IifNCVExenta & ", "
   Q1 = Q1 & IifSuperGroup
   Q1 = Q1 & " iif(Documento.TipoRetencion IS NULL, 0, Documento.TipoRetencion), "
   Q1 = Q1 & " EsRebaja, Documento.IdDoc,  Documento.NumDoc, NumDocHasta, CantBoletas "
   Q1 = Q1 & " ORDER BY " & SqlMonthLng("FEmision") & ", Documento.TipoLib, Documento.TipoDoc, Documento.IdDoc "

   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
   
      TipoDoc = vFld(Rs("TipoDoc"))
      
'      If vFld(Rs("NCV_Exenta")) <> 0 And ExpIVA Then     'si es una nota de crédito que sólo tiene monto exento se asimila a una Factura de Venta Exenta, desde el punto de vista del Form 29, ya que se traspasa a los mismos códigos
'         TipoDoc = TipoDocFVE                            'solicitado por Victro el 16 abr 2013)
'      End If
      
      For i = 0 To UBound(ResLib)
      
         EsIgual = (ResLib(i).Mes = vFld(Rs("Mes")))
         EsIgual = EsIgual And (ResLib(i).TipoLib = vFld(Rs("TipoLib")))
         EsIgual = EsIgual And (ResLib(i).TipoDoc = TipoDoc)
         EsIgual = EsIgual And ((ResLib(i).ActFijo <> 0 And vFld(Rs("ActFijo")) <> 0) Or (ResLib(i).ActFijo = 0 And vFld(Rs("ActFijo")) = 0))
         EsIgual = EsIgual And ((ResLib(i).Giro <> 0 And vFld(Rs("Giro")) <> 0) Or (ResLib(i).Giro = 0 And vFld(Rs("Giro")) = 0))
         EsIgual = EsIgual And ((ResLib(i).FacCompraRetParcial <> 0 And vFld(Rs("FacCompraRetParcial")) <> 0) Or (ResLib(i).FacCompraRetParcial = 0 And vFld(Rs("FacCompraRetParcial")) = 0))
         'EsIgual = EsIgual And ((ResLib(i).IVAIrrec <> 0 And vFld(Rs("IVAIrrecuperable")) <> 0) Or (ResLib(i).IVAIrrec = 0 And vFld(Rs("IVAIrrecuperable")) = 0))
         EsIgual = EsIgual And (ResLib(i).IVAIrrec = vFld(Rs("IVAIrrecuperable")))
         EsIgual = EsIgual And (ResLib(i).TipoReten = vFld(Rs("TipoReten")))
         If IifSuper <> "" Then
            EsIgual = EsIgual And (ResLib(i).EsSupermercado = vFld(Rs("EsSupermercado")))
         End If
         
         If EsIgual Then
            
            If vFld(Rs("TipoLib")) = LIB_COMPRAS Or vFld(Rs("TipoLib")) = LIB_RETEN Then    'compras y retenciones
               
               If vFld(Rs("IVAIrrecuperable")) <= IVAIRREC_PARCIAL Then   'compras sin IVA Irrec o IVA Irrec parcial y retenciones
                  
                  If Val(vFld(Rs("NumDocHasta"))) <> 0 Then
                     ResLib(i).CountTot = ResLib(i).CountTot + Val(vFld(Rs("NumDocHasta"))) - Val(vFld(Rs("NumDoc"))) + 1
                  Else
                     ResLib(i).CountTot = ResLib(i).CountTot + 1
                  End If
                  
               End If
                                    
               If vFld(Rs("IVAIrrecuperable")) >= IVAIRREC_PARCIAL Then    'compras con IVE Irrec parcial o total    (FCA 19 may 2015)
                  
                  If Val(vFld(Rs("NumDocHasta"))) <> 0 Then
                     ResLib(i).CountIVAIrrec = ResLib(i).CountIVAIrrec + Val(vFld(Rs("NumDocHasta"))) - Val(vFld(Rs("NumDoc"))) + 1
                  Else
                     ResLib(i).CountIVAIrrec = ResLib(i).CountIVAIrrec + 1
                  End If
                  
               End If
            
            ElseIf vFld(Rs("Giro")) <> 0 Then    'LIB_VENTAS del giro
                           
               If vFld(Rs("FacCompraRetParcial")) = 0 Then   'ventas del giro y fact compra con ret. total
                            
'                  If Not (vFld(Rs("Afecto")) = 0 And GetDiminutivoDoc(vFld(Rs("TipoLib")), TipoDoc) = "NCV") Then  'si es una NCV y sólo tiene exento, no se cuenta, FCA 15 abr 2013, solicitado por Victor Morales
                 
                     If Val(vFld(Rs("NumDocHasta"))) <> 0 And IsNumeric(vFld(Rs("NumDoc"))) Then
                        ResLib(i).CountTot = ResLib(i).CountTot + Val(vFld(Rs("NumDocHasta"))) - Val(vFld(Rs("NumDoc"))) + 1
                     ElseIf vFld(Rs("CantBoletas")) > 0 Then         'Para VPE
                        ResLib(i).CountTot = ResLib(i).CountTot + vFld(Rs("CantBoletas"))
                     Else
                        ResLib(i).CountTot = ResLib(i).CountTot + 1
                     End If
                     
'                  End If
                  
               Else                                          'facturas de compra de venta con ret. parcial
                  If Val(vFld(Rs("NumDocHasta"))) <> 0 Then
                     ResLib(i).CountRetParcial = ResLib(i).CountRetParcial + Val(vFld(Rs("NumDocHasta"))) - Val(vFld(Rs("NumDoc"))) + 1
                  Else
                     ResLib(i).CountRetParcial = ResLib(i).CountRetParcial + 1
                  End If
               
               End If
                
            Else           'ventas no del giro
               
'               If Not (vFld(Rs("Afecto")) = 0 And GetDiminutivoDoc(vFld(Rs("TipoLib")), TipoDoc) = "NCV") Then  'si es una NCV y sólo tiene exento, no se cuenta, FCA 15 abr 2013, solicitado por Victor Morales
                  
                  If Val(vFld(Rs("NumDocHasta"))) <> 0 Then
                     ResLib(i).CountTotNoGiro = ResLib(i).CountTotNoGiro + Val(vFld(Rs("NumDocHasta"))) - Val(vFld(Rs("NumDoc"))) + 1
                  ElseIf vFld(Rs("CantBoletas")) > 0 Then    'Para VPE
                     ResLib(i).CountTot = ResLib(i).CountTot + vFld(Rs("CantBoletas"))
                  Else
                     ResLib(i).CountTotNoGiro = ResLib(i).CountTotNoGiro + 1
                  End If
                  
'               End If
               
            End If
            
            'DTE
            If vFld(Rs("DTE")) <> 0 Then
               If Val(vFld(Rs("NumDocHasta"))) <> 0 Then
                  ResLib(i).CountDTE = ResLib(i).CountDTE + Val(vFld(Rs("NumDocHasta"))) - Val(vFld(Rs("NumDoc"))) + 1
               Else
                  ResLib(i).CountDTE = ResLib(i).CountDTE + 1
               End If
            End If
            
            If vFld(Rs("Exento")) <> 0 Then
            
               If vFld(Rs("TipoLib")) <> LIB_VENTAS Or vFld(Rs("Giro")) <> 0 Then
                  If Val(vFld(Rs("NumDocHasta"))) <> 0 Then
                     ResLib(i).CountExento = ResLib(i).CountExento + Val(vFld(Rs("NumDocHasta"))) - Val(vFld(Rs("NumDoc"))) + 1
                  Else
                     ResLib(i).CountExento = ResLib(i).CountExento + 1
                  End If
               Else
                  If Val(vFld(Rs("NumDocHasta"))) <> 0 Then
                     ResLib(i).CountExentoNoGiro = ResLib(i).CountExentoNoGiro + Val(vFld(Rs("NumDocHasta"))) - Val(vFld(Rs("NumDoc"))) + 1
                  Else
                     ResLib(i).CountExentoNoGiro = ResLib(i).CountExentoNoGiro + 1
                  End If
               End If
                              
            End If
            
            Exit For
            
         End If
         
      Next i
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   'ahora sumamos y desglosamos otros impuestos
   Call GenResOImp(Where, ResOImp)
   
   GenResLibros = True

End Function

'Esta función genera la estructura de datos de Otros Impuestos para llenar el reporte Resumen de Libros y
'para realizar la exportación a IVA estándar, tomando los movimientos del documento,
'considerando que ahora se ingresa el detalle.
'Se agrega parámetro UseAbs para cuando esta función es invocada desde el Resumen de IVA, en el Libro de Compras y ventas,
'en que no se requiere el Abs al calcular la diferencia entre Debe y Haber (no se usa nunca este parámetro)

Public Function GenResOImp(ByVal Where As String, ResOImp() As ResOImp_t, Optional ByVal UseAbs As Boolean = 1, Optional ByVal bCreaTmp As Boolean = 1) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim CurReg As String
   Dim NewReg As String
   Dim Wh As String
   Dim Dif As Double
   Dim AuxDif As Double
   Dim TmpTipoValor As String

   'TmpTipoValor = "tmp_TValor_" & gUsuario.Nombre
   TmpTipoValor = DbGenTmpName2(gDbType, "TValor_")
   If bCreaTmp Then ' 15 feb 2020
      Q1 = "DROP TABLE " & TmpTipoValor
      Call ExecSQL(DbMain, Q1)
   
      Q1 = "SELECT MovDocumento.IdMovDoc, TipoValor.Codigo, TipoValor.Valor as DescTipoValor, TipoValor.CodF29, TipoValor.CodF29_Adic, TipoValor.TipoIVARetenido, MovDocumento.IdEmpresa, MovDocumento.Ano"
      Q1 = Q1 & " INTO " & TmpTipoValor
      Q1 = Q1 & " FROM (MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc "
      Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
      Q1 = Q1 & " INNER JOIN TipoValor ON Documento.TipoLib= TipoValor.TipoLib AND MovDocumento.IdTipoValLib= TipoValor.Codigo  "
      Q1 = Q1 & " WHERE MovDocumento.IdEmpresa = " & gEmpresa.id & " AND MovDocumento.Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)
   End If
   
   Wh = "Documento.TipoLib IN(" & LIB_COMPRAS & "," & LIB_VENTAS & ") AND " & TmpTipoValor & ".Codigo >= " & LIBVENTAS_OTROSIMP  '(LIBCOMPRAS_OTROSIMP = LIBVENTAS_OTROSIMP)"
   If Where <> "" Then
      Wh = Wh & " AND " & Where
   End If
           
   Q1 = "SELECT " & SqlMonthLng("FEmision") & " As Mes, Documento.TipoLib, MovDocumento.IdTipoValLib, " & TmpTipoValor & ".CodF29, " & TmpTipoValor & ".CodF29_Adic, "
   Q1 = Q1 & " TipoDocs.EsRebaja, " & TmpTipoValor & ".DescTipoValor, Sum(MovDocumento.Debe) as SumDebe, Sum(MovDocumento.Haber) as SumHaber, "
   Q1 = Q1 & TmpTipoValor & ".TipoIVARetenido "
   Q1 = Q1 & " FROM (((Documento "
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc )"
   Q1 = Q1 & " LEFT JOIN MovDocumento ON Documento.IdDoc =  MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " LEFT JOIN " & TmpTipoValor & " ON MovDocumento.IdMovDoc = " & TmpTipoValor & ".IdMovDoc "
'   Q1 = Q1 & "  AND " & TmpTipoValor & ".IdEmpresa = MovDocumento.IdEmpresa AND " & TmpTipoValor & ".Ano = MovDocumento.Ano )"
   Q1 = Q1 & JoinEmpAno(gDbType, TmpTipoValor, "MovDocumento") & " )" ' 14 feb 2020
   Q1 = Q1 & " WHERE " & Wh & " AND (MovDocumento.EsTotalDoc = 0 OR MovDocumento.EsTotalDoc IS NULL)"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY " & SqlMonthLng("FEmision") & ", Documento.TipoLib, TipoDocs.EsRebaja, MovDocumento.IdTipoValLib, " & TmpTipoValor & ".CodF29, " & TmpTipoValor & ".CodF29_Adic, " & TmpTipoValor & ".DescTipoValor, "
   Q1 = Q1 & TmpTipoValor & ".TipoIVARetenido "
   Q1 = Q1 & " ORDER BY " & SqlMonthLng("FEmision") & ", Documento.TipoLib, MovDocumento.IdTipoValLib "

   Set Rs = OpenRs(DbMain, Q1)
      
   i = -1
   CurReg = ""
   
   ReDim ResOImp(30)
   
   Do While Rs.EOF = False
   
      If Not (vFld(Rs("TipoLib")) = LIB_COMPRAS And vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAACTFIJO) Then
         
         NewReg = vFld(Rs("Mes")) & "-" & vFld(Rs("TipoLib")) & "-" & vFld(Rs("IdTipoValLib"))
                                 
         If NewReg <> CurReg Then
                     
            i = i + 1
            
            If i > UBound(ResOImp) Then
               ReDim Preserve ResOImp(i + 10)
            End If
               
            ResOImp(i).Mes = vFld(Rs("Mes"))
            ResOImp(i).TipoLib = vFld(Rs("TipoLib"))
            ResOImp(i).CodValLib = vFld(Rs("IdTipoValLib"))
            ResOImp(i).DescValLib = vFld(Rs("DescTipoValor"))
            ResOImp(i).CodF29 = vFld(Rs("CodF29"))
            ResOImp(i).CodF29_Adic = vFld(Rs("CodF29_Adic"))
            ResOImp(i).TipoIVARetenido = vFld(Rs("CodF29_Adic"))
                     
            CurReg = NewReg
            
         End If
               
         Dif = vFld(Rs("SumDebe")) - vFld(Rs("SumHaber"))
         If UseAbs Then   'lo típico
            Dif = Abs(Dif)
         End If
         
         ResOImp(i).Valor = ResOImp(i).Valor + IIf(vFld(Rs("EsRebaja")) <> 0, Dif * -1, Dif)
      End If
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
'   Call ExecSQL(DbMain, "DROP TABLE " & TmpTipoValor)
   
   If i >= 0 Then
      ReDim Preserve ResOImp(i)
   Else
      ReDim ResOImp(0)
      Exit Function
   End If

End Function

'Esta función es idéntica a GenResOImp pero en este caso desgloza los impuestos por si EsRecuperable o No

Public Function GenResOImpEsRecup(ByVal Where As String, ResOImp() As ResOImp_t, Optional ByVal UseAbs = True) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim CurReg As String
   Dim NewReg As String
   Dim Wh As String
   Dim Dif As Double
   Dim AuxDif As Double
   Dim TmpTipoValor As String

   'TmpTipoValor = "tmp_TValor_" & gUsuario.Nombre
   TmpTipoValor = DbGenTmpName2(gDbType, "TValor_")
   
   Q1 = "DROP TABLE " & TmpTipoValor
   Call ExecSQL(DbMain, Q1)

   Q1 = "SELECT MovDocumento.IdMovDoc, TipoValor.Codigo, TipoValor.Valor as DescTipoValor, TipoValor.CodF29, TipoValor.CodF29_Adic, TipoValor.TipoIVARetenido, MovDocumento.IdEmpresa, MovDocumento.Ano INTO " & TmpTipoValor
   Q1 = Q1 & " FROM (MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " INNER JOIN TipoValor ON Documento.TipoLib= TipoValor.TipoLib AND MovDocumento.IdTipoValLib= TipoValor.Codigo  "
   Q1 = Q1 & " WHERE MovDocumento.IdEmpresa = " & gEmpresa.id & " AND MovDocumento.Ano = " & gEmpresa.Ano
   Call ExecSQL(DbMain, Q1)

   Wh = "Documento.TipoLib IN(" & LIB_COMPRAS & "," & LIB_VENTAS & ") AND " & TmpTipoValor & ".Codigo >= " & LIBVENTAS_OTROSIMP  '(LIBCOMPRAS_OTROSIMP = LIBVENTAS_OTROSIMP)"
   If Where <> "" Then
      Wh = Wh & " AND " & Where
   End If

           
   Q1 = "SELECT " & SqlMonthLng("FEmision") & " As Mes, Documento.TipoLib, MovDocumento.IdTipoValLib, MovDocumento.EsRecuperable, " & TmpTipoValor & ".CodF29, " & TmpTipoValor & ".CodF29_Adic, "
   Q1 = Q1 & " TipoDocs.EsRebaja, " & TmpTipoValor & ".DescTipoValor, Sum(MovDocumento.Debe) as SumDebe, Sum(MovDocumento.Haber) as SumHaber, "
   Q1 = Q1 & TmpTipoValor & ".TipoIVARetenido "
   Q1 = Q1 & " FROM (((Documento "
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc )"
   Q1 = Q1 & " LEFT JOIN MovDocumento ON Documento.IdDoc=MovDocumento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "Documento", "MovDocumento") & " )"
   Q1 = Q1 & " LEFT JOIN " & TmpTipoValor & " ON MovDocumento.IdMovDoc = " & TmpTipoValor & ".IdMovDoc "
'   Q1 = Q1 & "  AND " & TmpTipoValor & ".IdEmpresa = MovDocumento.IdEmpresa AND " & TmpTipoValor & ".Ano = MovDocumento.Ano )"
   Q1 = Q1 & JoinEmpAno(gDbType, TmpTipoValor, "MovDocumento", True) & " )" ' 14 feb 2020
   Q1 = Q1 & " WHERE " & Wh & " AND (MovDocumento.EsTotalDoc = 0 OR MovDocumento.EsTotalDoc IS NULL)"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY " & SqlMonthLng("FEmision") & ", Documento.TipoLib, TipoDocs.EsRebaja, MovDocumento.IdTipoValLib, MovDocumento.EsRecuperable, " & TmpTipoValor & ".CodF29, " & TmpTipoValor & ".CodF29_Adic, " & TmpTipoValor & ".DescTipoValor, "
   Q1 = Q1 & TmpTipoValor & ".TipoIVARetenido "
   Q1 = Q1 & " ORDER BY " & SqlMonthLng("FEmision") & ", Documento.TipoLib, MovDocumento.IdTipoValLib "

   Set Rs = OpenRs(DbMain, Q1)
      
   i = -1
   CurReg = ""
   
   ReDim ResOImp(30)
   
   Do While Rs.EOF = False
   
      If Not (vFld(Rs("TipoLib")) = LIB_COMPRAS And vFld(Rs("IdTipoValLib")) = LIBCOMPRAS_IVAACTFIJO) Then
         
         NewReg = vFld(Rs("Mes")) & "-" & vFld(Rs("TipoLib")) & "-" & vFld(Rs("IdTipoValLib")) & "-" & vFld(Rs("EsRecuperable"))
                                 
         If NewReg <> CurReg Then
                     
            i = i + 1
            
            If i > UBound(ResOImp) Then
               ReDim Preserve ResOImp(i + 10)
            End If
               
            ResOImp(i).Mes = vFld(Rs("Mes"))
            ResOImp(i).TipoLib = vFld(Rs("TipoLib"))
            ResOImp(i).CodValLib = vFld(Rs("IdTipoValLib"))
            ResOImp(i).DescValLib = vFld(Rs("DescTipoValor"))
            ResOImp(i).EsRecuperable = vFld(Rs("EsRecuperable"))
            ResOImp(i).CodF29 = vFld(Rs("CodF29"))
            ResOImp(i).CodF29_Adic = vFld(Rs("CodF29_Adic"))
            ResOImp(i).TipoIVARetenido = vFld(Rs("TipoIVARetenido"))
            
                     
            CurReg = NewReg
            
         End If
               
         Dif = vFld(Rs("SumDebe")) - vFld(Rs("SumHaber"))
         If UseAbs Then   'lo típico
            Dif = Abs(Dif)
         End If
         
         If UseAbs Then
            ResOImp(i).Valor = ResOImp(i).Valor + IIf(vFld(Rs("EsRebaja")) <> 0, Dif * -1, Dif)
         Else
            ResOImp(i).Valor = ResOImp(i).Valor + Dif
         End If
      End If
      
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Call ExecSQL(DbMain, "DROP TABLE " & TmpTipoValor)
   
   
   If i >= 0 Then
      ReDim Preserve ResOImp(i)
   Else
      ReDim ResOImp(0)
      Exit Function
   End If

End Function
Public Function GenExportF29(ByVal Where As String, ResLibCod() As ResLibCod_t, ByVal AnoMes As Long, Optional ByVal Msg As Boolean = True) As Boolean
   Dim ResLib() As ResLib_t
   Dim ResOImp() As ResOImp_t
   Dim i As Integer
   Dim j As Integer
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Rc As Integer
   Dim FLast As Long

   GenExportF29 = False
   
   'Where restringe año y mes
   
   Rc = GenResLibros(Where, ResLib, ResOImp, True)
   
   If Rc = False Then
      ReDim ResLibCod(0)
      Exit Function
   End If
   
   ReDim ResLibCod(40)        'en AppendCodF29 se hace ReDim si es necesario
   
   FLast = DateAdd("m", 1, AnoMes) - 1
   
   ' primero ponemos todos los códigos en el arreglo ResLibCod
   Q1 = "SELECT CodF29Count, CodF29Neto, CodF29IVA, CodF29IVADTE, CodF29AFCount, CodF29AFIVA, CodF29ExCount, CodF29Exento, CodF29RetHon, CodF29RetDieta, CodF29IVARet3ro, CodF29CountNoGiro, CodF29NetoNoGiro, CodF29IVANoGiro, CodF29CountSuper, CodF29IVASuper "
   Q1 = Q1 & " FROM TipoDocs"
   Set Rs = OpenRs(DbMain, Q1)
   Do Until Rs.EOF
      For i = 0 To Rs.Fields.Count - 1
         If vFld(Rs(i)) Then
            Call AppendCodF29(ResLibCod, vFld(Rs(i)), 0)
         End If
      Next i
   
      Rs.MoveNext
   Loop
   Call CloseRs(Rs)
      
   j = -1
   
   For i = 0 To UBound(ResLib)
   
      'por ahora asumimos que todas las facturas afectas tienen derecho a crédito en su totalidad
      'por ahora asumimos que todas las facturas de compra del libro de ventas tienen retención total (no genera débito)
      'por ahora asumimos que todas las facturas de compra del libro de compras tienen retención total (no genera crédito)
   
      Call AppendCodF29(ResLibCod, ResLib(i).CodF29IVADTE, ResLib(i).IVADTE)
      Call AppendCodF29(ResLibCod, ResLib(i).CodF29CountDTE, ResLib(i).CountDTE)
      Call AppendCodF29(ResLibCod, ResLib(i).CodF29NetoDTE, ResLib(i).NetoDTE)
      Call AppendCodF29(ResLibCod, ResLib(i).CodF29IVAIrrecDTE, ResLib(i).IVAIrrecDTE)
      
      If ResLib(i).ActFijo <> 0 Then
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29AFCount, ResLib(i).CountTot)
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29AFIVA, ResLib(i).IVA)
      
      ElseIf ResLib(i).TipoReten = 0 Then    'compras o ventas sin activo fijo
   
         If ResLib(i).EsSupermercado = 0 Or AnoMes < gFechaInicioTraspasoSupermercados Then
            Call AppendCodF29(ResLibCod, ResLib(i).CodF29Count, ResLib(i).CountTot)
            'Call AppendCodF29(ResLibCod, ResLib(i).CodF29Neto, ResLib(i).Afecto + ResLib(i).Exento)                       'victor morales 29 abr 2013
            Call AppendCodF29(ResLibCod, ResLib(i).CodF29Neto, ResLib(i).Afecto)
            Call AppendCodF29(ResLibCod, ResLib(i).CodF29IVA, ResLib(i).IVA)
         Else
            Call AppendCodF29(ResLibCod, ResLib(i).CodF29CountSuper, ResLib(i).CountTot)
            Call AppendCodF29(ResLibCod, ResLib(i).CodF29IVASuper, ResLib(i).IVA)
         End If
         
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29CountNoGiro, ResLib(i).CountTotNoGiro)
         'Call AppendCodF29(ResLibCod, ResLib(i).CodF29NetoNoGiro, ResLib(i).AfectoNoGiro + ResLib(i).ExentoNoGiro)     'victor morales 29 abr 2013
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29NetoNoGiro, ResLib(i).AfectoNoGiro)
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29IVANoGiro, ResLib(i).IVANoGiro)
         
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29ExCount, ResLib(i).CountExento)
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29Exento, ResLib(i).Exento)
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29ExCountNoGiro, ResLib(i).CountExentoNoGiro)
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29ExentoNoGiro, ResLib(i).ExentoNoGiro)
         
         'Call AppendCodF29(ResLibCod, ResLib(i).CodF29IVARet3ro, ResLib(i).IVA)   'antes se asumía que se retenía todo para facturas de compra del libro de compras, ahora el usuario indica cuánto retiene por el retenido parcial y el retenido total
         
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29CountRetParcial, ResLib(i).CountRetParcial)
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29NetoRetParcial, ResLib(i).NetoRetParcial)
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29DifIVARetParcial, ResLib(i).DifIVARetParcial)
         
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29CountIVAIrrec, ResLib(i).CountIVAIrrec)
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29NetoIVAIrrec, ResLib(i).NetoIVAIrrec)
         
      
      ElseIf ResLib(i).TipoReten = TR_HONORARIOS Then
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29RetHon, ResLib(i).OtroImp)
         
      ElseIf ResLib(i).TipoReten = TR_DIETA Then
         Call AppendCodF29(ResLibCod, ResLib(i).CodF29RetDieta, ResLib(i).OtroImp)
      
      End If
      
      j = j + 1
      
      If j > UBound(ResLibCod) Then
         ReDim Preserve ResLibCod(j + 10)
      End If

   Next i
   
   'Impuesto único
   If gCtasBas.IdCtaImpUnico = 0 Then
      If Msg Then
         MsgBox1 "Falta definir la cuenta de Impuesto Único a los Trabajadores. No se exportará este valor.", vbExclamation + vbOKOnly
      End If
   
   Else
      If Msg Then
         'MsgBox1 "Recuerde que debe tener cuadrada la cuenta de Impuesto Único de los meses anteriores, para que el saldo de esta cuenta sea correcto.", vbInformation + vbOKOnly
         MsgBox1 "Recuerde que debe tener saldada la cuenta de Impuesto Único de los meses anteriores, para que el saldo a traspasar sea el correcto." & vbCrLf & vbCrLf & "El saldo de esta cuenta se calcula a partir de los comprobantes cuyo Tipo de Ajuste sea TRIBUTARIO o AMBOS.", vbInformation + vbOKOnly       'Victor Morales 13 dic 2012
      End If
      
      Q1 = "SELECT Sum(Haber) - Sum(Debe) as Saldo "
      Q1 = Q1 & " FROM MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp"
      Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      Q1 = Q1 & " WHERE IdCuenta=" & gCtasBas.IdCtaImpUnico
      Q1 = Q1 & " AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & "))"
      Q1 = Q1 & " AND Fecha <= " & FLast
      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         Call AppendCodF29(ResLibCod, CODF29_IMPUNICO, vFld(Rs("Saldo")))
      End If
      
      Call CloseRs(Rs)
   
   End If
   
   'Otros impuestos
   
   'primero ponemos todos lod códigos de Otros Impuestos en el arreglo ResLibCod
   
   Q1 = "SELECT CodF29 FROM TipoValor WHERE CodF29 <> 0 AND (NOT CodF29 IS NULL)"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      If vFld(Rs("CodF29")) <> 0 Then
         Call AppendCodF29(ResLibCod, vFld(Rs("CodF29")), 0)
      End If
   
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
      
   Q1 = "SELECT CodF29_Adic FROM TipoValor WHERE CodF29_Adic <> 0 AND (NOT CodF29_Adic IS NULL)"
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      If vFld(Rs("CodF29_Adic")) <> 0 Then
         Call AppendCodF29(ResLibCod, vFld(Rs("CodF29_Adic")), 0)
      End If
   
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   'ahora los valores de cada tipo de impuesto
   For i = 0 To UBound(ResOImp)
   
      Call AppendCodF29(ResLibCod, ResOImp(i).CodF29, ResOImp(i).Valor)
      Call AppendCodF29(ResLibCod, ResOImp(i).CodF29_Adic, ResOImp(i).Valor)
            
   Next i
   
   GenExportF29 = True

End Function

Private Sub AppendCodF29(ResLibCod() As ResLibCod_t, ByVal CodF29 As Integer, ByVal Valor As Double)
   Dim i As Integer
   Dim IdxCodF29 As Integer
   Dim IdxVacio As Integer

   IdxCodF29 = -1
   IdxVacio = -1
   
'   If CodF29 = 142 Then
'      Beep
'   End If
   
   If CodF29 <> 0 Then
   
      For i = 0 To UBound(ResLibCod)
      
         If Abs(CodF29) = Abs(ResLibCod(i).CodF29) Then  'ya estaba en la lista, lo sumamos
         
            ResLibCod(i).Valor = ResLibCod(i).Valor + Valor * Sgn(CodF29)
            IdxCodF29 = i
            Exit For
         End If
         
         If ResLibCod(i).CodF29 = 0 Then    'termina la lista
            IdxVacio = i
            Exit For
         End If
      Next i
   
      If IdxCodF29 < 0 Then    'no se encontró, lo agregamos
      
         If IdxVacio < 0 Then
            IdxVacio = UBound(ResLibCod) + 1
            ReDim Preserve ResLibCod(IdxVacio + 10)
         End If
         
         ResLibCod(IdxVacio).CodF29 = CodF29
         ResLibCod(IdxVacio).Valor = Valor * Sgn(CodF29)
         
      End If
      
   End If

End Sub
Public Function GetInsertEntidad(ByVal Rut As String, ByVal Nombre As String, ByVal ClasifEnt As Integer, Optional ByVal NotValidRut As Boolean = False) As Long
   Dim Rs As Recordset
   Dim Q1 As String
   Dim IdEnt As Long
   Dim AuxRut As String
   Dim Codigo As String
   Dim RsCod As Recordset
   Dim Max As Long
   Dim FldArray(5) As AdvTbAddNew_t
   
   'Rut viene como string con dígito verificador, sin guión, y relleno con ceros a la izquierda
   If NotValidRut = False Then    'es Rut válido
      AuxRut = vFmtCID(Rut)
      
      Codigo = Trim(Rut)
      
   Else
      AuxRut = Trim(Rut)
      
      'veamos si es RUT ficticio creado en HR-LAU, de la forma "F000000000000001"
      
      If Len(AuxRut) > 12 Then    'Campo Rut tiene largo máximo 12 en DB Contabilidad
         AuxRut = "RF" & Right(AuxRut, 10)   'generamos un rut ficticio
      End If
      
      Codigo = AuxRut
      
   End If
   
   'veamos si ya existe una entidad con este Rut
   Q1 = "SELECT IdEntidad FROM Entidades WHERE Rut='" & AuxRut & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      'existe la entidad
      
      IdEnt = vFld(Rs("IdEntidad"))
      
      Call CloseRs(Rs)
      
      Q1 = "UPDATE Entidades SET Clasif" & ClasifEnt & "=1 WHERE IdEntidad =" & IdEnt
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      Call ExecSQL(DbMain, Q1)
      
      GetInsertEntidad = IdEnt
      Exit Function
      
   Else
      Call CloseRs(Rs)
      
      'veamos si ya existe una entidad con este Código
      Q1 = "SELECT IdEntidad FROM Entidades WHERE Codigo='" & ParaSQL(Left(ReplaceStr(Codigo, " ", ""), 15)) & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
      Set Rs = OpenRs(DbMain, Q1)
      
      If Rs.EOF = False Then
         
         'ya existe entidad con este código, creamos cod. ficticio correlativo
         Q1 = "SELECT Max(Codigo) As MaxCodigo FROM Entidades WHERE Left(Codigo,4)='#CF#'"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
         Set RsCod = OpenRs(DbMain, Q1)
         
         Max = 0
         If RsCod.EOF = False Then
            Max = Val(Mid(vFld(RsCod("MaxCodigo")), 5))
         End If
         Max = Max + 1
         Codigo = "#CF#" & Max
      
         Call CloseRs(RsCod)
         
      End If
      
      Call CloseRs(Rs)
                 
'      Set Rs = DbMain.OpenRecordset("Entidades", dbOpenTable)
'      Rs.AddNew
'
'      GetInsertEntidad = vFld(Rs("IdEntidad"))
'
'      Rs.Fields("Rut") = AuxRut
'      Rs.Fields("NotValidRut") = NotValidRut
'      Rs.Fields("Codigo") = ParaSQL(ReplaceStr(Codigo, " ", ""))
'      Rs.Fields("Nombre") = ParaSQL(Nombre)
'      Rs.Fields("Clasif" & ClasifEnt) = 1
'
'      Rs.Update
'      Rs.Close
      
      FldArray(0).FldName = "NotValidRut"
      FldArray(0).FldValue = NotValidRut
      FldArray(0).FldIsNum = True
      
      FldArray(1).FldName = "RUT"
      FldArray(1).FldValue = AuxRut
      FldArray(1).FldIsNum = False
            
      FldArray(2).FldName = "IdEmpresa"
      FldArray(2).FldValue = gEmpresa.id
      FldArray(2).FldIsNum = True
                        
      FldArray(3).FldName = "Codigo"
      FldArray(3).FldValue = ParaSQL(Left(ReplaceStr(Codigo, " ", ""), 15))
      FldArray(3).FldIsNum = False
                        
      FldArray(4).FldName = "Nombre"
      FldArray(4).FldValue = ParaSQL(Left(Nombre, 100))
      FldArray(4).FldIsNum = False
                        
      FldArray(5).FldName = "Clasif" & ClasifEnt
      FldArray(5).FldValue = 1
      FldArray(5).FldIsNum = True
                        
      GetInsertEntidad = AdvTbAddNewMult(DbMain, "Entidades", "IdEntidad", FldArray)
      
   End If

End Function
Public Sub ImportF29_Old(ByVal Mes As Integer)
   Dim Rc As Integer
   Dim Rs As Recordset

   If MsgBox1("Para realizar la importación desde HR-IVA, nadie debe estar trabajando en esta empresa en HR-IVA." & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If

   Rc = LinkLau()
   If Rc Then
      Exit Sub
   End If
   
   'probamos si no hay conflictos con la base de datos LAU
   On Error Resume Next
   
   Set Rs = OpenRs(DbMain, "SELECT sRUT FROM LAU_mPersonas", False)
   
   If ERR Then
      Call CloseRs(Rs)
      MsgBox1 "Problemas para realizar la importación. Verifique que nadie esté trabajando en esta empresa en HR-IVA.", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   Call CloseRs(Rs)
   
   On Error GoTo 0
   
   Call ImportLibF29(Mes, LIB_COMPRAS)
   Call ImportLibF29(Mes, LIB_VENTAS)


   'Rc = UnLinkLau()
   
   MsgBox1 "Importación finalizada.", vbInformation + vbOKOnly
   
End Sub
Public Function ImportLibF29(ByVal Mes As Integer, ByVal TipoLib As Integer) As Boolean
   Dim Rs As Recordset
   Dim Q1 As String
   Dim F1 As Long, F2 As Long
   Dim DefCtaBas As Boolean
   Dim Rc As Integer
   Dim NDocsCompras As Long
   Dim NDocsVentas As Long
   Dim NDocsReten As Long
   Dim DelFrom As String, DelWhere As String

   ImportLibF29 = False
   
   If Mes <= 0 Or TipoLib <= 0 Then
      Exit Function
   End If
   
   Rc = LinkLau()
   If Rc Then
      MsgBox1 "No se pudo vincular la información del Formulario 29.", vbExclamation
      Exit Function
   End If
   
   'probamos si no hay conflictos con la base de datos LAU
   On Error Resume Next
   
   Set Rs = OpenRs(DbMain, "SELECT sRUT FROM LAU_mPersonas", False)
   
   If ERR Then
      Call CloseRs(Rs)
      MsgBox1 "Problemas para realizar la importación. Verifique que nadie esté trabajando en esta empresa en HR-IVA.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   Call CloseRs(Rs)
   
   On Error GoTo 0
   
   DefCtaBas = True
   
   'Verificamos que se hayan definido las cuentas básicas en la configuración
   Call LoadCuentasDef(TipoLib)
   
   If (TipoLib = LIB_COMPRAS Or TipoLib = LIB_VENTAS) And (lIdCtaAfecto = 0 Or lIdCtaExento = 0 Or lIdCtaTotal = 0 Or lIdCtaIVA = 0 Or lIdCtaOtrosImp = 0) Then
      DefCtaBas = False
      
   ElseIf TipoLib = LIB_RETEN And (lIdCtaHonSinRet = 0 Or lIdCtaBruto = 0 Or lIdCtaImpRet = 0 Or lIdCtaNetoHon = 0 Or lIdCtaNetoDieta = 0) Then
      DefCtaBas = False
   End If
   
   If Not DefCtaBas Then
      MsgBox1 "No se han definido las cuentas básicas para el " & gTipoLib(TipoLib) & "." & vbNewLine & vbNewLine & "Es necesario hacer esta definición antes de realizar la importación desde HR-IVA. Para tal efecto, utilice la opción 'Definición de Cuentas Básicas' incluída en la Configuración Incial del Sistema.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   Rc = MsgBox1("¡¡¡ ATENCIÓN !!!" & vbLf & vbLf & "Antes de realizar la importación desde HR-IVA, se eliminarán TODOS los documentos IMPORTADOS del " & gTipoLib(TipoLib) & " del mes de " & gNomMes(Mes) & " que se encuentren en estado PENDIENTE o ANULADO." & vbNewLine & vbNewLine & "¿Está seguro que desea continuar?", vbExclamation + vbYesNo + vbDefaultButton2)
   If Rc = vbNo Then
      Exit Function
   End If
      
   'eliminamos docs del mes de TipoLib
   Call FirstLastMonthDay(DateSerial(gEmpresa.Ano, Mes, 1), F1, F2)
      
   DelFrom = " MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc "
   DelFrom = DelFrom & JoinEmpAno(gDbType, "Documento", "MovDocumento")
   DelWhere = " WHERE FEmision BETWEEN " & F1 & " AND " & F2
   DelWhere = DelWhere & " AND TipoLib = " & TipoLib
   DelWhere = DelWhere & " AND Estado IN (" & ED_PENDIENTE & "," & ED_ANULADO & ")"
   DelWhere = DelWhere & " AND FImporF29 <> 0 "
   DelWhere = DelWhere & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   
   Call DeleteJSQL(DbMain, "MovDocumento", DelFrom, DelWhere)
   
   Q1 = " WHERE FEmision BETWEEN " & F1 & " AND " & F2
   Q1 = Q1 & " AND TipoLib = " & TipoLib
   Q1 = Q1 & " AND Estado IN (" & ED_PENDIENTE & "," & ED_ANULADO & ")"
   Q1 = Q1 & " AND FImporF29 <> 0 "
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Call DeleteSQL(DbMain, "Documento", Q1)
   
   'insertamos los nuevos docs.
   
   Q1 = vbLf & vbLf & "Estos quedan de color azul y para ver su detalle puede usar la lupa."
   
   Select Case TipoLib
   
      Case LIB_COMPRAS
      
         NDocsCompras = ImportF29Compras(Mes)
         
         MsgBox1 "Se han importado " & NDocsCompras & " documentos del Libro de Compras del mes de " & gNomMes(Mes) & "." & Q1, vbInformation + vbOKOnly
         
      Case LIB_VENTAS
      
         NDocsVentas = NDocsVentas + ImportF29VentasNac(Mes)
         NDocsVentas = NDocsVentas + ImportF29VentasExp(Mes)
         NDocsVentas = NDocsVentas + ImportF29VentasBoleta(Mes)
         NDocsVentas = NDocsVentas + ImportF29VentasDevBoleta(Mes)
         
         MsgBox1 "Se han importado " & NDocsVentas & " documentos del Libro de Ventas del mes de " & gNomMes(Mes) & "." & Q1, vbInformation + vbOKOnly
         
      Case LIB_RETEN
         NDocsReten = ImportF29Retenciones(Mes)
         MsgBox1 "Se han importado " & NDocsReten & " documentos del Libro de Retenciones del mes de " & gNomMes(Mes) & "." & Q1, vbInformation + vbOKOnly
         
   End Select
   
   ImportLibF29 = True
   

End Function
Private Function ImportF29Compras(ByVal Mes As Integer) As Long
   Dim Rs As Recordset
   Dim RsDoc As Recordset
   Dim Q1 As String
   Dim IdDoc As Long
   Dim i As Integer, j As Integer
   Dim TipoDoc As Integer
   Dim orden As Integer
   Dim TipoLib As Integer
   Dim NDocsImp As Long
   Dim IdEnt As Long
   Dim RsOldDoc As Recordset
   Dim CeroAs As String
   Dim IdSucursal As Long
   Dim DTE As Integer
   Dim FldArray(3) As AdvTbAddNew_t
   
   TipoLib = LIB_COMPRAS

   'Facturas de Compra Nacionales e Importadas: tabla LAU_FacturaCompras
   
   If gEmpresa.Ano < 2005 Then
      CeroAs = " 0 as "
   End If
   
   Q1 = "SELECT LAU_FacturaCompras.sRUT, nlNumItemRefExp, nbTipoDoc, nbTipoCompra, nlNumDocTxt, dtFechaDoc, nlNumDocRef, sCodDatoUsuario, "
   Q1 = Q1 & " ndTotNeto, ndNetoOtros, ndNetoActFijo, ndNetoRealizable, ndNetoGastos, ndTotExento, ndExenOtros, ndExenActFijo, "
   Q1 = Q1 & " ndExenRealizable, ndExenGastos, ndIVA, ndTotOtrosImpto, nMesWrk, "
   Q1 = Q1 & " ndTotIvaRetenido, ndL24C39, ndL25C42, ndBaseHarina, ndBaseCarne, ndTotAnticipos, "
   Q1 = Q1 & " nd12Harina, ndIVAIrrecuperable, nd4Carne, ndTotDocumen, bDocE, " & CeroAs & " ndIvaRealizable, " & CeroAs & " ndIvaActFijo, " & CeroAs & " ndIvaGastos, "
   Q1 = Q1 & CeroAs & " ndIvaOtros, " & CeroAs & " ndIvaIrreRealizable, " & CeroAs & " ndIvaIrreActFijo, " & CeroAs & " ndIvaIrreGastos, " & CeroAs & " ndIvaIrreOtros, " & CeroAs & " ndImpto_c28, " & CeroAs & " ndImpto_c575, "
   Q1 = Q1 & CeroAs & " ndImpto_c576, " & CeroAs & " ndImpto_c574, " & CeroAs & " ndImpto_c33, " & CeroAs & " ndImpto_c580, " & CeroAs & " ndImpto_c149, " & CeroAs & " ndImpto_c582, " & CeroAs & " ndImpto_c85, " & CeroAs & " ndImpto_c127, "
   Q1 = Q1 & CeroAs & " ndImpto_c544, " & CeroAs & " ndIvaRetTot, " & CeroAs & " ndIvaRetPar, "
   Q1 = Q1 & " LAU_mPersonas.sNombre As NombrePers, "
   Q1 = Q1 & " LAU_ReferenciaExp.sNombre As NombreRef, LAU_ReferenciaExp.sNumRef, LAU_FacturaCompras.nlNumItemSuc, LAU_Sucursales.sCodSuc  "
   Q1 = Q1 & " FROM ((LAU_FacturaCompras "
   Q1 = Q1 & " LEFT JOIN LAU_mPersonas ON LAU_FacturaCompras.sRUT = LAU_mPersonas.sRUT) "
   Q1 = Q1 & " LEFT JOIN LAU_ReferenciaExp ON LAU_FacturaCompras.nlNumItemRefExp = LAU_ReferenciaExp.nlNumItem) "
   Q1 = Q1 & " LEFT JOIN LAU_Sucursales ON LAU_FacturaCompras.nlNumItemSuc = LAU_Sucursales.nlNumItem  "
   Q1 = Q1 & " WHERE bNula = 0 AND nMesWrk = " & Mes
   'Q1 = Q1 & " AND nbTipoCompra = " & LAU_COMPRASNAC
   Q1 = Q1 & " ORDER BY LAU_FacturaCompras.nlNumItem"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 1
   lIdCtaIVAIrrec = 0
   NDocsImp = 0
   
   Do While Rs.EOF = False
   
      If vFld(Rs("nlNumItemRefExp")) <> 0 And vFld(Rs("nbTipoDoc")) = LAU_COMP_FACT Then  'Factura de importación
         TipoDoc = GetTipoDocFromLAU(TipoLib, LAU_COMP_FACTIMP)
      Else  'nacionales y otros
         TipoDoc = GetTipoDocFromLAU(TipoLib, vFld(Rs("nbTipoDoc")))
      End If
      
      If vFld(Rs("nlNumItemRefExp")) <> 0 Then   'importaciones
         IdEnt = GetInsertEntidad(vFld(Rs("sNumRef"), True), vFld(Rs("NombreRef"), True), ENT_PROVEEDOR, True)
      Else
         IdEnt = GetInsertEntidad(vFld(Rs("sRUT"), True), vFld(Rs("NombrePers"), True), ENT_PROVEEDOR)
      End If

      'obtenemos la sucursal si el número en HR coincide con el código en Contabilidad
      IdSucursal = 0
      If vFld(Rs("sCodSuc")) <> "" Then
         IdSucursal = GetIdSucursal(vFld(Rs("sCodSuc")))
      End If
      
      DTE = IIf(vFld(Rs("bDocE")) <> 0, -1, 0)
     
      'vemos si el documento ya está en el libro
      
      Q1 = "SELECT IdDoc FROM Documento WHERE TipoLib=" & LIB_COMPRAS & " AND TipoDoc=" & TipoDoc & " AND NumDoc='" & vFld(Rs("nlNumDocTxt")) & "' AND IdEntidad =" & IdEnt & " AND DTE = " & DTE
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set RsOldDoc = OpenRs(DbMain, Q1)
      
      If RsOldDoc.EOF = True Then  'no existe, lo agregamos
   
'         Set RsDoc = DbMain.OpenRecordset("Documento")
'         RsDoc.AddNew
'
'         IdDoc = vFld(RsDoc("IdDoc"))
'         RsDoc.Fields("IdUsuario") = gUsuario.IdUsuario
'         RsDoc.Fields("FechaCreacion") = CLng(Int(Now))
'
'         On Error Resume Next
'
'         For j = 1 To 10
'            RsDoc.Fields("NumDoc") = CLng(Rnd * 321654356#)
'
'            RsDoc.Update
'
'            If Err = 0 Then
'               Exit For
'            End If
'         Next j
'
'         On Error GoTo 0
'
'         RsDoc.Close
'
'         Set RsDoc = Nothing
         
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
         
         NDocsImp = NDocsImp + 1
      
         'datos documento
               
         Q1 = "UPDATE Documento SET "
         Q1 = Q1 & "  TipoLib = " & TipoLib
         Q1 = Q1 & ", TipoDoc = " & TipoDoc
         Q1 = Q1 & ", NumDoc = '" & vFld(Rs("nlNumDocTxt")) & "'"
         Q1 = Q1 & ", NumDocHasta = '0'"
         Q1 = Q1 & ", IdEntidad = " & IdEnt
         Q1 = Q1 & ", TipoEntidad = " & ENT_PROVEEDOR
         Q1 = Q1 & ", FEmision = " & CLng(Int(DateSerial(gEmpresa.Ano, vFld(Rs("nMesWrk")), 1)))
         Q1 = Q1 & ", FVenc = 0"
         Q1 = Q1 & ", NumDocRef = '" & vFld(Rs("nlNumDocRef")) & "'"
         Q1 = Q1 & ", Descrip = '" & vFld(Rs("sCodDatoUsuario"), True) & "'"
         Q1 = Q1 & ", Estado = " & ED_PENDIENTE
         Q1 = Q1 & ", Exento = " & vFld(Rs("ndTotExento"))
         Q1 = Q1 & ", IdCuentaExento = " & lIdCtaExento
         Q1 = Q1 & ", Afecto = " & vFld(Rs("ndTotNeto"))
         Q1 = Q1 & ", IdCuentaAfecto = " & lIdCtaAfecto
         Q1 = Q1 & ", IVA = " & vFld(Rs("ndIVA"))
         Q1 = Q1 & ", IdCuentaIVA = " & lIdCtaIVA
         Q1 = Q1 & ", OtroImp = " & vFld(Rs("ndTotOtrosImpto"))
         Q1 = Q1 & ", IdCuentaOtroImp = " & lIdCtaOtrosImp
         Q1 = Q1 & ", Total = " & vFld(Rs("ndTotDocumen"))
         Q1 = Q1 & ", IdCuentaTotal = " & lIdCtaTotal
         Q1 = Q1 & ", FEmisionOri = " & CLng(Int(vFld(Rs("dtFechaDoc"))))
         Q1 = Q1 & ", CorrInterno = " & i
         Q1 = Q1 & ", SaldoDoc = NULL"
         Q1 = Q1 & ", DTE = " & vFld(Rs("bDocE"))
         Q1 = Q1 & ", PorcentRetencion = 0"
         Q1 = Q1 & ", TipoRetencion = 0"
         Q1 = Q1 & ", IdSucursal = " & IdSucursal
         Q1 = Q1 & ", FExported = 0"
         Q1 = Q1 & ", OldIdDoc = 0"
         Q1 = Q1 & ", MovEdited = 1"
         Q1 = Q1 & ", OtrosVal = " & vFld(Rs("ndTotDocumen")) - (vFld(Rs("ndTotNeto")) + vFld(Rs("ndTotExento")) + vFld(Rs("ndIVA")) + vFld(Rs("ndTotOtrosImpto")))
         Q1 = Q1 & ", FImporF29 = " & CLng(Int(Now))
         
         Q1 = Q1 & " WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
         Call ExecSQL(DbMain, Q1)
         
         'insertamos los MovDocumento
         orden = 1
   
         'Afecto
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndNetoRealizable")), lIdCtaAfecto, "Neto Realizable", LIBCOMPRAS_AFECTO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndNetoActFijo")), lIdCtaActFijoAfecto, "Neto Activo Fijo", LIBCOMPRAS_AFECTO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndNetoGastos")), lIdCtaAfecto, "Neto Gastos", LIBCOMPRAS_AFECTO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndNetoOtros")), lIdCtaAfecto, "Neto Otros", LIBCOMPRAS_AFECTO, False, False, orden)
         
         'Exento
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndExenRealizable")), lIdCtaExento, "Exento Realizable", LIBCOMPRAS_EXENTO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndExenActFijo")), lIdCtaActFijoExento, "Exento Activo Fijo", LIBCOMPRAS_EXENTO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndExenGastos")), lIdCtaExento, "Exento Gastos", LIBCOMPRAS_EXENTO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndExenOtros")), lIdCtaExento, "Exento Otros", LIBCOMPRAS_EXENTO, False, False, orden)
         
         'IVA
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaRealizable")), lIdCtaIVA, "IVA Créd. Fisc. Realizable", LIBCOMPRAS_IVACREDFISC, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaActFijo")), lIdCtaIVA, "IVA Créd. Fisc. Act. Fijo", LIBCOMPRAS_IVAACTFIJO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaGastos")), lIdCtaIVA, "IVA Créd. Fisc. Gastos", LIBCOMPRAS_IVACREDFISC, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaOtros")), lIdCtaIVA, "IVA Créd. Fisc. Otros", LIBCOMPRAS_IVACREDFISC, False, False, orden)
         
         'Otros impuestos
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c28")), lIdCtaOtrosImp, "Otros Imp., Art. Suntuarios", LIBCOMPRAS_OTROSIMP, False, False, orden)
'         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c576")), lIdCtaOtrosImp, "Otros Imp., Piscos, Licores, Whisky y Aguardiente", LIBCOMPRAS_OTROSIMP, False, False, Orden)
'         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c33")), lIdCtaOtrosImp, "Otros Imp., Vinos, Champ. y Otros", LIBCOMPRAS_OTROSIMP, False, False, Orden)
'         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c149")), lIdCtaOtrosImp, "Otros Imp., Cervezas", LIBCOMPRAS_OTROSIMP, False, False, Orden)
'         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c85")), lIdCtaOtrosImp, "Otros Imp., Bebidas Analcohólicas", LIBCOMPRAS_OTROSIMP, False, False, Orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c127")), lIdCtaOtrosImp, "Otros Imp., Impuesto Específico, General", LIBCOMPRAS_OTROSIMP, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c544")), lIdCtaOtrosImp, "Otros Imp., Impuesto Específico, Transporte", LIBCOMPRAS_OTROSIMP, False, False, orden)
         
         'PS Otros impuestos
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c575")), lIdCtaOtrosImp, "Otros Imp., Piscos, Licores, Whisky y Aguardiente", LIBCOMPRAS_IMPPISCO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c574")), lIdCtaOtrosImp, "Otros Imp., Vinos, Champ. y Otros", LIBCOMPRAS_IMPVINOS, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c580")), lIdCtaOtrosImp, "Otros Imp., Cervezas", LIBCOMPRAS_IMPCERVEZA, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c582")), lIdCtaOtrosImp, "Otros Imp., Bebidas Analcohólicas", LIBCOMPRAS_IMPBEBANALC, False, False, orden)
                 
        'PS IVA retenido
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaRetTot")), lIdCtaOtrosImp, "IVA Total Retenido a Terceras Personas", LIBCOMPRAS_IVARETTOT, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaRetPar")), lIdCtaOtrosImp, "IVA Parcial Retenido a Terceras Personas", LIBCOMPRAS_IVARETPARC, False, False, orden)
         
         'Anticipos: Harina, Carne
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("nd12Harina")), lIdCtaOtrosImp, "Anticipos: 12% Harina", LIBCOMPRAS_IVAANTICIPHARINA, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("nd4Carne")), lIdCtaOtrosImp, "Anticipos: 5% Carne", LIBCOMPRAS_IVAANTICIPCARNE, False, False, orden)
         
         'IVA Irrecuperable
         
         If lIdCtaIVAIrrec = 0 And (vFld(Rs("ndIvaIrreRealizable")) <> 0 Or vFld(Rs("ndIvaIrreActFijo")) <> 0 Or vFld(Rs("ndIvaIrreGastos")) <> 0 Or vFld(Rs("ndIvaIrreOtros")) <> 0) Then
            Call GetCtaIVAIrrec
         End If
        
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaIrreRealizable")), lIdCtaIVAIrrec, "IVA Irrecuperable Realizable", LIBCOMPRAS_IVAIRREC, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaIrreActFijo")), lIdCtaIVAIrrec, "IVA Irrecuperable Act. Fijo", LIBCOMPRAS_IVAIRRACTFIJO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaIrreGastos")), lIdCtaIVAIrrec, "IVA Irrecuperable Gastos", LIBCOMPRAS_IVAIRREC, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaIrreOtros")), lIdCtaIVAIrrec, "IVA Irrecuperable Otros", LIBCOMPRAS_IVAIRREC, False, False, orden)
        
         'Total
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndTotDocumen")), lIdCtaTotal, "Total Documento", LIBCOMPRAS_TOTAL, True, False, orden)
     
         i = i + 1
      
      End If
      Call CloseRs(RsOldDoc)
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   ImportF29Compras = NDocsImp

End Function

Private Function ImportF29VentasNac(ByVal Mes As Integer) As Long
   Dim Rs As Recordset
   Dim RsDoc As Recordset
   Dim Q1 As String
   Dim IdDoc As Long
   Dim i As Integer, j As Integer
   Dim TipoDoc As Integer
   Dim orden As Integer
   Dim TipoLib As Integer
   Dim NDocsImp As Long
   Dim RsOldDoc As Recordset
   Dim IdSucursal As Long
   Dim DTE As Integer
   Dim FldArray(3) As AdvTbAddNew_t
   
   TipoLib = LIB_VENTAS

   'Facturas de Venta Nacionales: tabla LAU_VentasFacturaNac
   
   Q1 = "SELECT LAU_VentasFacturaNac.sRUT, nbTipoDoc, nlNumDocTxt, dtFechaDoc, nlNumDocRef, sCodDatoUsuario, "
   Q1 = Q1 & " ndTotNeto, ndTotExento, ndIVA, ndTotOtrosImpto, ndRebaja65, nd12Harina, nd5Carne, "
   Q1 = Q1 & " ndIVAIrrecuperable, ndIVARetenido, ndTotDocumen, bDocE, nMesWrk, DELGIRO, "
   Q1 = Q1 & " ndImpto_c522, ndImpto_c526, ndImpto_c113, ndImpto_c577, ndImpto_c32, ndImpto_c150, ndImpto_c146, "
   Q1 = Q1 & " LAU_mPersonas.sNombre As NombrePers, LAU_VentasFacturaNac.nlNumItemSuc, LAU_Sucursales.sCodSuc "
   Q1 = Q1 & " FROM (LAU_VentasFacturaNac "
   Q1 = Q1 & " LEFT JOIN LAU_mPersonas ON LAU_VentasFacturaNac.sRUT = LAU_mPersonas.sRUT) "
   Q1 = Q1 & " LEFT JOIN LAU_Sucursales ON LAU_VentasFacturaNac.nlNumItemSuc = LAU_Sucursales.nlNumItem  "
   Q1 = Q1 & " WHERE bNula = 0 AND nMesWrk = " & Mes
   Q1 = Q1 & " ORDER BY LAU_VentasFacturaNac.nlNumItem"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 1
   lIdCtaIVAIrrec = 0
   NDocsImp = 0
   
   Do While Rs.EOF = False
  
      TipoDoc = GetTipoDocFromLAU(TipoLib, vFld(Rs("nbTipoDoc")))
      
      'obtenemos la sucursal si el número en HR coincide con el código en Contabilidad
      IdSucursal = 0
      If vFld(Rs("sCodSuc")) <> "" Then
         IdSucursal = GetIdSucursal(vFld(Rs("sCodSuc")))
      End If
      
      DTE = IIf(vFld(Rs("bDocE")) <> 0, -1, 0)
     
      'vemos si el documento ya está en el libro
                  
      Q1 = "SELECT IdDoc, FEmision FROM Documento WHERE TipoLib=" & LIB_VENTAS & " AND TipoDoc=" & TipoDoc & " AND NumDoc='" & vFld(Rs("nlNumDocTxt")) & "'" & " AND DTE = " & DTE
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set RsOldDoc = OpenRs(DbMain, Q1)
      
      If RsOldDoc.EOF = True Then  'no existe, lo agregamos
      
'         Set RsDoc = DbMain.OpenRecordset("Documento")
'         RsDoc.AddNew
'
'         IdDoc = vFld(RsDoc("IdDoc"))
'         RsDoc.Fields("IdUsuario") = gUsuario.IdUsuario
'         RsDoc.Fields("FechaCreacion") = CLng(Int(Now))
'
'         On Error Resume Next
'
'         For j = 1 To 10
'            RsDoc.Fields("NumDoc") = CLng(Rnd * 321654356#)
'
'            RsDoc.Update
'
'            If Err = 0 Then
'               Exit For
'            End If
'         Next j
'
'         On Error GoTo 0
'
'         RsDoc.Close
'
'         Set RsDoc = Nothing
         
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
                  
         NDocsImp = NDocsImp + 1
            
         'datos documento
         
         Q1 = "UPDATE Documento SET "
         Q1 = Q1 & "  TipoLib = " & TipoLib
         Q1 = Q1 & ", TipoDoc = " & TipoDoc
         Q1 = Q1 & ", NumDoc = '" & Trim(vFld(Rs("nlNumDocTxt"))) & "'"
         Q1 = Q1 & ", NumDocHasta = '0'"
         Q1 = Q1 & ", IdEntidad = " & GetInsertEntidad(vFld(Rs("sRUT"), True), vFld(Rs("NombrePers"), True), ENT_PROVEEDOR)
         Q1 = Q1 & ", TipoEntidad = " & ENT_CLIENTE
         'Q1 = Q1 & ", FEmision = " & CLng(Int(vFld(Rs("dtFechaDoc"))))
         Q1 = Q1 & ", FEmision = " & CLng(Int(DateSerial(gEmpresa.Ano, vFld(Rs("nMesWrk")), 1)))
         Q1 = Q1 & ", FVenc = 0"
         Q1 = Q1 & ", NumDocRef = '" & vFld(Rs("nlNumDocRef")) & "'"
         Q1 = Q1 & ", Descrip = '" & vFld(Rs("sCodDatoUsuario"), True) & "'"
         Q1 = Q1 & ", Estado = " & ED_PENDIENTE
         Q1 = Q1 & ", Exento = " & vFld(Rs("ndTotExento"))
         Q1 = Q1 & ", IdCuentaExento = " & lIdCtaExento
         Q1 = Q1 & ", Afecto = " & vFld(Rs("ndTotNeto"))
         Q1 = Q1 & ", IdCuentaAfecto = " & lIdCtaAfecto
         Q1 = Q1 & ", IVA = " & vFld(Rs("ndIVA"))
         Q1 = Q1 & ", IdCuentaIVA = " & lIdCtaIVA
         Q1 = Q1 & ", OtroImp = " & vFld(Rs("ndTotOtrosImpto"))
         Q1 = Q1 & ", IdCuentaOtroImp = " & lIdCtaOtrosImp
         Q1 = Q1 & ", Total = " & vFld(Rs("ndTotDocumen"))
         Q1 = Q1 & ", IdCuentaTotal = " & lIdCtaTotal
         Q1 = Q1 & ", FEmisionOri = " & CLng(Int(vFld(Rs("dtFechaDoc"))))
         Q1 = Q1 & ", CorrInterno = " & i
         Q1 = Q1 & ", SaldoDoc = NULL"
         Q1 = Q1 & ", DTE = " & vFld(Rs("bDocE"))
         Q1 = Q1 & ", PorcentRetencion = 0"
         Q1 = Q1 & ", TipoRetencion = 0"
         Q1 = Q1 & ", IdSucursal = " & IdSucursal
         Q1 = Q1 & ", FExported = 0"
         Q1 = Q1 & ", OldIdDoc = 0"
         Q1 = Q1 & ", MovEdited = 1"
         Q1 = Q1 & ", OtrosVal = " & vFld(Rs("ndTotDocumen")) - (vFld(Rs("ndTotNeto")) + vFld(Rs("ndTotExento")) + vFld(Rs("ndIVA")) + vFld(Rs("ndTotOtrosImpto")))
         Q1 = Q1 & ", FImporF29 = " & CLng(Int(Now))
         Q1 = Q1 & ", Giro = " & IIf(vFld(Rs("DELGIRO")) = 0, 1, 0)   'es al revés
         
         Q1 = Q1 & " WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
         Call ExecSQL(DbMain, Q1)
         
         'insertamos los MovDocumento
         orden = 1
   
         'Afecto
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndTotNeto")), lIdCtaAfecto, "Total Neto", LIBVENTAS_AFECTO, False, False, orden)
         
         'Exento
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndTotExento")), lIdCtaExento, "Total Exento", LIBVENTAS_EXENTO, False, False, orden)
         
         'IVA
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIva")), lIdCtaIVA, "IVA Débito Fiscal", LIBVENTAS_IVADEBFISC, False, False, orden)
         
         'Otros impuestos
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c522")), lIdCtaOtrosImp, "Otros Imp., Art. Suntuarios, L51 (15%)", LIBVENTAS_OTROSIMP, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c526")), lIdCtaOtrosImp, "Otros Imp., Art. Suntuarios, L52 (50%)", LIBVENTAS_OTROSIMP, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c113")), lIdCtaOtrosImp, "Otros Imp., Art. Suntuarios, L53 (15%)", LIBVENTAS_OTROSIMP, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c577")), lIdCtaOtrosImp, "Otros Imp., Piscos, Licores, Whisky y Aguardiente", LIBVENTAS_IMPPISCO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c32")), lIdCtaOtrosImp, "Otros Imp., Vinos, Champ. y Otros", LIBVENTAS_IMPVINOS, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c150")), lIdCtaOtrosImp, "Otros Imp., Cervezas", LIBVENTAS_IMPCERVEZA, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c146")), lIdCtaOtrosImp, "Otros Imp., Bebidas Analcohólicas", LIBVENTAS_IMPBEBANHALC, False, False, orden)
         
         'IVA retenido
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIVARetenido")), lIdCtaOtrosImp, "IVA Total Retenido Terceras Personas", LIBVENTAS_IVARETTOT, False, False, orden)
         
         'Anticipos: Harina, Carne
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("nd12Harina")), lIdCtaOtrosImp, "Anticipos: 12% Harina", LIBVENTAS_IVAANTICIPADOHARINA, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("nd5Carne")), lIdCtaOtrosImp, "Anticipos: 5% Carne", LIBVENTAS_RETANTCAMBIOSUJCARNE, False, False, orden)
         
         'IVA Rebaja 65%
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndRebaja65")), lIdCtaOtrosImp, "Rebaja 65% empresas contructoras", LIBVENTAS_REBAJA65, False, False, orden)
         
         'IVA Irrecuperable
         If lIdCtaIVAIrrec = 0 And vFld(Rs("ndIvaIrrecuperable")) <> 0 Then
            Call GetCtaIVAIrrec
         End If
        
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaIrrecuperable")), lIdCtaIVAIrrec, "IVA Irrecuperable", LIBVENTAS_OTROSIMP, False, False, orden)
        
         'Total
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndTotDocumen")), lIdCtaTotal, "Total Documento", LIBVENTAS_TOTAL, True, False, orden)
     
         i = i + 1
      Else
         Call MsgBox1("No se importará el documento " & vbNewLine & vbNewLine & GetNombreTipoDoc(LIB_VENTAS, TipoDoc) & " N° " & Trim(vFld(Rs("nlNumDocTxt"))) & vbNewLine & vbNewLine & " porque ya fue ingresado un documento de este mismo tipo y nombre en el mes de " & Format(vFld(RsOldDoc("FEmision")), "mmm yyyy"), vbInformation)
         
      End If
      
      Call CloseRs(RsOldDoc)
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   ImportF29VentasNac = NDocsImp

End Function

Private Function ImportF29VentasExp(ByVal Mes As Integer) As Long
   Dim Rs As Recordset
   Dim RsDoc As Recordset
   Dim Q1 As String
   Dim IdDoc As Long
   Dim i As Integer, j As Integer
   Dim TipoDoc As Integer
   Dim orden As Integer
   Dim TipoLib As Integer
   Dim LauTipoDoc As Integer
   Dim NDocsImp As Long
   Dim RsOldDoc As Recordset
   Dim IdSucursal As Long
   Dim FldArray(3) As AdvTbAddNew_t
   
   TipoLib = LIB_VENTAS

   'Facturas de Venta Exportación: tabla LAU_VentasFacturaExp
   
   Q1 = "SELECT LAU_VentasFacturaExp.nlNumItemRefExp, nbTipoDoc, nlNumDocTxt, dtFechaDoc, nlNumDocRef, sCodDatoUsuario, "
   Q1 = Q1 & " ndMontoPesos, nMesWrk, "
   Q1 = Q1 & " LAU_ReferenciaExp.sNombre As NombreRef, LAU_ReferenciaExp.sNumRef, LAU_VentasFacturaExp.nlNumItemSuc, LAU_Sucursales.sCodSuc  "
   Q1 = Q1 & " FROM (LAU_VentasFacturaExp "
   Q1 = Q1 & " LEFT JOIN LAU_ReferenciaExp ON LAU_VentasFacturaExp.nlNumItemRefExp = LAU_ReferenciaExp.nlNumItem) "
   Q1 = Q1 & " LEFT JOIN LAU_Sucursales ON LAU_VentasFacturaExp.nlNumItemSuc = LAU_Sucursales.nlNumItem  "
   Q1 = Q1 & " WHERE bNula = 0 AND nMesWrk = " & Mes
   Q1 = Q1 & " ORDER BY LAU_VentasFacturaExp.nlNumItem"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 1
   lIdCtaIVAIrrec = 0
   NDocsImp = 0
   
   Do While Rs.EOF = False
   
      Select Case vFld(Rs("nbTipoDoc"))
      
         Case 0
            LauTipoDoc = LAU_VENTA_FACTEXP
            
         Case 1
            LauTipoDoc = LAU_VENTA_NDEBEXP
         
         Case 2
            LauTipoDoc = LAU_VENTA_NCREDEXP
         
         Case 3
            LauTipoDoc = LAU_VENTA_OTRO
            
      End Select
      
      TipoDoc = GetTipoDocFromLAU(TipoLib, LauTipoDoc)
      
      'obtenemos la sucursal si el número en HR coincide con el código en Contabilidad
      IdSucursal = 0
      If vFld(Rs("sCodSuc")) <> "" Then
         IdSucursal = GetIdSucursal(vFld(Rs("sCodSuc")))
      End If
            
      'vemos si el documento ya está en el libro
      
      Q1 = "SELECT IdDoc FROM Documento WHERE TipoLib=" & LIB_VENTAS & " AND TipoDoc=" & TipoDoc & " AND NumDoc='" & vFld(Rs("nlNumDocTxt")) & "'"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set RsOldDoc = OpenRs(DbMain, Q1)
      
      If RsOldDoc.EOF = True Then  'no existe, lo agregamos
   
'         Set RsDoc = DbMain.OpenRecordset("Documento")
'         RsDoc.AddNew
'
'         IdDoc = vFld(RsDoc("IdDoc"))
'         RsDoc.Fields("IdUsuario") = gUsuario.IdUsuario
'         RsDoc.Fields("FechaCreacion") = CLng(Int(Now))
'
'         On Error Resume Next
'
'         For j = 1 To 10
'            RsDoc.Fields("NumDoc") = CLng(Rnd * 321654356#)
'
'            RsDoc.Update
'
'            If Err = 0 Then
'               Exit For
'            End If
'         Next j
'
'         On Error GoTo 0
'
'         RsDoc.Close
'
'         Set RsDoc = Nothing
         
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
         
         
         NDocsImp = NDocsImp + 1
   
         'datos documento
                  
         Q1 = "UPDATE Documento SET "
         Q1 = Q1 & "  TipoLib = " & TipoLib
         Q1 = Q1 & ", TipoDoc = " & TipoDoc
         Q1 = Q1 & ", NumDoc = '" & vFld(Rs("nlNumDocTxt")) & "'"
         Q1 = Q1 & ", NumDocHasta = '0'"
         Q1 = Q1 & ", IdEntidad = " & GetInsertEntidad(vFld(Rs("sNumRef"), True), vFld(Rs("NombreRef"), True), ENT_PROVEEDOR, True)
         Q1 = Q1 & ", TipoEntidad = " & ENT_CLIENTE
         'Q1 = Q1 & ", FEmision = " & CLng(Int(vFld(Rs("dtFechaDoc"))))
         Q1 = Q1 & ", FEmision = " & CLng(Int(DateSerial(gEmpresa.Ano, vFld(Rs("nMesWrk")), 1)))
         Q1 = Q1 & ", FVenc = 0"
         Q1 = Q1 & ", NumDocRef = '" & vFld(Rs("nlNumDocRef")) & "'"
         Q1 = Q1 & ", Descrip = '" & vFld(Rs("sCodDatoUsuario"), True) & "'"
         Q1 = Q1 & ", Estado = " & ED_PENDIENTE
         Q1 = Q1 & ", Exento = " & vFld(Rs("ndMontoPesos"))
         Q1 = Q1 & ", IdCuentaExento = " & lIdCtaExento
         Q1 = Q1 & ", Afecto = 0"
         Q1 = Q1 & ", IdCuentaAfecto = 0"
         Q1 = Q1 & ", IVA = 0"
         Q1 = Q1 & ", IdCuentaIVA = 0"
         Q1 = Q1 & ", OtroImp = 0"
         Q1 = Q1 & ", IdCuentaOtroImp = 0"
         Q1 = Q1 & ", Total = " & vFld(Rs("ndMontoPesos"))
         Q1 = Q1 & ", IdCuentaTotal = " & lIdCtaTotal
         Q1 = Q1 & ", FEmisionOri = " & CLng(Int(vFld(Rs("dtFechaDoc"))))
         Q1 = Q1 & ", CorrInterno = " & i
         Q1 = Q1 & ", SaldoDoc = NULL"
         Q1 = Q1 & ", DTE = 0"
         Q1 = Q1 & ", PorcentRetencion = 0"
         Q1 = Q1 & ", TipoRetencion = 0"
         Q1 = Q1 & ", IdSucursal = " & IdSucursal
         Q1 = Q1 & ", FExported = 0"
         Q1 = Q1 & ", OldIdDoc = 0"
         Q1 = Q1 & ", MovEdited = 1"
         Q1 = Q1 & ", OtrosVal = 0"
         Q1 = Q1 & ", FImporF29 = " & CLng(Int(Now))
         
         Q1 = Q1 & " WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
         Call ExecSQL(DbMain, Q1)
         
         'insertamos los MovDocumento
         orden = 1
         
         'Exento
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndMontoPesos")), lIdCtaExento, "Total Exento", LIBVENTAS_EXENTO, False, False, orden)
              
         'Total
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndMontoPesos")), lIdCtaTotal, "Total Documento", LIBVENTAS_TOTAL, True, False, orden)
     
         i = i + 1
      End If
      
      Call CloseRs(RsOldDoc)
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   ImportF29VentasExp = NDocsImp

End Function

Private Function ImportF29VentasBoleta(ByVal Mes As Integer) As Long
   Dim Rs As Recordset
   Dim RsDoc As Recordset
   Dim Q1 As String
   Dim IdDoc As Long
   Dim i As Integer, j As Integer
   Dim TipoDoc As Integer
   Dim orden As Integer
   Dim TipoLib As Integer
   Dim Tot As Double
   Dim LauTipoDoc As Integer
   Dim Afecto As Double
   Dim NDocsImp As Long
   Dim RsOldDoc As Recordset
   Dim IdSucursal As Long
   Dim DTE As Integer
   Dim FldArray(3) As AdvTbAddNew_t

   
   TipoLib = LIB_VENTAS

   'Venta con Boleta: tabla LAU_VentasBoletas
   
   Q1 = "SELECT nbTipoVta, nlBoletaIni, nlBoletaFin, dtFechaDoc, sCodDatoUsuario, "
   Q1 = Q1 & " ndTotNeto, ndTotExento, ndIVA, ndTotOtrosImpto, "
   Q1 = Q1 & " bDocE, nMesWrk, LAU_VentasBoletas.nlNumItemSuc, LAU_Sucursales.sCodSuc, "
   Q1 = Q1 & " ndImpto_c113, ndImpto_c577, ndImpto_c32, ndImpto_c150, ndImpto_c146, "
   Q1 = Q1 & " sMaqVCPME, ndCantVCPME "    'para mVPE
   Q1 = Q1 & " FROM LAU_VentasBoletas "
   Q1 = Q1 & " LEFT JOIN LAU_Sucursales ON LAU_VentasBoletas.nlNumItemSuc = LAU_Sucursales.nlNumItem  "
   Q1 = Q1 & " WHERE bNula = 0 AND nMesWrk = " & Mes
   Q1 = Q1 & " ORDER BY LAU_VentasBoletas.nlNumItem"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 1
   lIdCtaIVAIrrec = 0
   NDocsImp = 0
   
   Do While Rs.EOF = False
   
      Select Case vFld(Rs("nbTipoVta"))
         
         Case 0
            LauTipoDoc = LAU_VENTA_BOLETA
            
         Case 1
            LauTipoDoc = LAU_VENTA_VTAMENOR
            
         Case 2
            LauTipoDoc = LAU_VENTA_BOLEXENTA
            
         Case 3
            LauTipoDoc = LAU_VENTA_VALEPAGOELECTR
          '2814014 pipe
         Case 4
            LauTipoDoc = LAU_VENTA_BOLVENTAEXENTA
          'fin 2814014
      End Select
      
      TipoDoc = GetTipoDocFromLAU(TipoLib, LauTipoDoc)
      
      'obtenemos la sucursal si el número en HR coincide con el código en Contabilidad
      IdSucursal = 0
      If vFld(Rs("sCodSuc")) <> "" Then
         IdSucursal = GetIdSucursal(vFld(Rs("sCodSuc")))
      End If
   
      DTE = IIf(vFld(Rs("bDocE")) <> 0, -1, 0)
           
      'vemos si el documento ya está en el libro
      
      Q1 = "SELECT IdDoc FROM Documento WHERE TipoLib=" & LIB_VENTAS & " AND TipoDoc=" & TipoDoc & " AND NumDoc='" & vFld(Rs("nlBoletaIni")) & "' AND NumDocHasta='" & vFld(Rs("nlBoletaFin")) & "'" & " AND DTE = " & DTE
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set RsOldDoc = OpenRs(DbMain, Q1)
      
      If RsOldDoc.EOF = True Then  'no existe, lo agregamos
      
'         Set RsDoc = DbMain.OpenRecordset("Documento")
'         RsDoc.AddNew
'
'         IdDoc = vFld(RsDoc("IdDoc"))
'         RsDoc.Fields("IdUsuario") = gUsuario.IdUsuario
'         RsDoc.Fields("FechaCreacion") = CLng(Int(Now))
'
'         On Error Resume Next
'
'         For j = 1 To 10
'            RsDoc.Fields("NumDoc") = CLng(Rnd * 321654356#)
'
'            RsDoc.Update
'
'            If Err = 0 Then
'               Exit For
'            End If
'         Next j
'
'         On Error GoTo 0
'
'         RsDoc.Close
'
'         Set RsDoc = Nothing

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
         
         NDocsImp = NDocsImp + 1
     
         'datos documento
                  
         Q1 = "UPDATE Documento SET "
         Q1 = Q1 & "  TipoLib = " & TipoLib
         Q1 = Q1 & ", TipoDoc = " & TipoDoc
         
         If LauTipoDoc = LAU_VENTA_VALEPAGOELECTR Then
            Q1 = Q1 & ", NumDoc = '" & vFld(Rs("sMaqVCPME")) & "'"
            Q1 = Q1 & ", CantBoletas = '" & vFld(Rs("ndCantVCPME")) & "'"
         Else
            Q1 = Q1 & ", NumDoc = '" & vFld(Rs("nlBoletaIni")) & "'"
         End If
         
         Q1 = Q1 & ", NumDocHasta = '" & vFld(Rs("nlBoletaFin")) & "'"
            
         Q1 = Q1 & ", IdEntidad = 0"
         Q1 = Q1 & ", TipoEntidad = 0"
         'Q1 = Q1 & ", FEmision = " & CLng(Int(vFld(Rs("dtFechaDoc"))))
         Q1 = Q1 & ", FEmision = " & CLng(Int(DateSerial(gEmpresa.Ano, vFld(Rs("nMesWrk")), 1)))
         Q1 = Q1 & ", FVenc = 0"
         Q1 = Q1 & ", Descrip = '" & vFld(Rs("sCodDatoUsuario"), True) & "'"
         Q1 = Q1 & ", Estado = " & ED_PENDIENTE
         Q1 = Q1 & ", Exento = " & vFld(Rs("ndTotExento"))
         Q1 = Q1 & ", IdCuentaExento = " & lIdCtaExento
         
         Afecto = vFld(Rs("ndTotNeto")) - vFld(Rs("ndIVA"))
         Q1 = Q1 & ", Afecto = " & Afecto
         Q1 = Q1 & ", IdCuentaAfecto = " & lIdCtaAfecto
         Q1 = Q1 & ", IVA = " & vFld(Rs("ndIVA"))
         Q1 = Q1 & ", IdCuentaIVA = " & lIdCtaIVA
         Q1 = Q1 & ", OtroImp = " & vFld(Rs("ndTotOtrosImpto"))
         Q1 = Q1 & ", IdCuentaOtroImp = " & lIdCtaOtrosImp
         
         Tot = vFld(Rs("ndTotExento")) + Afecto + vFld(Rs("ndIVA")) + vFld(Rs("ndTotOtrosImpto"))
         Q1 = Q1 & ", Total = " & Tot
         
         Q1 = Q1 & ", IdCuentaTotal = " & lIdCtaTotal
         Q1 = Q1 & ", FEmisionOri = " & CLng(Int(vFld(Rs("dtFechaDoc"))))
         Q1 = Q1 & ", CorrInterno = " & i
         Q1 = Q1 & ", SaldoDoc = NULL"
         Q1 = Q1 & ", DTE = " & vFld(Rs("bDocE"))
         Q1 = Q1 & ", PorcentRetencion = 0"
         Q1 = Q1 & ", TipoRetencion = 0"
         Q1 = Q1 & ", IdSucursal = " & IdSucursal
         Q1 = Q1 & ", FExported = 0"
         Q1 = Q1 & ", OldIdDoc = 0"
         Q1 = Q1 & ", MovEdited = 1"
         Q1 = Q1 & ", FImporF29 = " & CLng(Int(Now))
         
         Q1 = Q1 & " WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
         Call ExecSQL(DbMain, Q1)
         
         'insertamos los MovDocumento
         orden = 1
   
         'Afecto
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, Afecto, lIdCtaAfecto, "Total Neto", LIBVENTAS_AFECTO, False, False, orden)
         
         'Exento
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndTotExento")), lIdCtaExento, "Total Exento", LIBVENTAS_EXENTO, False, False, orden)
         
         'IVA
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIva")), lIdCtaIVA, "IVA Débito Fiscal", LIBVENTAS_IVADEBFISC, False, False, orden)
         
         'Otros impuestos
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c113")), lIdCtaOtrosImp, "Otros Imp., Art. Suntuarios, L53 (15%)", LIBVENTAS_OTROSIMP, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c577")), lIdCtaOtrosImp, "Otros Imp., Piscos, Licores, Whisky y Aguardiente", LIBVENTAS_IMPPISCO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c32")), lIdCtaOtrosImp, "Otros Imp., Vinos, Champ. y Otros", LIBVENTAS_IMPVINOS, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c150")), lIdCtaOtrosImp, "Otros Imp., Cervezas", LIBVENTAS_IMPCERVEZA, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c146")), lIdCtaOtrosImp, "Otros Imp., Bebidas Analcohólicas", LIBVENTAS_IMPBEBANHALC, False, False, orden)
                    
         'Total
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, Tot, lIdCtaTotal, "Total Documento", LIBVENTAS_TOTAL, True, False, orden)
     
         i = i + 1
         
      End If
      
      Call CloseRs(RsOldDoc)
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   ImportF29VentasBoleta = NDocsImp

End Function

Private Function ImportF29VentasDevBoleta(ByVal Mes As Integer) As Long
   Dim Rs As Recordset
   Dim RsDoc As Recordset
   Dim Q1 As String
   Dim IdDoc As Long
   Dim i As Integer, j As Integer
   Dim TipoDoc As Integer
   Dim orden As Integer
   Dim TipoLib As Integer
   Dim Tot As Double
   Dim LauTipoDoc As Integer
   Dim Afecto As Double
   Dim NDocsImp As Long
   Dim RsOldDoc As Recordset
   Dim IdSucursal As Long
   Dim DTE As Integer
   Dim FldArray(3) As AdvTbAddNew_t

   
   TipoLib = LIB_VENTAS

   'Devoluciones de Ventas con Boleta: tabla LAU_VentasDevolucion
   
   Q1 = "SELECT nbTipoVta, nlNumBoleta, dtFechaDoc, sCodDatoUsuario, "
   Q1 = Q1 & " ndTotNeto, ndTotExento, ndIVA, ndTotOtrosImpto, ndIvaIrrecuperable, "
   Q1 = Q1 & " bDocE, nMesWrk, LAU_VentasDevolucion.nlNumItemSuc, LAU_Sucursales.sCodSuc, "
   Q1 = Q1 & " ndImpto_c113, ndImpto_c577, ndImpto_c32, ndImpto_c150, ndImpto_c146 "
   Q1 = Q1 & " FROM LAU_VentasDevolucion "
   Q1 = Q1 & " LEFT JOIN LAU_Sucursales ON LAU_VentasDevolucion.nlNumItemSuc = LAU_Sucursales.nlNumItem  "
   Q1 = Q1 & " WHERE bNula = 0 AND nMesWrk = " & Mes
   Q1 = Q1 & " ORDER BY LAU_VentasDevolucion.nlNumItem"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 1
   lIdCtaIVAIrrec = 0
   NDocsImp = 0
   
   Do While Rs.EOF = False
   
      Select Case vFld(Rs("nbTipoVta"))
         
         Case 0
            LauTipoDoc = LAU_VENTA_DEVBOLETA
            
         Case Else
            LauTipoDoc = LAU_VENTA_DEVBOLETA
                        
      End Select
      
      TipoDoc = GetTipoDocFromLAU(TipoLib, LauTipoDoc)
      
      'obtenemos la sucursal si el número en HR coincide con el código en Contabilidad
      IdSucursal = 0
      If vFld(Rs("sCodSuc")) <> "" Then
         IdSucursal = GetIdSucursal(vFld(Rs("sCodSuc")))
      End If
   
      DTE = IIf(vFld(Rs("bDocE")) <> 0, -1, 0)
                 
     'vemos si el documento ya está en el libro
      
      Q1 = "SELECT IdDoc FROM Documento WHERE TipoLib=" & LIB_VENTAS & " AND TipoDoc=" & TipoDoc & " AND NumDoc='" & vFld(Rs("nlNumBoleta")) & "'" & " AND DTE = " & DTE
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set RsOldDoc = OpenRs(DbMain, Q1)
      
      If RsOldDoc.EOF = True Then  'no existe, lo agregamos
      
'         Set RsDoc = DbMain.OpenRecordset("Documento")
'         RsDoc.AddNew
'
'         IdDoc = vFld(RsDoc("IdDoc"))
'         RsDoc.Fields("IdUsuario") = gUsuario.IdUsuario
'         RsDoc.Fields("FechaCreacion") = CLng(Int(Now))
'
'         On Error Resume Next
'
'         For j = 1 To 10
'            RsDoc.Fields("NumDoc") = CLng(Rnd * 321654356#)
'
'            RsDoc.Update
'
'            If Err = 0 Then
'               Exit For
'            End If
'         Next j
'
'         On Error GoTo 0
'
'         RsDoc.Close
         
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
         
         NDocsImp = NDocsImp + 1
   
         Set RsDoc = Nothing
      
         'datos documento
                  
         Q1 = "UPDATE Documento SET "
         Q1 = Q1 & "  TipoLib = " & TipoLib
         Q1 = Q1 & ", TipoDoc = " & TipoDoc
         Q1 = Q1 & ", NumDoc = '" & vFld(Rs("nlNumBoleta")) & "'"
         Q1 = Q1 & ", NumDocHasta = '0'"
         Q1 = Q1 & ", IdEntidad = 0"
         Q1 = Q1 & ", TipoEntidad = 0"
         'Q1 = Q1 & ", FEmision = " & CLng(Int(vFld(Rs("dtFechaDoc"))))
         Q1 = Q1 & ", FEmision = " & CLng(Int(DateSerial(gEmpresa.Ano, vFld(Rs("nMesWrk")), 1)))
         Q1 = Q1 & ", FVenc = 0"
         Q1 = Q1 & ", Descrip = '" & vFld(Rs("sCodDatoUsuario"), True) & "'"
         Q1 = Q1 & ", Estado = " & ED_PENDIENTE
         Q1 = Q1 & ", Exento = " & vFld(Rs("ndTotExento"))
         Q1 = Q1 & ", IdCuentaExento = " & lIdCtaExento
         
         Afecto = vFld(Rs("ndTotNeto")) - vFld(Rs("ndIVA"))
         Q1 = Q1 & ", Afecto = " & Afecto
         Q1 = Q1 & ", IdCuentaAfecto = " & lIdCtaAfecto
         Q1 = Q1 & ", IVA = " & vFld(Rs("ndIVA"))
         Q1 = Q1 & ", IdCuentaIVA = " & lIdCtaIVA
         Q1 = Q1 & ", OtroImp = " & vFld(Rs("ndTotOtrosImpto"))
         Q1 = Q1 & ", IdCuentaOtroImp = " & lIdCtaOtrosImp
         
         Tot = vFld(Rs("ndTotExento")) + Afecto + vFld(Rs("ndIVA")) + vFld(Rs("ndTotOtrosImpto")) + vFld(Rs("ndIvaIrrecuperable"))
         Q1 = Q1 & ", Total = " & Tot
         
         Q1 = Q1 & ", IdCuentaTotal = " & lIdCtaTotal
         Q1 = Q1 & ", FEmisionOri = " & CLng(Int(vFld(Rs("dtFechaDoc"))))
         Q1 = Q1 & ", CorrInterno = " & i
         Q1 = Q1 & ", SaldoDoc = NULL"
         Q1 = Q1 & ", DTE = " & vFld(Rs("bDocE"))
         Q1 = Q1 & ", PorcentRetencion = 0"
         Q1 = Q1 & ", TipoRetencion = 0"
         Q1 = Q1 & ", IdSucursal = " & IdSucursal
         Q1 = Q1 & ", FExported = 0"
         Q1 = Q1 & ", OldIdDoc = 0"
         Q1 = Q1 & ", MovEdited = 1"
         Q1 = Q1 & ", FImporF29 = " & CLng(Int(Now))
         
         Q1 = Q1 & " WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
         Call ExecSQL(DbMain, Q1)
         
         'insertamos los MovDocumento
         orden = 1
   
         'Afecto
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, Afecto, lIdCtaAfecto, "Total Neto", LIBVENTAS_AFECTO, False, False, orden)
         
         'Exento
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndTotExento")), lIdCtaExento, "Total Exento", LIBVENTAS_EXENTO, False, False, orden)
         
         'IVA
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIva")), lIdCtaIVA, "IVA Débito Fiscal", LIBVENTAS_IVADEBFISC, False, False, orden)
         
         'Otros impuestos
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c113")), lIdCtaOtrosImp, "Otros Imp., Art. Suntuarios, L53 (15%)", LIBVENTAS_OTROSIMP, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c577")), lIdCtaOtrosImp, "Otros Imp., Piscos, Licores, Whisky y Aguardiente", LIBVENTAS_IMPPISCO, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c32")), lIdCtaOtrosImp, "Otros Imp., Vinos, Champ. y Otros", LIBVENTAS_IMPVINOS, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c150")), lIdCtaOtrosImp, "Otros Imp., Cervezas", LIBVENTAS_IMPCERVEZA, False, False, orden)
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto_c146")), lIdCtaOtrosImp, "Otros Imp., Bebidas Analcohólicas", LIBVENTAS_IMPBEBANHALC, False, False, orden)
         
         'IVA Irrecuperable
         If lIdCtaIVAIrrec = 0 And vFld(Rs("ndIvaIrrecuperable")) <> 0 Then
            Call GetCtaIVAIrrec
         End If
        
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndIvaIrrecuperable")), lIdCtaIVAIrrec, "IVA Irrecuperable", LIBVENTAS_OTROSIMP, False, False, orden)
                    
         'Total
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, Tot, lIdCtaTotal, "Total Documento", LIBVENTAS_TOTAL, True, False, orden)
     
         i = i + 1
      End If
      
      Call CloseRs(RsOldDoc)
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   ImportF29VentasDevBoleta = NDocsImp

End Function
Private Function ImportF29Retenciones(ByVal Mes As Integer) As Long
   Dim Rs As Recordset
   Dim RsDoc As Recordset
   Dim Q1 As String
   Dim IdDoc As Long
   Dim i As Integer, j As Integer
   Dim TipoDoc As Integer
   Dim orden As Integer
   Dim TipoLib As Integer
   Dim PReten As Integer
   Dim TReten As Integer
   Dim DescTipo As String
   Dim NDocsImp As Long
   Dim RsOldDoc As Recordset
   Dim IdEnt As Long
   Dim IdSucursal As Long
   Dim FldArray(3) As AdvTbAddNew_t
   
   TipoLib = LIB_RETEN

   'Retenciones: tabla LAU_Retenciones
   
   Q1 = "SELECT LAU_Retenciones.sRUT, nlNumItemRefExp, nbTipo, nlNumDocTxt, dtFechaDoc, sCodDatoUsuario, "
   Q1 = Q1 & " ndBruto, ndImpto, ndCredito, ndNeto, nMesWrk, "
   Q1 = Q1 & " LAU_mPersonas.sNombre As NombrePers, "
   Q1 = Q1 & " LAU_ReferenciaExp.sNombre As NombreRef, LAU_ReferenciaExp.sNumRef, LAU_Retenciones.nlNumItemSuc, LAU_Sucursales.sCodSuc "
   Q1 = Q1 & " FROM ((LAU_Retenciones "
   Q1 = Q1 & " LEFT JOIN LAU_mPersonas ON LAU_Retenciones.sRUT = LAU_mPersonas.sRUT) "
   Q1 = Q1 & " LEFT JOIN LAU_ReferenciaExp ON LAU_Retenciones.nlNumItemRefExp = LAU_ReferenciaExp.nlNumItem) "
   Q1 = Q1 & " LEFT JOIN LAU_Sucursales ON LAU_Retenciones.nlNumItemSuc = LAU_Sucursales.nlNumItem  "
   Q1 = Q1 & " WHERE nMesWrk = " & Mes
   Q1 = Q1 & " ORDER BY LAU_Retenciones.nlNumItem"
   
   Set Rs = OpenRs(DbMain, Q1)
   
   i = 1
   NDocsImp = 0
   
   Do While Rs.EOF = False
   
      TipoDoc = TIPODOC_HONOR
      
      'obtenemos la sucursal si el número en HR coincide con el código en Contabilidad
      IdSucursal = 0
      If vFld(Rs("sCodSuc")) <> "" Then
         IdSucursal = GetIdSucursal(vFld(Rs("sCodSuc")))
      End If
      
       If vFld(Rs("nlNumItemRefExp")) <> 0 Then
          IdEnt = GetInsertEntidad(vFld(Rs("sNumRef"), True), vFld(Rs("NombreRef"), True), ENT_PROVEEDOR, True)
       Else
          IdEnt = GetInsertEntidad(vFld(Rs("sRUT"), True), vFld(Rs("NombrePers"), True), ENT_PROVEEDOR)
       End If
            
      'vemos si el documento ya está en el libro
      
      Q1 = "SELECT IdDoc FROM Documento WHERE TipoLib=" & LIB_RETEN & " AND TipoDoc=" & TipoDoc & " AND NumDoc='" & vFld(Rs("nlNumDocTxt")) & "' AND IdEntidad = " & IdEnt
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set RsOldDoc = OpenRs(DbMain, Q1)
      
      If RsOldDoc.EOF = True Then  'no existe, lo agregamos
      
'         Set RsDoc = DbMain.OpenRecordset("Documento")
'         RsDoc.AddNew
'
'         IdDoc = vFld(RsDoc("IdDoc"))
'         RsDoc.Fields("IdUsuario") = gUsuario.IdUsuario
'         RsDoc.Fields("FechaCreacion") = CLng(Int(Now))
'
'         On Error Resume Next
'
'         For j = 1 To 10
'            RsDoc.Fields("NumDoc") = CLng(Rnd * 321654356#)
'
'            RsDoc.Update
'
'            If Err = 0 Then
'               Exit For
'            End If
'         Next j
'
'         On Error GoTo 0
'
'         RsDoc.Close
'
'         Set RsDoc = Nothing

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
         
         NDocsImp = NDocsImp + 1
   
         'datos documento
                  
         Select Case vFld(Rs("nbTipo"))
         
            Case LAU_TRET_RENTASCAP
               PReten = IMPRET_OTRO
               TReten = TR_OTRO
               DescTipo = "Rentas Capitales Mobiliarios (sug. 17%)"
               
            Case LAU_TRET_HONORARIOS
               PReten = IMPRET_NAC
               TReten = TR_HONORARIOS
               DescTipo = "Honorarios (sug. 10%)"
               
            Case LAU_TRET_PARTDIR10
               PReten = IMPRET_NAC
               TReten = TR_DIETA
               DescTipo = "Participaciones a Directores (sug. 10%)"
               
            Case LAU_TRET_PARTDIR20
               PReten = IMPRET_EXT
               TReten = TR_DIETA
               DescTipo = "Participaciones a Directores (sug. 20%)"
                      
            Case LAU_TRET_RETMINEROS
               PReten = IMPRET_OTRO
               TReten = TR_OTRO
               DescTipo = "Reteción a Mineros"
                        
         End Select
         
         If vFld(Rs("sCodDatoUsuario"), True) <> "" Then
            DescTipo = vFld(Rs("sCodDatoUsuario"), True) & " - " & DescTipo
         End If
   
         Q1 = "UPDATE Documento SET "
         Q1 = Q1 & "  TipoLib = " & TipoLib
         Q1 = Q1 & ", TipoDoc = " & TipoDoc
         Q1 = Q1 & ", NumDoc = '" & vFld(Rs("nlNumDocTxt")) & "'"
         Q1 = Q1 & ", NumDocHasta = '0'"
         Q1 = Q1 & ", IdEntidad = " & IdEnt
         Q1 = Q1 & ", TipoEntidad = " & ENT_PROVEEDOR
         'Q1 = Q1 & ", FEmision = " & CLng(Int(vFld(Rs("dtFechaDoc"))))
         Q1 = Q1 & ", FEmision = " & CLng(Int(DateSerial(gEmpresa.Ano, vFld(Rs("nMesWrk")), 1)))
         Q1 = Q1 & ", FVenc = 0"
         Q1 = Q1 & ", NumDocRef = '0'"
         Q1 = Q1 & ", Descrip = '" & DescTipo & "'"
         Q1 = Q1 & ", Estado = " & ED_PENDIENTE
         
         If vFld(Rs("ndImpto")) = 0 Then
            Q1 = Q1 & ", Exento = " & vFld(Rs("ndBruto"))
            Q1 = Q1 & ", IdCuentaExento = " & lIdCtaHonSinRet
            Q1 = Q1 & ", Afecto = 0"
            Q1 = Q1 & ", IdCuentaAfecto = 0"
         Else
            Q1 = Q1 & ", Exento = 0"
            Q1 = Q1 & ", IdCuentaExento = 0"
            Q1 = Q1 & ", Afecto = " & vFld(Rs("ndBruto"))
            Q1 = Q1 & ", IdCuentaAfecto = " & lIdCtaBruto
         End If
         
         Q1 = Q1 & ", IVA = 0"
         Q1 = Q1 & ", IdCuentaIVA = 0"
         Q1 = Q1 & ", OtroImp = " & vFld(Rs("ndImpto"))
         Q1 = Q1 & ", IdCuentaOtroImp = " & lIdCtaImpRet
         Q1 = Q1 & ", Total = " & vFld(Rs("ndNeto"))
         
         If TReten = TR_HONORARIOS Then
            Q1 = Q1 & ", IdCuentaTotal = " & lIdCtaNetoHon
         Else
            Q1 = Q1 & ", IdCuentaTotal = " & lIdCtaNetoDieta
         End If
            
         Q1 = Q1 & ", FEmisionOri = " & CLng(Int(vFld(Rs("dtFechaDoc"))))
         Q1 = Q1 & ", CorrInterno = " & i
         Q1 = Q1 & ", SaldoDoc = NULL"
         Q1 = Q1 & ", DTE = 0" 'vFld(Rs("bDocE"))
       
         Q1 = Q1 & ", PorcentRetencion = " & PReten
         Q1 = Q1 & ", TipoRetencion = " & TReten
         Q1 = Q1 & ", IdSucursal = " & IdSucursal
         
         Q1 = Q1 & ", FExported = 0"
         Q1 = Q1 & ", OldIdDoc = 0"
         Q1 = Q1 & ", MovEdited = 1"
         Q1 = Q1 & ", OtrosVal = 0"
         Q1 = Q1 & ", FImporF29 = " & CLng(Int(Now))
         
         Q1 = Q1 & " WHERE IdDoc = " & IdDoc
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
         Call ExecSQL(DbMain, Q1)
         
         'insertamos los MovDocumento
         orden = 1
   
         'Bruto/Honor. sin Retención
         If vFld(Rs("ndImpto")) = 0 Then
            Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndBruto")), lIdCtaHonSinRet, "Honorarios sin Retención", LIBRETEN_HONORSINRET, False, False, orden)
         Else
            Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndBruto")), lIdCtaBruto, "Bruto", LIBRETEN_BRUTO, False, False, orden)
         End If
         
         'Impuesto
         Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndImpto")), lIdCtaImpRet, "Impuesto Retenido", LIBRETEN_IMPUESTO, False, True, orden)
         
         'Neto
         If TReten = TR_HONORARIOS Then
            Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndNeto")), lIdCtaNetoHon, "Neto Honorarios", LIBRETEN_NETO, True, False, orden)
         Else
            Call InsertMovDocumento(IdDoc, TipoLib, TipoDoc, vFld(Rs("ndNeto")), lIdCtaNetoDieta, "Neto Dieta", LIBRETEN_NETO, True, False, orden)
         End If
     
         i = i + 1
      
      End If
      
      Call CloseRs(RsOldDoc)
      
      Rs.MoveNext
   Loop

   Call CloseRs(Rs)
   
   ImportF29Retenciones = NDocsImp

End Function

Private Function GetTipoDocFromLAU(ByVal TipoLib As Integer, ByVal TipoDocLAU As Integer) As Integer
   Dim i As Integer
   
   GetTipoDocFromLAU = 0
   
   For i = 0 To UBound(gTipoDoc)
   
      If gTipoDoc(i).TipoLib = TipoLib And gTipoDoc(i).TipoDocLAU = TipoDocLAU Then
         GetTipoDocFromLAU = gTipoDoc(i).TipoDoc
      End If
      
   Next i
   
End Function
Private Sub LoadCuentasDef(ByVal TipoLib As Integer)
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Item As String
      
   lIdCtaAfecto = 0
   lIdCtaExento = 0
   lIdCtaTotal = 0
   lIdCtaActFijoAfecto = 0
   lIdCtaActFijoExento = 0

   lIdCtaIVA = 0
   lIdCtaOtrosImp = 0

   lIdCtaHonSinRet = 0
   lIdCtaBruto = 0

   lIdCtaImpRet = 0
   lIdCtaNetoHon = 0
   lIdCtaNetoDieta = 0
      
      
   If TipoLib > 0 Then
   
      Q1 = "SELECT CuentasBasicas.IdCuenta, Cuentas.Codigo, Cuentas.Nombre, Cuentas.Descripcion, TipoValor, Cuentas.Atrib" & ATRIB_ACTIVOFIJO & " As ActFijo "
      Q1 = Q1 & " FROM CuentasBasicas INNER JOIN Cuentas ON CuentasBasicas.IdCuenta = Cuentas.IdCuenta "
      Q1 = Q1 & JoinEmpAno(gDbType, "CuentasBasicas", "Cuentas")
      Q1 = Q1 & " WHERE TipoLib = " & TipoLib
      Q1 = Q1 & " AND CuentasBasicas.IdEmpresa = " & gEmpresa.id & " AND CuentasBasicas.Ano = " & gEmpresa.Ano
      Q1 = Q1 & " ORDER BY TipoValor, CuentasBasicas.Id "
      
      Set Rs = OpenRs(DbMain, Q1)
   
      Do While Rs.EOF = False
                      
         If TipoLib = LIB_COMPRAS Or TipoLib = LIB_VENTAS Then
                      
            Select Case vFld(Rs("TipoValor"))
            
               Case LIBVENTAS_AFECTO, LIBCOMPRAS_AFECTO
               
                  If lIdCtaAfecto = 0 Then
                     lIdCtaAfecto = vFld(Rs("IdCuenta"))
                  End If
                  
                  If vFld(Rs("ActFijo")) <> 0 And lIdCtaActFijoAfecto = 0 Then
                     lIdCtaActFijoAfecto = vFld(Rs("IdCuenta"))
                  End If
                  
               Case LIBVENTAS_EXENTO, LIBCOMPRAS_EXENTO
               
                  If lIdCtaExento = 0 Then
                     lIdCtaExento = vFld(Rs("IdCuenta"))
                  End If
               
                  If vFld(Rs("ActFijo")) <> 0 And lIdCtaActFijoExento = 0 Then
                     lIdCtaActFijoExento = vFld(Rs("IdCuenta"))
                  End If
                  
               Case LIBVENTAS_TOTAL, LIBCOMPRAS_TOTAL
               
                  If lIdCtaTotal = 0 Then
                     lIdCtaTotal = vFld(Rs("IdCuenta"))
                  End If
                  
            End Select
                           
         ElseIf TipoLib = LIB_RETEN Then
         
            Select Case vFld(Rs("TipoValor"))
            
               Case LIBRETEN_HONORSINRET
               
                  If lIdCtaHonSinRet = 0 Then
                     lIdCtaHonSinRet = vFld(Rs("IdCuenta"))
                  End If
                  
               Case LIBRETEN_BRUTO
               
                  If lIdCtaBruto = 0 Then
                     lIdCtaBruto = vFld(Rs("IdCuenta"))
                  End If
               
            End Select
         
         End If
         
         Rs.MoveNext
         
      Loop
      Call CloseRs(Rs)
      
   End If
   
   If lIdCtaActFijoAfecto = 0 Then
      lIdCtaActFijoAfecto = lIdCtaAfecto
   End If
   
   If lIdCtaActFijoExento = 0 Then
      lIdCtaActFijoExento = lIdCtaExento
   End If
   
   If TipoLib = LIB_COMPRAS Then
      lIdCtaIVA = gCtasBas.IdCtaIVACred
      lIdCtaOtrosImp = gCtasBas.IdCtaOtrosImpCred
   
   ElseIf TipoLib = LIB_VENTAS Then
      lIdCtaIVA = gCtasBas.IdCtaIVADeb               'LIB_VENTAS
      lIdCtaOtrosImp = gCtasBas.IdCtaOtrosImpDeb
   
   ElseIf TipoLib = LIB_RETEN Then
       lIdCtaImpRet = gCtasBas.IdCtaImpRet
       lIdCtaNetoHon = gCtasBas.IdCtaNetoHon
       lIdCtaNetoDieta = gCtasBas.IdCtaNetoDieta
   End If

End Sub
Private Sub InsertMovDocumento(ByVal IdDoc As Long, ByVal TipoLib As Long, ByVal TipoDoc As Integer, ByVal Valor As Double, ByVal IdCuenta As Double, ByVal Glosa As String, ByVal IdTipoValLib As Integer, ByVal EsTotalDoc As Integer, ByVal EsImptoReten As Integer, orden As Integer)
   Dim Q1 As String
   Dim QBase As String
   Dim i As Integer
   Dim TipoDocNC As Boolean
   Dim Idx As Integer
         
   QBase = "INSERT INTO MovDocumento"
   QBase = QBase & "(IdDoc, IdEmpresa, Ano, Orden, IdCuenta, Debe, Haber, Glosa, IdTipoValLib, EsTotalDoc, IdCCosto, IdAreaNeg ) "
   QBase = QBase & " VALUES(" & IdDoc & "," & gEmpresa.id & "," & gEmpresa.Ano & ","
   
   Idx = GetTipoDoc(TipoLib, TipoDoc)
   If Idx >= 0 Then
      TipoDocNC = gTipoDoc(Idx).EsRebaja
   End If
         
   If Valor <> 0 Then
      Q1 = QBase & orden & ","                                    'Orden
      Q1 = Q1 & IdCuenta & ","                                    'IdCuenta
         
      If TipoLib = LIB_COMPRAS Then
      
         If Not EsTotalDoc Then
            If TipoDocNC Or IdTipoValLib = LIBCOMPRAS_IVARETTOT Or IdTipoValLib = LIBCOMPRAS_IVARETPARC Then
               Q1 = Q1 & "0" & ","                                'Debe
               Q1 = Q1 & Valor & ","                              'Haber
            Else
               Q1 = Q1 & Valor & ","                              'Debe
               Q1 = Q1 & "0" & ","                                'Haber
            End If
         Else
            If TipoDocNC Then
               Q1 = Q1 & Valor & ","                              'Debe
               Q1 = Q1 & "0" & ","                                'Haber
            Else
               Q1 = Q1 & "0" & ","                                'Debe
               Q1 = Q1 & Valor & ","                              'Haber
            End If
         End If
         
      ElseIf TipoLib = LIB_VENTAS Then
      
         If Not EsTotalDoc Then
            If TipoDocNC Or IdTipoValLib = LIBVENTAS_REBAJA65 Then
               Q1 = Q1 & Valor & ","                              'Debe
               Q1 = Q1 & "0" & ","                                'Haber
            Else
               Q1 = Q1 & "0" & ","                                'Debe
               Q1 = Q1 & Valor & ","                              'Haber
            End If
         Else
            If TipoDocNC Then
               Q1 = Q1 & "0" & ","                                'Debe
               Q1 = Q1 & Valor & ","                              'Haber
            Else
               Q1 = Q1 & Valor & ","                              'Debe
               Q1 = Q1 & "0" & ","                                'Haber
            End If
         End If
      
      ElseIf TipoLib = LIB_RETEN Then
      
         If Not EsTotalDoc Then
            If EsImptoReten Then
               Q1 = Q1 & "0" & ","                                'Debe
               Q1 = Q1 & Valor & ","                              'Haber
            Else
               Q1 = Q1 & Valor & ","                              'Debe
               Q1 = Q1 & "0" & ","                                'Haber
            End If
         Else
            Q1 = Q1 & "0" & ","                                   'Debe
            Q1 = Q1 & Valor & ","                                 'Haber
         End If
      End If
      
      Q1 = Q1 & "'" & Glosa & "',"                                'Glosa
      Q1 = Q1 & IdTipoValLib & ","                                'IdTipoValLib
      Q1 = Q1 & EsTotalDoc & ",0,0" & ")"                         'EsTotalDoc, IdCCosto, IdAreaNeg
      
      Call ExecSQL(DbMain, Q1)
      
      orden = orden + 1
         
   End If

End Sub
Private Sub GetCtaIVAIrrec()
   Dim Q1 As String
   Dim Rs As Recordset
   
   If lIdCtaIVAIrrec <> 0 Then
      Exit Sub
   End If
   
   Q1 = "SELECT IdCuenta FROM Cuentas "
   Q1 = Q1 & " WHERE " & GenLike(DbMain, "IVA No Recuperable", "Descripcion")
   Q1 = Q1 & " OR " & GenLike(DbMain, "I.V.A. No Recuperable", "Descripcion")
   Q1 = Q1 & " OR " & GenLike(DbMain, "IVA Irrecuperable", "Descripcion")
   Q1 = Q1 & " OR " & GenLike(DbMain, "I.V.A. Irrecuperable", "Descripcion")
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

   Set Rs = OpenRs(DbMain, Q1)
   
   If Rs.EOF = False Then
      lIdCtaIVAIrrec = vFld(Rs("IdCuenta"))
   End If
   
   Call CloseRs(Rs)
   
   If lIdCtaIVAIrrec = 0 Then
      MsgBox1 "No se ha encontrado la cuenta de IVA Irrecuperable en del plan de cuentas definido para esta empresa." & vbNewLine & vbNewLine & "Los valores de IVA Irrecuperable que se importen desde HR-IVA, serán asignados momentáneamente a la cuenta de Otros Impuestos." & vbNewLine & vbNewLine & "Recuerde asignar estos valores a la(s) cuenta(s) de pérdida correspondiente(s).", vbInformation + vbOKOnly
      lIdCtaIVAIrrec = lIdCtaOtrosImp
   End If
      
End Sub
Private Function GetIdSucursal(ByVal HRCodSuc As String) As Long
   Dim Q1 As String
   Dim Rs As Recordset
   
   GetIdSucursal = 0
   
   If HRCodSuc = "" Then
      Exit Function
   End If

   Q1 = "SELECT IdSucursal FROM Sucursales WHERE Codigo = '" & HRCodSuc & "'"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
   Set Rs = OpenRs(DbMain, Q1)
   
   If Not Rs.EOF Then
      GetIdSucursal = vFld(Rs("IdSucursal"))
   End If
   
   Call CloseRs(Rs)
   
End Function
Public Function GenDB_F29(ByVal Mes As Integer, Optional ByVal Msg As Boolean = True) As Boolean
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rs As Recordset
   Dim RsDao As dao.Recordset
   Dim Q1 As String, Q2 As String
   Dim Where As String
   Dim TblName As String
   Dim j As Integer
   Dim UpdOK As Boolean
   Dim Rc As Integer
   Dim DbF29 As Database
   Dim DbF29Path As String
   Dim AnoMes As Long
   Dim Buf As String
   Dim RutMdb As String
   Dim FLast As Long
   Dim DbExpHR As String
   
   If Not gFunciones.NuevoTraspasoIVA Then
      Exit Function
   End If
         
   GenDB_F29 = False
   
   Where = SqlYearLng("FEmision") & " = " & gEmpresa.Ano
            
   If Mes > 0 Then
      Where = Where & " AND " & SqlMonthLng("FEmision") & " = " & Mes
   Else
      Exit Function
      
   End If
   
   AnoMes = DateSerial(gEmpresa.Ano, Mes, 1)
   
   'creamos la db
      
   ERR.Clear
   
   On Error Resume Next
   
   'vemos si existe la carpeta del RUT
   DbF29Path = gHRPath & "\RUTS\" & Right("000000000" & gEmpresa.Rut, 8)
   'Buf = DbF29Path & "\SDF29" & Format(gEmpresa.Ano Mod 100, "00") & ".MDB"     'AVF29-año.MDB
   Buf = DbF29Path & "\AVF29" & Format(gEmpresa.Ano Mod 100, "00") & ".MDB"
   If Not ExistFile(Buf) Then     'no existe el RUT
      MsgBox1 "Este contribuyente no ha sido creado en HR-IVA.", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   DbF29Path = DbF29Path & "\ImpConta"
   MkDir DbF29Path
   
   DbF29Path = DbF29Path & "\F29_" & Format(AnoMes, "mmyy") & ".mdb"   'Antes era: "\LPConta.mdb"
     
   If ExistFile(DbF29Path) Then   'si existe, la eliminamos para evitar problemas con archivos dañados
      Kill (DbF29Path)
   End If
        
   If Dir(DbF29Path) = "" Then    'no existe, lo creamos
'      Set DbF29 = CreateDatabase(DbF29Path, dbLangGeneral)

      DbExpHR = gDbPath & "\ExpHR_F22_F29.mdb"           'dado que usamos DAO 3.6 y HR no, usamos una base vacia creada en DAO 3.5 como base para exportar a HR
      If ExistFile(DbExpHR) Then
         Call CopyFile(DbExpHR, DbF29Path, True)
         
         If ERR Then
            Call MsgErr(DbExpHR)
            Exit Function
         End If
      Else
         MsgBox1 "No se encuentra el archivo" & vbCrLf & vbCrLf & DbExpHR & vbCrLf & vbCrLf & "Comuníquese con soporte para que le puedan ayudar.", vbExclamation
         Exit Function
      End If
   End If
'   Else
      Set DbF29 = OpenDatabase(DbF29Path)
      If ERR Then
         MsgBox1 "La base de exportación está dañada, elimine el siguiente archivo y vuelva a ejecutar." & vbCrLf & vbCrLf & DbF29Path, vbExclamation
         Exit Function
      End If
      
'   End If

   On Error GoTo 0
   
   'creamos la tabla Param con la fecha y version de exportacion
   TblName = "Param"
   
   'vemos si la tabla existe
   Q1 = "SELECT count(*) FROM " & TblName
   Set RsDao = OpenRsDao(DbF29, Q1, False)
   
   If RsDao Is Nothing Then
      'no existe
      
      'Creamos la tabla Psram
   
'      Set Tbl = New TableDef
'      Tbl.Name = TblName
      On Error Resume Next
      
      Set Tbl = DbF29.CreateTableDef(TblName)
      
      If ERR Then
         MsgErr TblName
      End If
      On Error GoTo 0
         
      ERR.Clear
'      Set Fld = Tbl.CreateField("Codigo", dbText, 15)
'      Tbl.Fields.Append Fld
      Tbl.Fields.Append Tbl.CreateField("Codigo", dbText, 15)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Codigo", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
'      Set Fld = Tbl.CreateField("Valor", dbText, 30)
'      Tbl.Fields.Append Fld
      Tbl.Fields.Append Tbl.CreateField("Valor", dbText, 30)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Valor", vbExclamation
         UpdOK = False
      End If
                 
      DbF29.TableDefs.Append Tbl
      If ERR = 0 Then
         DbF29.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON " & TblName & " (Codigo)"
         Rc = ExecSQLDao(DbF29, Q1, False)
                  
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla " & TblName, vbExclamation
         UpdOK = False
         
      End If
      
      Set Tbl = Nothing
      
   Else    'existe la tabla
   
      'eliminamos todos los registros
      Call ExecSQLDao(DbF29, "DELETE * FROM " & TblName)
      
   End If
     
   Call CloseRs(RsDao)
   
   'insertamos registros de la exportación
   Q1 = "INSERT INTO " & TblName & "(Codigo, Valor) VALUES ( 'Version','" & W.Version & "')"
   Call ExecSQLDao(DbF29, Q1)
   
   Q1 = "INSERT INTO " & TblName & "(Codigo, Valor) VALUES ( 'Fecha Version','" & Format(W.FVersion, "dd mmm yy") & "')"
   Call ExecSQLDao(DbF29, Q1)
   
   Q1 = "INSERT INTO " & TblName & "(Codigo, Valor) VALUES ( 'Fecha Export','" & Format(Now, "dd mmm yy hh:mm") & "')"
   Call ExecSQLDao(DbF29, Q1)
   
   Q1 = "INSERT INTO " & TblName & "(Codigo, Valor) VALUES ( 'RUT','" & FmtCID(gEmpresa.Rut) & "')"
   Call ExecSQLDao(DbF29, Q1)
   
   Q1 = "INSERT INTO " & TblName & "(Codigo, Valor) VALUES ( 'Mes','" & Format(AnoMes, "mm/yy") & "')"
   Call ExecSQLDao(DbF29, Q1)

   'creamos la tabla Otros para almacenar otros valores, por ejemplo el impuesto único
   FLast = DateAdd("m", 1, AnoMes) - 1

   TblName = "Otros"
   
   'vemos si la tabla existe
   Q1 = "SELECT count(*) FROM " & TblName
   Set RsDao = OpenRsDao(DbF29, Q1, False)
   
   If RsDao Is Nothing Then
      'no existe
      
      'Creamos la tabla Otros
   
'      Set Tbl = New TableDef
'      Tbl.Name = TblName
      Set Tbl = DbF29.CreateTableDef(TblName)
      
            
      ERR.Clear
'      Set Fld = Tbl.CreateField("Codigo", dbText, 15)
'      Tbl.Fields.Append Fld
      Tbl.Fields.Append Tbl.CreateField("Codigo", dbText, 15)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Codigo", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
'      Set Fld = Tbl.CreateField("Valor", dbDouble)
'      Tbl.Fields.Append Fld
      Tbl.Fields.Append Tbl.CreateField("Valor", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Valor", vbExclamation
         UpdOK = False
      End If
                 
      DbF29.TableDefs.Append Tbl
      If ERR = 0 Then
         DbF29.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON " & TblName & " (Codigo)"
         Rc = ExecSQLDao(DbF29, Q1, False)
                  
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla " & TblName, vbExclamation
         UpdOK = False
         
      End If
      
      Set Tbl = Nothing
      
   Else    'existe la tabla
   
      'eliminamos todos los registros
      Call ExecSQLDao(DbF29, "DELETE * FROM " & TblName)
      
   End If
     
   Call CloseRs(RsDao)
   
   'Impuesto único
   If gCtasBas.IdCtaImpUnico = 0 Then
      If Msg Then
         MsgBox1 "Falta definir la cuenta de Impuesto Único a los Trabajadores. No se exportará este valor.", vbExclamation + vbOKOnly
      End If
   
   Else
      If Msg Then
         'MsgBox1 "Recuerde que debe tener cuadrada la cuenta de Impuesto Único de los meses anteriores, para que el saldo de esta cuenta sea correcto.", vbInformation + vbOKOnly
         MsgBox1 "Recuerde que debe tener saldada la cuenta de Impuesto Único de los meses anteriores, para que el saldo a traspasar sea el correcto." & vbCrLf & vbCrLf & "El saldo de esta cuenta se calcula a partir de los comprobantes cuyo Tipo de Ajuste sea TRIBUTARIO o AMBOS.", vbInformation + vbOKOnly       'Victor Morales 13 dic 2012
      End If
      
      Q1 = "SELECT Sum(Haber) - Sum(Debe) as Saldo "
      Q1 = Q1 & " FROM MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
      Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      Q1 = Q1 & " WHERE IdCuenta=" & gCtasBas.IdCtaImpUnico
      Q1 = Q1 & " AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & "))"

'      Q1 = Q1 & " AND Fecha <= " & FLast                            'Victor Morales 28 oct. 2020
      Q1 = Q1 & " AND Comprobante.Tipo = " & TC_TRASPASO               'Victor Morales 11 nov 2020
      Q1 = Q1 & " AND " & SqlYearLng("Fecha") & " = " & gEmpresa.Ano
      Q1 = Q1 & " AND " & SqlMonthLng("Fecha") & " = " & Mes
      
      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         'insertamos Impuesto Único
         Q1 = "INSERT INTO " & TblName & "(Codigo, Valor) VALUES ( 'IMPUNICO'," & vFld(Rs("Saldo")) & ")"
         Call ExecSQLDao(DbF29, Q1)
      End If
      
      Call CloseRs(Rs)
      

   
   End If
   
   
      '******** 3% Ret Centralizacion Remuneracion ********** ADO 2678534 Victor Morales 11-11-2021
   If gCtasBas.IdCta3PorcCentraRem = 0 Then
      If Msg Then
         MsgBox1 "Falta definir la cuenta del 3% Ret Centralizacion Remuneracion. No se exportará este valor.", vbExclamation + vbOKOnly
      End If
   
   Else
      If Msg Then
         'MsgBox1 "Recuerde que debe tener cuadrada la cuenta de Impuesto Único de los meses anteriores, para que el saldo de esta cuenta sea correcto.", vbInformation + vbOKOnly
         MsgBox1 "Recuerde que debe tener saldada la cuenta del 3% Ret Centralizacion Remuneracion de los meses anteriores, para que el saldo a traspasar sea el correcto." & vbCrLf & vbCrLf & "El saldo de esta cuenta se calcula a partir de los comprobantes cuyo Tipo de Ajuste sea TRIBUTARIO o AMBOS.", vbInformation + vbOKOnly       'Victor Morales 13 dic 2012
      End If
      
      Q1 = "SELECT Sum(Haber) - Sum(Debe) as Saldo "
      Q1 = Q1 & " FROM MovComprobante INNER JOIN Comprobante ON MovComprobante.IdComp = Comprobante.IdComp "
      Q1 = Q1 & JoinEmpAno(gDbType, "Comprobante", "MovComprobante")
      Q1 = Q1 & " WHERE IdCuenta=" & gCtasBas.IdCta3PorcCentraRem
      Q1 = Q1 & " AND (Comprobante.TipoAjuste IS NULL OR Comprobante.TipoAjuste IN (" & TAJUSTE_TRIBUTARIO & "," & TAJUSTE_AMBOS & "))"

      Q1 = Q1 & " AND Comprobante.Tipo = " & TC_TRASPASO
      Q1 = Q1 & " AND " & SqlYearLng("Fecha") & " = " & gEmpresa.Ano
      Q1 = Q1 & " AND " & SqlMonthLng("Fecha") & " = " & Mes
      
      Q1 = Q1 & " AND MovComprobante.IdEmpresa = " & gEmpresa.id & " AND MovComprobante.Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Rs.EOF = False Then
         'insertamos 3% Ret Centralizacion Remuneracion
         Q1 = "INSERT INTO " & TblName & "(Codigo, Valor) VALUES ( '3CENREMU'," & vFld(Rs("Saldo")) & ")"
         Call ExecSQLDao(DbF29, Q1)
      End If
      
      Call CloseRs(Rs)
      
   End If
   '*****************************************************
   
   'Creamos la tabla TipoDocs
   TblName = "TipoDocs"
   
   'vemos si la tabla existe
   Q1 = "SELECT count(*) FROM " & TblName
   Set RsDao = OpenRsDao(DbF29, Q1, False)
   
   If RsDao Is Nothing Then
      'no existe
   
      Set Tbl = DbF29.CreateTableDef(TblName)
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoLib", dbInteger)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".TipoLib", vbExclamation
         UpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoDoc", dbInteger)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".TipoDoc", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Nombre", dbText, 30)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Nombre", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Diminutivo", dbText, 10)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Diminutivo", vbExclamation
         UpdOK = False
      End If
                            
      DbF29.TableDefs.Append Tbl
      If ERR = 0 Then
         DbF29.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON " & TblName & " (TipoLib, TipoDoc)"
         Rc = ExecSQLDao(DbF29, Q1, False)
                  
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla " & TblName, vbExclamation
         UpdOK = False
         
      End If
      
      Set Tbl = Nothing
      
   Else    'existe la tabla
   
      'eliminamos todos los registros
      Call ExecSQLDao(DbF29, "DELETE * FROM " & TblName)
      
   End If
     
   Call CloseRs(RsDao)

   'copiamos TipoDocs
   Q1 = "SELECT TipoLib, TipoDoc, Nombre, Diminutivo "
   Q1 = Q1 & " FROM TipoDocs"
   Q1 = Q1 & " WHERE TipoLib IN(" & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
   Q1 = Q1 & " ORDER BY TipoLib, TipoDoc "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      Q1 = "INSERT INTO TipoDocs(TipoLib, TipoDoc, Nombre, Diminutivo)"
      Q1 = Q1 & " VALUES(" & vFld(Rs("TipoLib"))
      Q1 = Q1 & " , " & vFld(Rs("TipoDoc"))
      Q1 = Q1 & " , '" & vFld(Rs("Nombre")) & "'"
      Q1 = Q1 & " , '" & vFld(Rs("Diminutivo")) & "')"
      Call ExecSQLDao(DbF29, Q1, False)
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   
   'Creamos la tabla TipoValor
   TblName = "TipoValor"
   
   'vemos si la tabla existe
   Q1 = "SELECT count(*) FROM " & TblName
   Set RsDao = OpenRsDao(DbF29, Q1, False)
   
   If RsDao Is Nothing Then
      'no existe
   
      Set Tbl = DbF29.CreateTableDef(TblName)
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("idTValor", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".idTValor", vbExclamation
         UpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoLib", dbInteger)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".TipoLib", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Codigo", dbByte)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Codigo", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Valor", dbText, 50)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Valor", vbExclamation
         UpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29", dbInteger)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".CodF29", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodF29_Adic", dbInteger)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".CodF29_Adic", vbExclamation
         UpdOK = False
      End If
           
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodImpSII", dbText, 3)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".CodImpSII", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      
      Tbl.Fields.Append Tbl.CreateField("CodSIIDTE", dbText, 3)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".CodSIIDTE", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Atributo", dbText, 10)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Atributo", vbExclamation
         UpdOK = False
      End If
                                       
      DbF29.TableDefs.Append Tbl
      If ERR = 0 Then
         DbF29.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON " & TblName & " (TipoLib, Codigo)"
         Rc = ExecSQLDao(DbF29, Q1, False)
                  
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla " & TblName, vbExclamation
         UpdOK = False
         
      End If
      
      Set Tbl = Nothing
      
   Else    'existe la tabla
   
      'eliminamos todos los registros
      Call ExecSQLDao(DbF29, "DELETE * FROM " & TblName)
      
   End If
     
   Call CloseRs(RsDao)

   Call AddDebug("GenDB_F29: Ln3273 ")
   
   'copiamos TipoValor
   Q1 = "SELECT idTValor, TipoLib, Codigo, Valor, CodF29, CodF29_Adic, CodImpSII, CodSIIDTE, Atributo "
   Q1 = Q1 & " FROM TipoValor"
   Q1 = Q1 & " WHERE TipoLib IN(" & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ") "
   Q1 = Q1 & " ORDER BY TipoLib, Codigo "
   Set Rs = OpenRs(DbMain, Q1)
   
   Call AddDebug("GenDB_F29: Ln3282")
   
   If Not Rs Is Nothing Then
   
     Call AddDebug("GenDB_F29: Ln3286 ")
     
     Do While Not Rs.EOF
         Q1 = "INSERT INTO TipoValor(idTValor, TipoLib, Codigo, Valor, CodF29, CodF29_Adic, CodImpSII, CodSIIDTE, Atributo)"
         Q1 = Q1 & " VALUES(" & vFld(Rs("idTValor"))
         Q1 = Q1 & " , " & vFld(Rs("TipoLib"))
         Q1 = Q1 & " , " & vFld(Rs("Codigo"))
         Q1 = Q1 & " , '" & vFld(Rs("Valor")) & "'"
         Q1 = Q1 & " , " & vFld(Rs("CodF29"))
         Q1 = Q1 & " , " & vFld(Rs("CodF29_Adic"))
         Q1 = Q1 & " , '" & vFld(Rs("CodImpSII")) & "'"
         Q1 = Q1 & " , '" & vFld(Rs("CodSIIDTE")) & "'"
         Q1 = Q1 & " , '" & vFld(Rs("Atributo")) & "')"
         
         Call AddDebug("GenDB_F29: Ln 3300 - Q1=" & Q1)
         
         Call ExecSQLDao(DbF29, Q1, False)
         
         Rs.MoveNext
      Loop
      
      Call CloseRs(Rs)
      
      Call AddDebug("GenDB_F29: Ln 3309 =")
      
   End If
   
   'Creamos la tabla Documento
   TblName = "Documento"
   
   'vemos si la tabla existe
   Q1 = "SELECT count(*) FROM " & TblName
   Set RsDao = OpenRsDao(DbF29, Q1, False)
   
   If RsDao Is Nothing Then
      'no existe
   
      Set Tbl = DbF29.CreateTableDef(TblName)
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdDoc", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".IdDoc", vbExclamation
         UpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoLib", dbInteger)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".TipoLib", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoDoc", dbInteger)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".TipoDoc", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NumDoc", dbText, 20)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".NumDoc", vbExclamation
         UpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("NumDocHasta", dbText, 20)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".NumDocHasta", vbExclamation
         UpdOK = False
      End If
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CantBoletas", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".CantBoletas", vbExclamation
         UpdOK = False
      End If
                      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("DTE", dbText, 1)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".DTE", vbExclamation
         UpdOK = False
      End If
                      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Giro", dbText, 1)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Giro", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("EsDeSuper", dbText, 1)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".EsDeSuper", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("FEmision", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".FEmision", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Descrip", dbText, 100)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Descrip", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Exento", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Exento", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Afecto", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Afecto", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IVA", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".IVA", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("OtroImp", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".OtroImp", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Total", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Total", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("PorcentRetencion", dbByte)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".PorcentRetencion", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoRetencion", dbByte)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".TipoRetencion", vbExclamation
         UpdOK = False
      End If
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("ValRet3Porc", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".ValRet3Porc", vbExclamation
         UpdOK = False
      End If
           
           
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IVAInmueble", dbText, 1)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".IVAInmueble", vbExclamation
         UpdOK = False
      End If
                                       
      DbF29.TableDefs.Append Tbl
      If ERR = 0 Then
         DbF29.TableDefs.Refresh
         
         Q1 = "DROP INDEX Idx ON " & TblName
         Rc = ExecSQLDao(DbF29, Q1, False)
         
         Q1 = "CREATE INDEX Idx ON " & TblName & " (TipoLib, TipoDoc, NumDoc)"
         Rc = ExecSQLDao(DbF29, Q1, False)
                  
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla " & TblName, vbExclamation
         UpdOK = False
         
      End If
      
      Set Tbl = Nothing
      
      Call CloseRs(RsDao)
     
   Else    'existe la tabla
      Call CloseRs(RsDao)
  
      'eliminamos todos los registros
      Call ExecSQLDao(DbF29, "DELETE * FROM " & TblName)
      
      'eliminamos y volvemos a crear índice porque antes se creaba como índice único y no sirve porque pueden haber documentos con el mismo número y con distinta entidad y como la entidad no se traspasa, se duplica el registro
      Q1 = "DROP INDEX Idx ON " & TblName
      Rc = ExecSQLDao(DbF29, Q1, False)
         
      Q1 = "CREATE INDEX Idx ON " & TblName & " (TipoLib, TipoDoc, NumDoc)"
      Rc = ExecSQLDao(DbF29, Q1, False)
  
   End If
     

   'copiamos Documento
   Q1 = "SELECT IdDoc, TipoLib, TipoDoc, NumDoc, NumDocHasta, CantBoletas, iif(Documento.DTE <> 0, 'S', 'N') As DTE, "
   Q1 = Q1 & " Entidades.Rut, iif(Documento.Giro <> 0, 'S', 'N') As Giro, "
   Q1 = Q1 & " iif(Entidades.EsSupermercado <> 0, 'S', 'N' ) As EsDeSuper, "
   Q1 = Q1 & " FEmision, Descrip, Exento, Afecto, IVA, OtroImp, Total, PorcentRetencion, TipoRetencion, ValRet3Porc, "      'FCA - 12/10/2021
   Q1 = Q1 & " iif(Documento.IVAInmueble <> 0, 'S', 'N') As IVAInmueble "
   Q1 = Q1 & " FROM Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " AND Documento.IdEmpresa = Entidades.IdEmpresa "
   Q1 = Q1 & " WHERE TipoLib IN(" & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
   Q1 = Q1 & " AND " & Where
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY FEmision, IdDoc "
   Set Rs = OpenRs(DbMain, Q1)
   
   Do While Not Rs.EOF
      
      Q1 = "INSERT INTO Documento(IdDoc, TipoLib, TipoDoc, NumDoc, NumDocHasta, CantBoletas, DTE, "
      Q1 = Q1 & " Giro, EsDeSuper, FEmision, Descrip, Exento, Afecto, IVA, OtroImp, Total, PorcentRetencion, TipoRetencion, ValRet3Porc, IVAInmueble) "
      Q1 = Q1 & " VALUES(" & vFld(Rs("IdDoc"))
      Q1 = Q1 & " , " & vFld(Rs("TipoLib"))
      Q1 = Q1 & " , " & vFld(Rs("TipoDoc"))
      Q1 = Q1 & " , '" & vFld(Rs("NumDoc")) & "'"
      Q1 = Q1 & " , '" & vFld(Rs("NumDocHasta")) & "'"
      Q1 = Q1 & " , " & vFld(Rs("CantBoletas"))
      Q1 = Q1 & " , '" & vFld(Rs("DTE")) & "'"
      Q1 = Q1 & " , '" & vFld(Rs("Giro")) & "'"
      Q1 = Q1 & " , '" & vFld(Rs("EsDeSuper")) & "'"
      Q1 = Q1 & " , " & vFld(Rs("FEmision"))
      Q1 = Q1 & " , '" & Replace(vFld(Rs("Descrip")), "'", "") & "'"
      Q1 = Q1 & " , " & vFld(Rs("Exento"))
      Q1 = Q1 & " , " & vFld(Rs("Afecto"))
      Q1 = Q1 & " , " & vFld(Rs("IVA"))
      Q1 = Q1 & " , " & vFld(Rs("OtroImp"))
      Q1 = Q1 & " , " & vFld(Rs("Total"))
      Q1 = Q1 & " , " & vFld(Rs("PorcentRetencion"))
      Q1 = Q1 & " , " & vFld(Rs("TipoRetencion"))
      Q1 = Q1 & " , " & vFld(Rs("ValRet3Porc"))
      Q1 = Q1 & " , '" & vFld(Rs("IVAInmueble")) & "')"
      Call ExecSQLDao(DbF29, Q1, False)
            
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
  
   'Creamos la tabla Detalle
   TblName = "Detalle"
   
   'vemos si la tabla existe
   Q1 = "SELECT count(*) FROM " & TblName
   Set RsDao = OpenRsDao(DbF29, Q1, False)
   
   If RsDao Is Nothing Then
      'no existe
   
      Set Tbl = DbF29.CreateTableDef(TblName)
      
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdMovDoc", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".IdMovDoc", vbExclamation
         UpdOK = False
      End If
  
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdDoc", dbLong)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".IdDoc", vbExclamation
         UpdOK = False
      End If
  
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Debe", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Debe", vbExclamation
         UpdOK = False
      End If
  
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Haber", dbDouble)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Haber", vbExclamation
         UpdOK = False
      End If
  
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("IdTipoValLib", dbInteger)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".IdTipoValLib", vbExclamation
         UpdOK = False
      End If
  
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("EsTotalDoc", dbBoolean)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".EsTotalDoc", vbExclamation
         UpdOK = False
      End If
  
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("Tasa", dbSingle)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".Tasa", vbExclamation
         UpdOK = False
      End If
  
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("CodSIIDTE", dbText, 2)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".CodSIIDTE", vbExclamation
         UpdOK = False
      End If
  
      ERR.Clear
      Tbl.Fields.Append Tbl.CreateField("EsRecuperable", dbBoolean)
      
      If ERR = 0 Then
         Tbl.Fields.Refresh
      ElseIf ERR <> 3191 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & TblName & ".EsRecuperable", vbExclamation
         UpdOK = False
      End If
  
      DbF29.TableDefs.Append Tbl
      If ERR = 0 Then
         DbF29.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON " & TblName & " (IdDoc, IdMovDoc)"
         Rc = ExecSQLDao(DbF29, Q1, False)
                  
      ElseIf ERR <> 3010 Then ' ya existe
         MsgBox1 "Error " & ERR & ", " & Error & vbLf & "Tabla " & TblName, vbExclamation
         UpdOK = False
         
      End If
      
      Set Tbl = Nothing
      
   Else    'existe la tabla
   
      'eliminamos todos los registros
      Call ExecSQLDao(DbF29, "DELETE * FROM " & TblName)
      
   End If
     
   Call CloseRs(RsDao)
  
   Call AddDebug("GenDB_F29: Ln 3722 ")
  
   'copiamos MovDocumento (Detalle en la base DBF29)
   Q1 = "SELECT IdMovDoc, MovDocumento.IdDoc, Debe, Haber, IdTipoValLib, EsTotalDoc, Tasa, CodSIIDTE, EsRecuperable "
   Q1 = Q1 & " FROM MovDocumento INNER JOIN Documento ON MovDocumento.IdDoc = Documento.IdDoc "
   Q1 = Q1 & JoinEmpAno(gDbType, "MovDocumento", "Documento")
   Q1 = Q1 & " WHERE Documento.TipoLib IN(" & LIB_COMPRAS & "," & LIB_VENTAS & "," & LIB_RETEN & ")"
   Q1 = Q1 & " AND " & Where
   Q1 = Q1 & " AND MovDocumento.IdEmpresa = " & gEmpresa.id & " AND MovDocumento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY FEmision, Documento.IdDoc, IdMovDoc "
   Set Rs = OpenRs(DbMain, Q1)
  
   Call AddDebug("GenDB_F29: Ln 3734 ")
  
   Do While Not Rs.EOF
      
      Q1 = "INSERT INTO Detalle(IdMovDoc, IdDoc, Debe, Haber, IdTipoValLib, EsTotalDoc, Tasa, CodSIIDTE, EsRecuperable )"
      Q1 = Q1 & " VALUES(" & vFld(Rs("IdMovDoc"))
      Q1 = Q1 & " , " & vFld(Rs("IdDoc"))
      Q1 = Q1 & " , " & Round(vFld(Rs("Debe")))
      Q1 = Q1 & " , " & Round(vFld(Rs("Haber")))
      Q1 = Q1 & " , " & vFld(Rs("IdTipoValLib"))
      Q1 = Q1 & " , " & vFld(Rs("EsTotalDoc"))
      Q1 = Q1 & " , " & str(vFld(Rs("Tasa")))
      Q1 = Q1 & " , '" & vFld(Rs("CodSIIDTE")) & "'"
      Q1 = Q1 & " , " & vFld(Rs("EsRecuperable")) & ")"
      
      Call AddDebug("GenDB_F29: Ln 3734 Q1 = " & Q1)
      
      Call ExecSQLDao(DbF29, Q1, False)
      
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   Call AddDebug("GenDB_F29: Ln 3758 ")
   
   Call CloseDb(DbF29)
   
   Call AddDebug("GenDB_F29: Ln 3762 ")
   
   GenDB_F29 = True
   
End Function

