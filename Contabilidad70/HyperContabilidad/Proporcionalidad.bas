Attribute VB_Name = "Proporcionalidad"
Option Explicit

Type ValPropIVA_t
   TotAfecto As Double
   TotExento As Double
   Total As Double
   CalcProp As Boolean
   AcumAfecto As Double
   AcumTotal As Double
   Proporcion As Double
End Type

Public gValPropIVA(12) As ValPropIVA_t
Public gPrimerMesProp As Integer


Public Const PIVA_SINPROP = 0
Public Const PIVA_TOTAL = 1
Public Const PIVA_NULO = 2
Public Const PIVA_PROP = 3       'proporcional

Public gStrPropIVA(PIVA_PROP)
Public gDescPropIVA(PIVA_PROP)

Public Sub InitPropIVA()

   gStrPropIVA(PIVA_SINPROP) = ""
   gStrPropIVA(PIVA_TOTAL) = "Total"
   gStrPropIVA(PIVA_NULO) = "Nulo"
   gStrPropIVA(PIVA_PROP) = "Proporcional"

   gDescPropIVA(PIVA_SINPROP) = ""
   gDescPropIVA(PIVA_TOTAL) = "Total: IVA CF del documento se aprovecha 100%"
   gDescPropIVA(PIVA_NULO) = "Nulo: IVA se aprovecha un 0% (también se aplica para las facturas fuera de plazo)"
   gDescPropIVA(PIVA_PROP) = "Proporcional: efectuar proporcionalidad al IVA del documento"

End Sub



'Actualiza los valores de totales mensuales en la tabla PropIVA en la base de datos y los deja cargados en el arreglo en memoria
Public Sub PropIVA_UpdateTblTotMensual(Optional ByVal FillMes As Integer = 0, Optional ByVal bForce As Boolean = False)
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Mes As Integer, n As Integer
   Dim i As Integer
   Dim DocExp As String

   n = 0

   If FillMes = 0 Then
      If Not bForce Then
         Q1 = "SELECT Count(*)  FROM PropIVA_TotMensual"
         Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Set Rs = OpenRs(DbMain, Q1)
         
         If Not Rs.EOF Then
            n = vFld(Rs(0))
         End If
         
         Call CloseRs(Rs)
         
         If n = 12 Then
            'se supone que la tabla está llena y actualizada
            
            Call PropIVA_LoadTotMensual    'cargamos arreglo en memoria
            Exit Sub
         
         End If
      End If

      If bForce Or n < 12 Then
         Q1 = " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call DeleteSQL(DbMain, "PropIVA_TotMensual", Q1)   'por si las moscas
         
         'insertamos todos los meses con totales en cero
         For i = 1 To 12
            Q1 = "INSERT INTO PropIVA_TotMensual (Mes, TotalAfecto, TotalExento, IdEmpresa, Ano) VALUES( " & i & ", 0, 0, " & gEmpresa.id & "," & gEmpresa.Ano & ")"
            Call ExecSQL(DbMain, Q1)
         Next i
      
      End If

   ElseIf FillMes < 1 Or FillMes > 12 Then
      MsgBox1 "Mes inválido.", vbExclamation
      Exit Sub
      
   End If

   DocExp = " (TipoDocs.Diminutivo = 'EXP' or TipoDocs.Diminutivo = 'NCE' or TipoDocs.Diminutivo = 'NDE') "
   
   'ahora llenamos la tabla con los totales Afecto y Exento del Libro de Ventas
   Q1 = "SELECT " & SqlMonthLng("Documento.FEmision") & " as Mes, "
   Q1 = Q1 & " Sum( iif( EsRebaja <> 0, -1 * Documento.Afecto, Documento.Afecto)) as SumAfecto, "
   'Q1 = Q1 & " Sum( iif(EsRebaja, -1 * Documento.Exento, Documento.Exento)) As SumExento "
   
   Q1 = Q1 & " Sum( iif (Not " & DocExp & ", iif(EsRebaja <> 0, -1 * Documento.Exento, Documento.Exento), 0)) As SumExento, "
   Q1 = Q1 & " Sum( iif (" & DocExp & ", iif(EsRebaja <> 0, -1 * Documento.Exento, Documento.Exento), 0)) As SumExentoExp "
   
   Q1 = Q1 & " FROM TipoDocs INNER JOIN Documento ON (TipoDocs.TipoDoc = Documento.TipoDoc) AND (TipoDocs.TipoLib = Documento.TipoLib)"
   Q1 = Q1 & " WHERE Documento.TipoLib = " & LIB_VENTAS & " AND " & SqlYearLng("FEmision") & " = " & gEmpresa.Ano
   'Q1 = Q1 & " AND TipoDocs.Diminutivo NOT IN ('EXP', 'NCE', 'NDE')"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   


   If FillMes > 0 Then
      Q1 = Q1 & " AND " & SqlMonthLng("Documento.FEmision") & " = " & FillMes
   End If
   
   Q1 = Q1 & " GROUP BY  " & SqlMonthLng("Documento.FEmision")
   Q1 = Q1 & " ORDER BY  " & SqlMonthLng("Documento.FEmision")

   Set Rs = OpenRs(DbMain, Q1)

   Do While Not Rs.EOF
      Mes = vFld(Rs("Mes"))
                  
      Q1 = "UPDATE PropIVA_TotMensual "
      Q1 = Q1 & " SET TotalAfecto = " & vFld(Rs("SumAfecto")) + vFld(Rs("SumExentoExp"))
      Q1 = Q1 & ",    TotalExento = " & vFld(Rs("SumExento"))
      Q1 = Q1 & " WHERE Mes = " & Mes
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Call ExecSQL(DbMain, Q1)

      Rs.MoveNext
   Loop

   Call CloseRs(Rs)

   Call PropIVA_LoadTotMensual(FillMes)


End Sub

'carga desde la tabla en la DB los totales de Afecto y Exento mensuales y calcula el total mensual
Public Sub PropIVA_LoadTotMensual(Optional ByVal LoadMes As Integer = 0)
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Es As Integer
   Dim i As Integer
   Dim Mes As Integer
   

   If LoadMes = 0 Then
      For i = 0 To 12
         gValPropIVA(i).TotAfecto = 0
         gValPropIVA(i).TotExento = 0
         gValPropIVA(i).Total = 0
         gValPropIVA(i).AcumAfecto = 0
         gValPropIVA(i).AcumTotal = 0
         gValPropIVA(i).CalcProp = False
         gValPropIVA(i).Proporcion = False
      Next i
   End If
   
   Q1 = "SELECT Mes, TotalAfecto, TotalExento "
   Q1 = Q1 & " FROM PropIVA_TotMensual"
   Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   If LoadMes > 0 Then
      Q1 = Q1 & " AND Mes = " & LoadMes
   End If
   
   Q1 = Q1 & " ORDER BY Mes"
   
   Set Rs = OpenRs(DbMain, Q1)
         
   Do While Not Rs.EOF
      Mes = vFld(Rs("Mes"))
                  
      gValPropIVA(Mes).TotAfecto = vFld(Rs("TotalAfecto"))
      gValPropIVA(Mes).TotExento = vFld(Rs("TotalExento"))
      gValPropIVA(Mes).Total = gValPropIVA(Mes).TotAfecto + gValPropIVA(Mes).TotExento
            
      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)
   
   Call PropIVA_CalcTotMensual
   
End Sub


'calcula los totales acumulados y la proporción de IVA para cada mes
Public Sub PropIVA_CalcTotMensual()
   Dim Rs As Recordset
   Dim Q1 As String
   Dim Es As Integer
   Dim CurrMes As Integer
   Dim i As Integer
   Dim Mes As Integer
   
   gPrimerMesProp = 0
   
   For i = 0 To 12
      gValPropIVA(i).AcumAfecto = 0
      gValPropIVA(i).AcumTotal = 0
      gValPropIVA(i).CalcProp = False
   Next i

   
   For Mes = 1 To 12
      gValPropIVA(Mes).Total = gValPropIVA(Mes).TotAfecto + gValPropIVA(Mes).TotExento
      
      If gPrimerMesProp = 0 Then
      
         If Not gPrimerMesProp And gValPropIVA(Mes).TotAfecto <> 0 And gValPropIVA(Mes).TotExento <> 0 Then   'marcamos el primer mes a partir del cual se acumulan los totales y parte la proporcionalidad
            gPrimerMesProp = Mes
         End If
         
      End If
               
      If gPrimerMesProp > 0 Then
      
         gValPropIVA(Mes).CalcProp = True
         
         If Mes > 1 And Mes > gPrimerMesProp Then
            gValPropIVA(Mes).AcumAfecto = gValPropIVA(Mes - 1).AcumAfecto
            gValPropIVA(Mes).AcumTotal = gValPropIVA(Mes - 1).AcumTotal
         End If
         
         gValPropIVA(Mes).AcumAfecto = gValPropIVA(Mes).AcumAfecto + gValPropIVA(Mes).TotAfecto
         gValPropIVA(Mes).AcumTotal = gValPropIVA(Mes).AcumTotal + gValPropIVA(Mes).Total
         
         If gValPropIVA(Mes).AcumTotal > 0 Then
            gValPropIVA(Mes).Proporcion = gValPropIVA(Mes).AcumAfecto / gValPropIVA(Mes).AcumTotal
            
            If gValPropIVA(Mes).Proporcion > 1 Then
               gValPropIVA(Mes).Proporcion = 1
            End If
            
         Else
            gValPropIVA(Mes).Proporcion = 1      'Victor Morales entregó esta especificación en reporte de error 111, el día 7 oct. 2013
         
         End If
      
      ElseIf gValPropIVA(Mes).TotAfecto > 0 Then
         gValPropIVA(Mes).Proporcion = 1      'Victor Morales entregó esta especificación en reporte de error 111, el día 7 oct. 2013
         
      Else
         gValPropIVA(Mes).Proporcion = 0
         
      End If
      
   Next Mes

End Sub

Public Function PropIVA_UpdateMovDoc(Optional ByVal Mes As Integer = 0, Optional ProgBar As Object = Nothing) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset, RsAux As Recordset
   Dim Idx As Long
   Dim TipoLib As Integer
   Dim TipoDoc As Integer
   Dim TipoDocNC As Boolean
   Dim Wh As String
   Dim PropIVA As Double
   Dim IdMovIvaCredFiscal As Long
   Dim IdMovIvaActFijo As Long
   Dim IdMovIvaIrrec As Long
   Dim QBase As String
   Dim Glosa As String
   Dim orden As Integer
   Dim IVAIrrec As Integer
   Dim nDocs As Long
   Dim n As Long, i As Integer
   Dim nTipoDOc() As Integer
   Dim Msg As String
   Dim IdDoc As Long
   Dim ValIVAIrrec As Double
   Dim CalcPropIVA As Boolean
   Dim CodSIIDTEIvaIrrec As Integer
   Dim TasaIVAIrrec As Single
   Dim TipoIVAIrrec As Integer
   
   ReDim nTipoDOc(UBound(gTipoDoc))

   PropIVA_UpdateMovDoc = False
     
   If gCtasBas.IdCtaIVAIrrec = 0 Then
      Exit Function
   End If

   QBase = "INSERT INTO MovDocumento"
   QBase = QBase & "(IdEmpresa, Ano, IdDoc, Orden, IdCuenta, Debe, Haber, Glosa, IdTipoValLib, EsTotalDoc, IdCCosto, IdAreaNeg, Tasa, EsRecuperable, CodSIIDTE) "
   QBase = QBase & " VALUES(" & gEmpresa.id & "," & gEmpresa.Ano & ","
      
   PropIVA_UpdateMovDoc = False
   
   Wh = " WHERE " & SqlYearLng("FEmision") & " = " & gEmpresa.Ano
   Wh = Wh & " AND Estado = " & ED_PENDIENTE
   Wh = Wh & " AND PropIVA > 0 "


   If Mes > 0 Then
      'Wh = Wh & " AND IdDoc = " & IdDoc
      Wh = Wh & " AND " & SqlMonthLng("FEmision") & " = " & Mes '  & " AND " & SqlYearLng("FEmision") & " = " & gEmpresa.Ano (ya está arriba)
   End If
   
   'primero contamos los registros
   Q1 = "SELECT Count(*) "
   Q1 = Q1 & " FROM Documento " & Wh
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

   Set Rs = OpenRs(DbMain, Q1)

   If Not Rs.EOF Then
      nDocs = Rs(0)
   End If
   Call CloseRs(Rs)
   
   'ahora obtenemos los documentos uno a uno
   
   Q1 = "SELECT IdDoc, " & SqlMonthLng("FEmision") & " as MesDoc, TipoLib, TipoDoc, PropIVA, IVA, Estado, Descrip "
   Q1 = Q1 & " FROM Documento " & Wh
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY FEmision "

   Set Rs = OpenRs(DbMain, Q1)

   
   If Rs.EOF Then
      Call CloseRs(Rs)
      Exit Function
   End If
      
   Do While Not Rs.EOF

      IdDoc = vFld(Rs("IdDoc"))
   
      TipoLib = vFld(Rs("TipoLib"))
      TipoDoc = vFld(Rs("TipoDoc"))
      
      Glosa = Left(vFld(Rs("Descrip")), 50)
      
      IVAIrrec = 0
      ValIVAIrrec = 0

      If TipoLib <> LIB_COMPRAS Then
         Call CloseRs(Rs)
         Exit Function
      End If
   
      Idx = GetTipoDoc(TipoLib, TipoDoc)
      If Idx >= 0 Then
         TipoDocNC = gTipoDoc(Idx).EsRebaja
         nTipoDOc(Idx) = nTipoDOc(Idx) + 1
      End If
                  
      'buscamos los movimientos que nos interesan: IVA Crédito Fiscal, IVA Activo Fijo e IVA Irrecuperable
      Q1 = "SELECT IdMovDoc, IdTipoValLib, CodSIIDTE FROM MovDocumento "
      Q1 = Q1 & " WHERE IdDoc = " & IdDoc
      Q1 = Q1 & " AND IdTipoValLib IN( " & LIBCOMPRAS_IVACREDFISC & "," & LIBCOMPRAS_IVAACTFIJO & "," & LIBCOMPRAS_IVAIRREC & "," & LIBCOMPRAS_IVAIRREC1 & "," & LIBCOMPRAS_IVAIRREC2 & "," & LIBCOMPRAS_IVAIRREC3 & "," & LIBCOMPRAS_IVAIRREC4 & "," & LIBCOMPRAS_IVAIRREC9 & ")"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

      Set RsAux = OpenRs(DbMain, Q1)

      IdMovIvaCredFiscal = 0
      IdMovIvaActFijo = 0
      IdMovIvaIrrec = 0
      TipoIVAIrrec = 0
      CodSIIDTEIvaIrrec = 0
      TasaIVAIrrec = 0
      
      Do While Not RsAux.EOF
      
         Select Case vFld(RsAux("IdTipoValLib"))
         
            Case LIBCOMPRAS_IVACREDFISC
               IdMovIvaCredFiscal = vFld(RsAux("idMovDoc"))    'se supone que a lo más hay 1 registro de cada uno de estos TipoValLib, si marcó proporcionalidad de algún tipo
            
            Case LIBCOMPRAS_IVAACTFIJO
               IdMovIvaActFijo = vFld(RsAux("idMovDoc"))
            
            Case LIBCOMPRAS_IVAIRREC, LIBCOMPRAS_IVAIRREC1, LIBCOMPRAS_IVAIRREC2, LIBCOMPRAS_IVAIRREC3, LIBCOMPRAS_IVAIRREC4, LIBCOMPRAS_IVAIRREC9
               IdMovIvaIrrec = vFld(RsAux("idMovDoc"))
               TipoIVAIrrec = vFld(RsAux("IdTipoValLib"))
               CodSIIDTEIvaIrrec = Val(vFld(RsAux("CodSIIDTE")))
         End Select
         
         
         RsAux.MoveNext
         
      Loop

      Call CloseRs(RsAux)

      
      Select Case vFld(Rs("PropIVA"))
            
         Case PIVA_TOTAL
         
            IVAIrrec = IVAIRREC_CERO
            ValIVAIrrec = 0
         
            'asignamos el total del IVA al crédito fiscal o crédito activo fijo y nada a IVA Irrecuperable
            
            If IdMovIvaCredFiscal = 0 And IdMovIvaActFijo = 0 Then   'no hay registro de IVA Cred o IVA Act Fijo, lo insertamos
            
               orden = GetOrdenMovDoc(IdDoc)
            
               Q1 = QBase & IdDoc & "," & orden & ","             'IdDoc, Orden
               Q1 = Q1 & gCtasBas.IdCtaIVACred & ","              'IdCuenta
               
               If TipoDocNC Then
                  Q1 = Q1 & "0" & ","                             'Debe
                  Q1 = Q1 & Abs(vFld(Rs("IVA"))) & ","            'Haber
               Else
                  Q1 = Q1 & Abs(vFld(Rs("IVA"))) & ","            'Debe
                  Q1 = Q1 & "0" & ","                             'Haber
               End If
               
               Q1 = Q1 & "'" & Glosa & "',"                       'Glosa
               Q1 = Q1 & LIBCOMPRAS_IVACREDFISC & ","             'IdTipoValLib
                                 
               Q1 = Q1 & "0,0,0,0,0,''" & ")"                     'EsTotalDoc, IdCCosto, IdAreaNeg, Tasa, EsRecuperable, CodSIIDTE

               Call ExecSQL(DbMain, Q1)


            Else   'hay un registro de IVA Cred o de IVA Act Fijo => actualizamos el que hay
               
               Q1 = "UPDATE MovDocumento SET "

               If TipoDocNC Then
                  Q1 = Q1 & "  Debe = 0 "
                  Q1 = Q1 & ", Haber = " & Abs(vFld(Rs("IVA")))
               Else
                  Q1 = Q1 & "  Debe = " & Abs(vFld(Rs("IVA")))
                  Q1 = Q1 & ", Haber = 0 "
               End If

               If IdMovIvaCredFiscal <> 0 Then
                  Q1 = Q1 & " WHERE IdMovDoc = " & IdMovIvaCredFiscal
               Else
                  Q1 = Q1 & " WHERE IdMovDoc = " & IdMovIvaActFijo
               End If
               
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

               Call ExecSQL(DbMain, Q1)

            End If
               
            'eliminamos, si lo hubiera, un movimiento de IVA irrecuperable
            If IdMovIvaIrrec > 0 Then
               Wh = " WHERE IdMovDoc = " & IdMovIvaIrrec
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'ffv delete
              Call DeleteSQL(DbMain, "MovDocumento", Wh)

               CodSIIDTEIvaIrrec = 0
            End If

         Case PIVA_NULO
         
            IVAIrrec = IVAIRREC_TOTAL
            ValIVAIrrec = Abs(vFld(Rs("IVA")))
         
            'asignamos el total del IVA al IVA Irrecuperable
            
            If IdMovIvaIrrec = 0 Then        'no existe el reg de IVA Irrecuperable, lo agregamos a IVA Irrecuperable 1
                              
               orden = GetOrdenMovDoc(IdDoc)
               
               Q1 = QBase & IdDoc & "," & orden & ","             'IdDoc, Orden
               Q1 = Q1 & gCtasBas.IdCtaIVAIrrec & ","             'IdCuenta
               
               If TipoDocNC Then
                  Q1 = Q1 & "0" & ","                             'Debe
                  Q1 = Q1 & Abs(vFld(Rs("IVA"))) & ","            'Haber
               Else
                  Q1 = Q1 & Abs(vFld(Rs("IVA"))) & ","            'Debe
                  Q1 = Q1 & "0" & ","                             'Haber
               End If
               
               Q1 = Q1 & "'" & Glosa & "',"                       'Glosa
               Q1 = Q1 & LIBCOMPRAS_IVAIRREC1 & ","                'IdTipoValLib
                                 
               Q1 = Q1 & "0,0,0,100,0,'1'" & ")"                  'EsTotalDoc, IdCCosto, IdAreaNeg, Tasa, EsRecuperable, CodSIIDTE

               Call ExecSQL(DbMain, Q1)

            
               CodSIIDTEIvaIrrec = 1
               
            Else        'actualizamos el registro de IVA Irrecuperable con el total del IVA
            
            
               Q1 = "UPDATE MovDocumento SET "

               If TipoDocNC Then
                  Q1 = Q1 & "  Debe = 0 "
                  Q1 = Q1 & ", Haber = " & Abs(vFld(Rs("IVA")))
               Else
                  Q1 = Q1 & "  Debe = " & Abs(vFld(Rs("IVA")))
                  Q1 = Q1 & ", Haber = 0 "
               End If

               'Si tiene IVA Irrecuperable general, lo cambiamos por IVA Irrecuperable 1
               If TipoIVAIrrec = LIBCOMPRAS_IVAIRREC Or TipoIVAIrrec = LIBCOMPRAS_IVAIRREC1 Then    'TipoIVAIrrec = LIBCOMPRAS_IVAIRREC1 es para corregir un error
                  Q1 = Q1 & ", CodSIIDTE = 1"
                  CodSIIDTEIvaIrrec = 1
               End If
               
               Q1 = Q1 & ", Tasa = 100 "
               Q1 = Q1 & ", EsRecuperable = 0 "
               
               Q1 = Q1 & " WHERE IdMovDoc = " & IdMovIvaIrrec
               
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

               Call ExecSQL(DbMain, Q1)

               
            End If
            
            'eliminamos, si lo hubiera, un movimiento de IVA crédito fiscal
            If IdMovIvaCredFiscal <> 0 Or IdMovIvaActFijo <> 0 Then
               Wh = " WHERE IdMovDoc IN (" & IdMovIvaCredFiscal & "," & IdMovIvaActFijo & ")"
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'ffv delete
               Call DeleteSQL(DbMain, "MovDocumento", Wh)

            End If

         Case PIVA_PROP
         
            CalcPropIVA = gValPropIVA(vFld(Rs("MesDoc"))).CalcProp
            
            IVAIrrec = IVAIRREC_PARCIAL
            
            'asignamos el porcentaje correspondiente al IVA y el resto al IVA Irrecuperable (funciona para todos los casos, sea con proporción o todo irrecuperable o todo Crédito Fiscal)
            
            PropIVA = Abs(vFld(Rs("IVA"))) * gValPropIVA(vFld(Rs("MesDoc"))).Proporcion    'la Proporción está en 0 si es todo Irrecuperable o en 1 si es todo Crédito Fiscal
            PropIVA = Round(PropIVA, 0)
                           
            ValIVAIrrec = Abs(vFld(Rs("IVA"))) - PropIVA
            TasaIVAIrrec = gValPropIVA(vFld(Rs("MesDoc"))).Proporcion * 100
            
            If PropIVA = 0 Then
               'eliminamos, si lo hubiera, un movimiento de IVA crédito fiscal
               If IdMovIvaCredFiscal <> 0 Or IdMovIvaActFijo <> 0 Then
                  Wh = " WHERE IdMovDoc IN (" & IdMovIvaCredFiscal & "," & IdMovIvaActFijo & ")"
                  Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'ffv delete
                 Call DeleteSQL(DbMain, "MovDocumento", Wh)

               End If

            ElseIf IdMovIvaCredFiscal = 0 And IdMovIvaActFijo = 0 Then   'no hay registro de IVA Cred o IVA Act Fijo, insertamos uno de IVA Créd.
            
               orden = GetOrdenMovDoc(IdDoc)
               
               Q1 = QBase & IdDoc & "," & orden & ","             'IdDoc, Orden
               Q1 = Q1 & gCtasBas.IdCtaIVACred & ","              'IdCuenta
               
               If TipoDocNC Then
                  Q1 = Q1 & "0" & ","                             'Debe
                  Q1 = Q1 & PropIVA & ","                         'Haber
               Else
                  Q1 = Q1 & PropIVA & ","                         'Debe
                  Q1 = Q1 & "0" & ","                             'Haber
               End If
               
               Q1 = Q1 & "'" & Glosa & "',"                      'Glosa
               Q1 = Q1 & LIBCOMPRAS_IVACREDFISC & ","             'IdTipoValLib
                                 
               Q1 = Q1 & "0,0,0,0,0,''" & ")"                     'EsTotalDoc, IdCCosto, IdAreaNeg, Tasa, EsRecuperable, CodSIIDTE

               Call ExecSQL(DbMain, Q1)

               
            Else   'hay un registro de IVA Cred o de IVA Act Fijo => actualizamos el que hay
               
               Q1 = "UPDATE MovDocumento SET "

               If TipoDocNC Then
                  Q1 = Q1 & "  Debe = 0 "
                  Q1 = Q1 & ", Haber = " & PropIVA
               Else
                  Q1 = Q1 & "  Debe = " & PropIVA
                  Q1 = Q1 & ", Haber = 0 "
               End If

               If IdMovIvaCredFiscal <> 0 Then
                  Q1 = Q1 & " WHERE IdMovDoc = " & IdMovIvaCredFiscal
               Else
                  Q1 = Q1 & " WHERE IdMovDoc = " & IdMovIvaActFijo
               End If
               
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

               Call ExecSQL(DbMain, Q1)

               
            End If
            
            If ValIVAIrrec = 0 Then    'no hay IVA Irrecuperable
            
               'eliminamos, si lo hubiera, un movimiento de IVA irrecuperable
               If IdMovIvaIrrec > 0 Then
                  Wh = " WHERE IdMovDoc = " & IdMovIvaIrrec
                  Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
'ffv delete
                  Call DeleteSQL(DbMain, "MovDocumento", Wh)
                  CodSIIDTEIvaIrrec = 0
              End If
               
            'hay IVA Irrecuperable
            ElseIf IdMovIvaIrrec = 0 Then         'no existe el reg de IVA Irrecuperable, lo agregamos con IVA Irrecuperable 1
                              
               orden = GetOrdenMovDoc(IdDoc)
               
               Q1 = QBase & IdDoc & "," & orden & ","             'IdDoc, Orden
               Q1 = Q1 & gCtasBas.IdCtaIVAIrrec & ","             'IdCuenta
               
               If TipoDocNC Then
                  Q1 = Q1 & "0" & ","                             'Debe
                  Q1 = Q1 & ValIVAIrrec & ","                     'Haber
               Else
                  Q1 = Q1 & ValIVAIrrec & ","                     'Debe
                  Q1 = Q1 & "0" & ","                             'Haber
               End If
               
               Q1 = Q1 & "'" & Glosa & "',"                      'Glosa
               Q1 = Q1 & LIBCOMPRAS_IVAIRREC1 & ","               'IdTipoValLib
                                 
               Q1 = Q1 & "0,0,0," & str(TasaIVAIrrec) & ",0,'1'" & ")" 'EsTotalDoc, IdCCosto, IdAreaNeg, Tasa, EsRecuperable, CodSIIDTE

               Call ExecSQL(DbMain, Q1)

               CodSIIDTEIvaIrrec = 1
               
            Else
            
               'si está el movimento de IVA irrecuperable, lo actualizamos
               Q1 = "UPDATE MovDocumento SET "

               If TipoDocNC Then
                  Q1 = Q1 & "  Debe = 0 "
                  Q1 = Q1 & ", Haber = " & ValIVAIrrec
               Else
                  Q1 = Q1 & "  Debe = " & ValIVAIrrec
                  Q1 = Q1 & ", Haber = 0 "
               End If
               
               'Si tiene IVA Irrecuperable general, lo cambiamos por IVA Irrecuperable 1
               If TipoIVAIrrec = LIBCOMPRAS_IVAIRREC Or TipoIVAIrrec = LIBCOMPRAS_IVAIRREC1 Then    'TipoIVAIrrec = LIBCOMPRAS_IVAIRREC1 es para corregir un error
                  Q1 = Q1 & ", CodSIIDTE = 1"
                  CodSIIDTEIvaIrrec = 1
               End If
               
               Q1 = Q1 & ", Tasa = " & str(TasaIVAIrrec)
               Q1 = Q1 & ", EsRecuperable = 0 "

               Q1 = Q1 & " WHERE IdMovDoc = " & IdMovIvaIrrec
               
               Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

               Call ExecSQL(DbMain, Q1)

               
            End If
         
      End Select
      
      'actualizamos los campos
      Q1 = "UPDATE Documento SET IVAIrrecuperable = " & IVAIrrec & ", ValIVAIrrec = " & ValIVAIrrec & ", CodSIIDTEIVAIrrec = " & CodSIIDTEIvaIrrec & " WHERE IdDoc=" & IdDoc
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano

      Call ExecSQL(DbMain, Q1)

            
      'actualizamos la barra de proceso
      n = n + 1
      If Not ProgBar Is Nothing Then
         ProgBar.Value = n / nDocs * 100
      End If
      
      'Tracking 3227543
      Call SeguimientoDocumento(IdDoc, gEmpresa.id, gEmpresa.Ano, "Proporcionalidad.PropIVA_UpdateMovDoc", "", 1, "", gUsuario.IdUsuario, 1, 2)
      Call SeguimientoMovDocumento(IdDoc, gEmpresa.id, gEmpresa.Ano, "Proporcionalidad.PropIVA_UpdateMovDoc", "", 1, "", 1, 2)
      ' fin 3227543
      
      
      Rs.MoveNext
      
   Loop

   Call CloseRs(Rs)

   Msg = "RESULTADO: El cálculo de proporcionalidad fue aplicado a los siguientes documentos:" & vbCrLf & vbCrLf
   For i = 0 To UBound(nTipoDOc)
      If nTipoDOc(i) > 0 Then
         Msg = Msg & "     - " & nTipoDOc(i) & " " & gTipoDoc(i).Nombre & vbCrLf
      End If
   Next i

   MsgBox1 Msg, vbInformation
   
   PropIVA_UpdateMovDoc = True

End Function

Private Function GetOrdenMovDoc(ByVal IdDoc As Long) As Integer
   Dim Q1 As String
   Dim Rs As Recordset, RsAux As Recordset
   Dim orden As Integer
   
   GetOrdenMovDoc = 0

   If IdDoc = 0 Then
      Exit Function
   End If


   'obtenemos el Orden del último movimiento
   Q1 = "SELECT Max(Orden) FROM MovDocumento "
   Q1 = Q1 & " WHERE IdDoc = " & IdDoc
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   
   orden = 0
   
   If Not Rs.EOF Then
      orden = vFld(Rs(0))
   End If
   
   Call CloseRs(Rs)
   
   orden = orden + 1
   
   GetOrdenMovDoc = orden
   

End Function


