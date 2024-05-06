VERSION 5.00
Begin VB.Form FrmImportEmpresas 
   Caption         =   "Capturador de Empresas"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmImportEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Valida(ByVal vRut As String, ByVal vNCorto As String) As Boolean
   Dim Row As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Rut As String
   
   Valida = False
   
   Rut = Trim(vRut)
   
   If Rut = "" Then
      MsgBox1 "Debe ingresar el RUT de la empresa", vbExclamation
      Exit Function
   End If
   
   If Rut <> "" And Trim(vNCorto) = "" Then
      MsgBox1 "Debe ingresar nombre corto", vbExclamation
      Exit Function
   End If
   
   If gAppCode.Demo = True Then
      If Rut <> "1-9" And Rut <> "2-7" And Rut <> "3-5" Then
         MsgBox1 "En la versín DEMO sólo puede usar los siguietes RUTs:" & vbCrLf & vbCrLf & "1-9, 2-7 y 3-5", vbExclamation
         Exit Function
      End If
   End If
         
   
   'Ver en BDatos
   Q1 = "SELECT Rut FROM Empresas WHERE Rut='" & vFmtCID(vRut) & "'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False And lOper = OPER_NEW Then
      MsgBox1 "Ya existe una empresa con este RUT." & vRut, vbExclamation
      Call CloseRs(Rs)
      Exit Function
      
   End If
   Call CloseRs(Rs)
   
   'Ver en BDatos
   Q1 = "SELECT idEmpresa FROM Empresas WHERE NombreCorto='" & vNCorto & "'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF = False Then
      If (lOper = OPER_NEW) Or (lOper = OPER_EDIT And vFld(Rs("idEmpresa")) <> lId) Then
         MsgBox1 "Ya existe este nombre corto asociado a otra empresa " & vNCorto, vbExclamation
         Call CloseRs(Rs)
         Exit Function
      End If
   End If
   Call CloseRs(Rs)
   
   Valida = True
   
End Function

Private Function ImportFromFile() As Boolean
   Dim FName As String
   Dim Buf As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim ImpEnable As Boolean
   Dim IdEnt As Long
   Dim NotValidRut As Boolean
   Dim i As Integer, l As Integer
   Dim j As Integer, p As Long, k As Integer
   Dim NUpd As Long
   Dim NIns As Long
   Dim Rc As Integer
   Dim Fd As Long
   Dim Aux As String
   Dim vEstado As Integer
   Dim NRecErroneos As Integer, StrNRecErroneos As String
   Dim CampoInvalido As String
   Dim lId As String
   Dim vRut As String
   Dim vNomCorto As String
   
   
  
   Dim FldArray(3) As AdvTbAddNew_t
   
   ImportFromFile = False
      
   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = ""
   FrmMain.Cm_ComDlg.InitDir = gImportPath
   FrmMain.Cm_ComDlg.Filter = "Archivos de Texto (*.txt)|*.txt"
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
   If Err = cdlCancel Then
      Exit Function
   ElseIf Err Then
      MsgBox1 "Error " & Err & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Function
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Function
   End If
   Err.Clear
   
   FName = FrmMain.Cm_ComDlg.Filename
   
   MousePointer = vbHourglass
   DoEvents
      
   Rc = MsgBox1("Atención:" & vbNewLine & vbNewLine & "Se importará el archivo:" & vbNewLine & vbNewLine & FName & vbNewLine & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2)
   If Rc = vbNo Then
      Exit Function
   End If
   
   'abrimos el archivo
   Fd = FreeFile
   Open FName For Input As #Fd
   If Err Then
      MsgErr FName
      ImportFromFile = -Err
      Exit Function
   End If
   
   Row = i
   r = 0
   
   Grid.FlxGrid.Redraw = False
   
   Do Until EOF(Fd)
   
      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf
         
      p = 1
      Buf = Trim(Buf)
      

      '1er registro con nombres de campos
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 And InStr(1, Buf, "RUT", vbTextCompare) Then
         GoTo NextRec
      End If
      
      CampoInvalido = ""
      
      'Rut
      
       If Not ValidRut(Me.Txt_UsuarioSII.Text) Then
          CampoInvalido = CampoInvalido & "," & p
          Call AddLogImp(lFNameLogImp, FName, l, "Rut inválida.")
       Else
      
        Aux = Trim(NextField2(Buf, p, Sep))
        DtRec = vFmtCID(Aux, False)
        If DtRec = "0" Or DtRec = "" Then
           CampoInvalido = CampoInvalido & "," & p
           Call AddLogImp(lFNameLogImp, FName, l, "Rut inválida.")
        End If
      
      End If
      'Tipo Doc
      TipoDoc = Trim(NextField2(Buf, p))
      IdTipoDoc = FindTipoDoc(lTipoLib, TipoDoc)
      If IdTipoDoc = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, FName, l, "Tipo de documento inválido o no corresponde al libro en edición.")
      Else
         DocImpExp = CInt(gTipoDoc(GetTipoDoc(lTipoLib, IdTipoDoc)).DocImpExp)
         AceptaPropIVA = CInt(gTipoDoc(GetTipoDoc(lTipoLib, IdTipoDoc)).AceptaPropIVA)

      End If
         
      IdxTipoDoc = GetTipoDoc(lTipoLib, IdTipoDoc)
         
      'Del Giro
      If lTipoLib = LIB_VENTAS Then
         DelGiro = IIf(Trim(NextField2(Buf, p)) = "", 1, 0)
      End If
      
      
      'DTE
      TxtDTE = Trim(NextField2(Buf, p))
      DTE = IIf(Val(TxtDTE) = 0 Or Trim(TxtDTE) = "", 0, 1)
      
      If lTipoLib = LIB_VENTAS Then
         'N° Fiscal Impresora
         NumFiscImp = Trim(NextField2(Buf, p))
         
         'N° Informe Z
         NumInfZ = Trim(NextField2(Buf, p))
         
         If TipoDoc <> "MRG" And (NumFiscImp <> "" Or NumInfZ <> "") Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "N° Fiscal Impresora y/o N° Informe Z ser cero o blanco")
            NumFiscImp = ""
            NumInfZ = ""
         End If
      End If
      
      'NumDoc
      NumDoc = Trim(NextField2(Buf, p))
      If NumDoc = "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, FName, l, "N° de documento inválido.")
      End If
      
      If lTipoLib = LIB_VENTAS Then
         'NumDocHasta
         NumDocHasta = Trim(NextField2(Buf, p))
         If Val(NumDocHasta) <> 0 Then
         
            'NumDocHasta no se permite para VPE y otros documentos
'            If gTipoDoc(IdxTipoDoc).Diminutivo = "VPE" Then
            If gTipoDoc(IdxTipoDoc).TieneNumDocHasta = VAL_NOPERMITIDO Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, FName, l, "N° de Documento Hasta debe ser cero o blanco.")
         
            ElseIf Val(NumDocHasta) < Val(NumDoc) Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, FName, l, "N° de documento hasta inválido.")
            
'            ElseIf Val(NumDocHasta) = Val(NumDoc) Then
'               NumDocHasta = "0"
            End If
            
         ElseIf gTipoDoc(IdxTipoDoc).TieneNumDocHasta = VAL_OBLIGATORIO Then   'NumDocHasta = 0
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Num. Doc. Hasta debe ser mayor que cero.")
         
         End If
         
         'CantBoletas
         CantBoletas = Val(Trim(NextField2(Buf, p)))
         
         If CantBoletas > 0 Then
            If gTipoDoc(IdxTipoDoc).TieneCantBoletas = VAL_NOPERMITIDO Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, FName, l, "Cantidad de Boletas debe ser cero o blanco.")
            End If
         ElseIf gTipoDoc(IdxTipoDoc).TieneCantBoletas = VAL_OBLIGATORIO Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Cantidad de Boletas debe ser mayor que cero.")
         End If
      End If
      
      If lTipoLib = LIB_COMPRAS Then
      
         'Prop IVA
         Aux = Trim(NextField2(Buf, p))
         IdPropIVA = -1
         PropIVA = ""
         For k = 0 To UBound(gStrPropIVA)
            If Aux = Left(gStrPropIVA(k), 1) Then
               IdPropIVA = k
               PropIVA = Aux
            End If
         Next k
         
         If IdPropIVA < 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Opción Proporcionalidad de IVA inválida.")
         
         ElseIf Not AceptaPropIVA And IdPropIVA > 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Este tipo de documento no acepta Proporcionalidad de IVA.")
         End If
         
         
         'Fecha emisión
         Aux = Trim(NextField2(Buf, p))
         DtEmi = ValFmtDate(Aux, False)
         If DtEmi = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Fecha emisión inválida.")
         End If
      
      Else
         DtEmi = DtRec
         
      End If
      
      'Entidad
      IdEnt = 0
      NotValidRut = False
      Aux = Trim(NextField2(Buf, p))
      If Aux = "0-0" Or Aux = "" Then
         RutEnt = ""
      ElseIf Aux = "NULO" Then
         RutEnt = "NULO"
         Estado = ED_ANULADO
      Else
         RutEnt = vFmtCID(Aux)
         If RutEnt = "0" Or RutEnt = "" Then    'es inválido
            NotValidRut = True
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "RUT inválido")
         End If
      End If
      
      CodEnt = RutEnt
      
      NombEnt = Trim(NextField2(Buf, p))
      If NombEnt = "" And RutEnt <> "" Then
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, FName, l, "Falta ingresar nombre o razón social entidad.")
      End If
      
      Descrip = Trim(NextField2(Buf, p))
      If Descrip = "NULO" Then
         Estado = ED_ANULADO
      End If
      CodSuc = Trim(NextField2(Buf, p))
      IdSucursal = 0
      Sucursal = ""
      
      If CodSuc <> "" Then
         Q1 = "SELECT IdSucursal, Descripcion FROM Sucursales WHERE Codigo ='" & CodSuc & "'"
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Código de sucursal inválido")
         Else
            IdSucursal = vFld(Rs("IdSucursal"))
            Sucursal = vFld(Rs("Descripcion"))
         End If
         
         Call CloseRs(Rs)
      End If
      
      'Valores
      Afecto = vFmt(Trim(NextField2(Buf, p)))
      'código cuenta Afecto
      AuxCodCtaAfecto = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      
      'Afecto sólo si corresponde
      If (Afecto <> 0 Or AuxCodCtaAfecto <> "") And gTipoDoc(IdxTipoDoc).TieneAfecto = VAL_NOPERMITIDO Then
         CampoInvalido = CampoInvalido & "," & "Exento"
         Call AddLogImp(lFNameLogImp, FName, l, "Valor Afecto o Código de Cuenta Afecto no permitido.")
      End If
      
      'Exento
      Exento = vFmt(Trim(NextField2(Buf, p)))
      
      'código cuenta Exento
      AuxCodCtaExento = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      
      'Exento sólo si corresponde
      If (Exento <> 0 Or AuxCodCtaExento <> "") And gTipoDoc(IdxTipoDoc).TieneExento = VAL_NOPERMITIDO Then
         CampoInvalido = CampoInvalido & "," & "Exento"
         Call AddLogImp(lFNameLogImp, FName, l, "Valor Exento o Código de Cuenta Exento no permitido.")
      End If
      
      IVA = vFmt(Trim(NextField2(Buf, p)))
      
      OtroImp = vFmt(Trim(NextField2(Buf, p)))
      'código cuenta OtrosImp
      AuxCodCtaOtroImp = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      
      If TipoDoc = "VPE" And OtroImp <> 0 Then
         CampoInvalido = CampoInvalido & "," & "Otro Impuesto"
         Call AddLogImp(lFNameLogImp, FName, l, "Valor de Otro Impuesto inválido.")
      End If

      Total = vFmt(Trim(NextField2(Buf, p)))
      'código cuenta Total
      AuxCodCtaTotal = VFmtCodigoCta(Trim(NextField2(Buf, p)))
      
      If gTipoDoc(IdxTipoDoc).EsRebaja Then
         Afecto = Abs(Afecto) * -1
         Exento = Abs(Exento) * -1
         IVA = Abs(IVA) * -1
         Total = Abs(Total) * -1
         OtroImp = OtroImp * -1   'para permitir ingresar valores negativos en los otros impuestos. Al ser rebaja y el usuario ingresa un valor negativo, queda positivo
      
      Else
         If Afecto < 0 Or Exento < 0 Or IVA < 0 Or Total < 0 Then
            CampoInvalido = CampoInvalido & "," & "Afecto, Exento, IVA, Total"
            Call AddLogImp(lFNameLogImp, FName, l, "Valor de Afecto, Exento, IVA y/o Total inválido.")
         End If
                     
      End If
      
      If lTipoLib = LIB_VENTAS Then
         'Ventas Acum Informe Z
         VentasAcumInfZ = vFmt(Trim(NextField2(Buf, p)))
                  
         If TipoDoc <> "MRG" And VentasAcumInfZ <> 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Ventas Acum. Informe Z inválido")
            VentasAcumInfZ = 0
         End If
         
      End If

      
      'Fecha Vencim
      Aux = Trim(NextField2(Buf, p))
      DtVenc = 0
      If Aux <> "" Then
         DtVenc = ValFmtDate(Aux, False)
         If DtVenc = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Fecha vencimiento inválida.")
         End If
      Else
         CampoInvalido = CampoInvalido & "," & p
         Call AddLogImp(lFNameLogImp, FName, l, "Falta ingresar Fecha Vencimiento.")
      End If
            
      NumInterno = vFmt(Trim(NextField2(Buf, p)))
      CodANeg = Trim(NextField2(Buf, p))
      CodCCosto = Trim(NextField2(Buf, p))
      
      IdANeg = GetAreaNegocio(CodANeg)
      IdCCosto = GetCentroCosto(CodCCosto)
      
      If CodANeg <> "" And IdANeg = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         pCCosto = p
         Call AddLogImp(lFNameLogImp, FName, l, "Área de Negocio inválida.")
      End If
           
      If CodCCosto <> "" And IdCCosto = 0 Then
         CampoInvalido = CampoInvalido & "," & p
         pANeg = p
         Call AddLogImp(lFNameLogImp, FName, l, "Centro de Gestión inválido.")
      End If
      
      'códigos cuentas
      
      NomCta = ""
      
      If AuxCodCtaAfecto <> "" Then
         If Afecto = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Código de cuenta Afecto debe ser cero o blanco")
         
         Else
            AuxIdCtaAfecto = GetIdCuenta(NomCta, AuxCodCtaAfecto, AuxDescCtaAfecto, UltNivel)
            If AuxIdCtaAfecto <= 0 Or Not UltNivel Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, FName, l, "Código de cuenta Afecto inválido")
            End If
         End If
      Else
         AuxIdCtaAfecto = 0
         AuxDescCtaAfecto = ""
      End If
      
      NomCta = ""
      
      If AuxCodCtaExento <> "" Then
         If Exento = 0 Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Código de cuenta Exento debe ser cero o blanco")
         Else
            AuxIdCtaExento = GetIdCuenta(NomCta, AuxCodCtaExento, AuxDescCtaExento, UltNivel)
            If AuxIdCtaExento <= 0 Or Not UltNivel Then
               CampoInvalido = CampoInvalido & "," & p
               Call AddLogImp(lFNameLogImp, FName, l, "Código de cuenta Exento inválido")
            End If
         End If
      Else
         AuxIdCtaExento = 0
         AuxDescCtaExento = ""
      End If
            
      NomCta = ""
      
      If AuxCodCtaTotal <> "" Then
         AuxIdCtaTotal = GetIdCuenta(NomCta, AuxCodCtaTotal, AuxDescCtaTotal, UltNivel)
         If AuxIdCtaTotal <= 0 Or Not UltNivel Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Código de cuenta Total inválido")
         End If
      Else
         AuxIdCtaTotal = 0
         AuxDescCtaTotal = ""
      End If
            
      NomCta = ""
    
      If AuxCodCtaOtroImp <> "" Then
         AuxIdCtaOtroImp = GetIdCuenta(NomCta, AuxCodCtaOtroImp, AuxDescCtaOtroImp, UltNivel)
         If AuxIdCtaOtroImp <= 0 Or Not UltNivel Then
            CampoInvalido = CampoInvalido & "," & p
            Call AddLogImp(lFNameLogImp, FName, l, "Código de cuenta Otros Impuestos inválido")
         End If
      Else
         AuxIdCtaOtroImp = 0
         AuxDescCtaOtroImp = ""
      End If
            
      NomCta = ""
    
      'Cuentas Contables Default
                  
      If Exento <> 0 Then
         If AuxIdCtaExento > 0 Then
            IdCtaExento = AuxIdCtaExento
            CodCtaExento = FmtCodCuenta(AuxCodCtaExento)
            DescCtaExento = AuxDescCtaExento
         Else
            IdCtaExento = lCtaExento.id
            CodCtaExento = FmtCodCuenta(lCtaExento.Codigo)
            DescCtaExento = lCtaExento.Descripcion
         End If
      Else
         IdCtaExento = 0
         CodCtaExento = ""
         DescCtaExento = ""
      End If
         
      
      If Afecto <> 0 Then
         If AuxIdCtaAfecto > 0 Then
            IdCtaAfecto = AuxIdCtaAfecto
            CodCtaAfecto = FmtCodCuenta(AuxCodCtaAfecto)
            DescCtaAfecto = AuxDescCtaAfecto
         Else
            IdCtaAfecto = lCtaAfecto.id
            CodCtaAfecto = FmtCodCuenta(lCtaAfecto.Codigo)
            DescCtaAfecto = lCtaAfecto.Descripcion
         End If
      Else
         IdCtaAfecto = 0
         CodCtaAfecto = 0
         DescCtaAfecto = 0
      End If
                     
                  
      If AuxIdCtaTotal > 0 Then
         IdCtaTotal = AuxIdCtaTotal
         CodCtaTotal = FmtCodCuenta(AuxCodCtaTotal)
         DescCtaTotal = AuxDescCtaTotal
      Else
         IdCtaTotal = lCtaTotal.id
         CodCtaTotal = FmtCodCuenta(lCtaTotal.Codigo)
         DescCtaTotal = lCtaTotal.Descripcion
      End If
            
      If IVA <> 0 Then
         IdCtaIVA = lIdCuentaIVA
      Else
         IdCtaIVA = 0
      End If
      
      If OtroImp <> 0 Then
         If AuxIdCtaOtroImp > 0 Then
            IdCtaOtroImp = AuxIdCtaOtroImp
'            CodCtaOtroImp = FmtCodCuenta(AuxCodCtaOtroImp)
'            DescCtaOtroImp = AuxDescCtaOtroImp
         Else
            'si es factura de compras, nota de crédito de fac. compras o nota de débito de fac. compras, se pone la cuenta al revés
            If TipoDoc = "FCC" Or TipoDoc = "NCF" Or TipoDoc = "NDF" Or TipoDoc = "FCV" Then
               IdCtaOtroImp = lIdCuentaOtrosImpFacCompra
            Else
               IdCtaOtroImp = lIdCuentaOtrosImp
            End If
         End If
      Else
         IdCtaOtroImp = 0
      End If
      
      'validamos si ingresó area de negocio y centro de costo si corresponde
      AtribANeg = GetAtribCuenta(IdCtaAfecto, ATRIB_AREANEG) Or GetAtribCuenta(IdCtaExento, ATRIB_AREANEG) Or GetAtribCuenta(IdCtaTotal, ATRIB_AREANEG)
      
      AtribCCosto = GetAtribCuenta(IdCtaAfecto, ATRIB_CCOSTO) Or GetAtribCuenta(IdCtaExento, ATRIB_CCOSTO) Or GetAtribCuenta(IdCtaTotal, ATRIB_CCOSTO)
      
      If AtribANeg And IdANeg = 0 Then
         CampoInvalido = CampoInvalido & "," & pANeg
         Call AddLogImp(lFNameLogImp, FName, l, "Falta indicar Área de Negocio")
      End If
      
      If AtribCCosto And IdCCosto = 0 Then
         CampoInvalido = CampoInvalido & "," & pCCosto
         Call AddLogImp(lFNameLogImp, FName, l, "Falta indicar Centro de Costo")
      End If
      
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
            
            If lId = 0 Then  'no existe
         
                FldArray(0).FldName = "Rut"
                FldArray(0).FldValue = vFmtCID(Tx_RUT)
                FldArray(0).FldIsNum = False
                
                FldArray(1).FldName = "NombreCorto"
                FldArray(1).FldValue = ParaSQL(Tx_NCorto)
                FldArray(1).FldIsNum = False
                            
                lId = AdvTbAddNewMult(DbMain, "Empresas", "IdEmpresa", FldArray)

            End If
            
         End If
         
         'para que no ingrese documento que ya existen en la grilla
         
         Q1 = "SELECT IdDoc FROM Documento "
         Q1 = Q1 & " WHERE TipoLib = " & lTipoLib & " AND TipoDoc = " & IdTipoDoc
         Q1 = Q1 & " AND NumDoc = '" & NumDoc & "'"
         Q1 = Q1 & " AND IdEntidad = " & IdEnt
         Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         
         Set Rs = OpenRs(DbMain, Q1)
         
         If Rs.EOF = True Then 'documento no existe
      
         'si no hay errores, ingresamos el registro a la grilla
         Grid.TextMatrix(Row, C_NUMLIN) = vFmt(Grid.TextMatrix(Row - 1, C_NUMLIN)) + 1
         Grid.TextMatrix(Row, C_FECHA) = Day(DtRec)
         Grid.TextMatrix(Row, C_IDTIPODOC) = IdTipoDoc
         Grid.TextMatrix(Row, C_TIPODOC) = TipoDoc
         Grid.TextMatrix(Row, C_DOCIMPEXP) = DocImpExp
         
         If lTipoLib = LIB_VENTAS Then
            Grid.TextMatrix(Row, C_GIRO) = IIf(DelGiro = 0, "No", "")
         Else
            Grid.TextMatrix(Row, C_GIRO) = ""
         End If
         
         Grid.TextMatrix(Row, C_DTE) = IIf(DTE <> 0, "x", "")
                  
         If lTipoLib = LIB_VENTAS Then
            Grid.TextMatrix(Row, C_NUMFISCIMPR) = NumFiscImp
            Grid.TextMatrix(Row, C_NUMINFORMEZ) = NumInfZ
            Grid.TextMatrix(Row, C_VENTASACUM) = Format(VentasAcumInfZ, NUMFMT)
         End If
         
         Grid.TextMatrix(Row, C_NUMDOC) = NumDoc
         Grid.TextMatrix(Row, C_NUMDOCHASTA) = NumDocHasta
         
         
         If lTipoLib = LIB_VENTAS Then
            If Val(NumDocHasta) > 0 And Val(NumDocHasta) >= Val(NumDoc) Then
               Grid.TextMatrix(Row, C_CANTBOLETAS) = Format(Val(NumDocHasta) - Val(NumDoc), NUMFMT) + 1
            ElseIf CantBoletas > 0 Then
               Grid.TextMatrix(Row, C_CANTBOLETAS) = Format(CantBoletas, NUMFMT)
            Else
               Grid.TextMatrix(Row, C_CANTBOLETAS) = ""
            End If
         End If
         
         If lTipoLib = LIB_COMPRAS Then
            Grid.TextMatrix(Row, C_IDPROPIVA) = IdPropIVA
            Grid.TextMatrix(Row, C_PROPIVA) = PropIVA
         End If
         
         Grid.TextMatrix(Row, C_FECHAEMIORI) = Format(DtEmi, SDATEFMT)
         Grid.TextMatrix(Row, C_LNGFECHAEMIORI) = DtEmi
         Grid.TextMatrix(Row, C_RUT) = FmtRut(RutEnt)
         Grid.TextMatrix(Row, C_NOMBRE) = NombEnt
         Grid.TextMatrix(Row, C_IDENTIDAD) = IdEnt
         Grid.TextMatrix(Row, C_DESCRIP) = Descrip
         Grid.TextMatrix(Row, C_IDSUCURSAL) = IdSucursal
         Grid.TextMatrix(Row, C_SUCURSAL) = Sucursal
         
         Grid.TextMatrix(Row, C_AFECTO) = Format(Afecto, NUMFMT)
         Grid.TextMatrix(Row, C_AF_IDCUENTA) = IdCtaAfecto
         Grid.TextMatrix(Row, C_AF_CODCUENTA) = CodCtaAfecto
         Grid.TextMatrix(Row, C_AF_CUENTA) = DescCtaAfecto
         
         Grid.TextMatrix(Row, C_EXENTO) = Format(Exento, NUMFMT)
         Grid.TextMatrix(Row, C_EX_IDCUENTA) = IdCtaExento
         Grid.TextMatrix(Row, C_EX_CODCUENTA) = CodCtaExento
         Grid.TextMatrix(Row, C_EX_CUENTA) = DescCtaExento
         
         Grid.TextMatrix(Row, C_IVA) = Format(IVA, NUMFMT)
         Grid.TextMatrix(Row, C_IVA_IDCUENTA) = IdCtaIVA
         
         Grid.TextMatrix(Row, C_OTROIMP) = Format(OtroImp, NUMFMT)
         Grid.TextMatrix(Row, C_OIMP_IDCUENTA) = IdCtaOtroImp
         
         Grid.TextMatrix(Row, C_TOTAL) = Format(Total, NUMFMT)
         Grid.TextMatrix(Row, C_TOT_IDCUENTA) = IdCtaTotal
         Grid.TextMatrix(Row, C_TOT_CODCUENTA) = CodCtaTotal
         Grid.TextMatrix(Row, C_TOT_CUENTA) = DescCtaTotal
         
         If IdANeg > 0 Or IdCCosto > 0 Then
            Grid.TextMatrix(Row, C_IDANEG_CCOSTO) = IdANeg & "-" & IdCCosto
         End If
         
         Grid.TextMatrix(Row, C_DETALLE) = TX_DETALLE
         Grid.TextMatrix(Row, C_FECHAVENC) = IIf(DtVenc <> 0, Format(DtVenc, SDATEFMT), "")
         Grid.TextMatrix(Row, C_LNGFECHAVENC) = DtVenc
         Grid.TextMatrix(Row, C_CORRINTERNO) = NumInterno
         Grid.TextMatrix(Row, C_DETACTFIJO) = TX_ACTFIJO
         Grid.TextMatrix(Row, C_ESTADO) = gEstadoDoc(Estado)
         Grid.TextMatrix(Row, C_IDESTADO) = Estado
         Grid.TextMatrix(Row, C_USUARIO) = gUsuario.Nombre
         Grid.TextMatrix(Row, C_UPDATE) = FGR_I
                  
                 
         If EsIngresoTotal(Row) Then
            Dim Value As String
            Value = Total
            Call CalcIngresoTotal(Row, C_TOTAL, Value)
            '3071158
            'Call CalcTot
            '3071158

         Else
            Call CalcTotRow(Row, False)   'no recalcula IVA, deja el que viene
            
         End If
         
         
         Row = Row + 1
         
         r = r + 1
         
        '3071158
         Call CloseRs(Rs)
         
          End If
        '3071158
         
         If gDbType = SQL_ACCESS Then
            If r = 3000 Then
                Exit Do
            End If
         Else
            If r = 3000 Then
                Mayor3000reg = True
                Exit Do
            End If
         End If
      
         Grid.rows = Grid.rows + 1
         
'         If FGrChkMaxSize(Grid) = True Then
'            Exit loop
'         End If
    
   

      Else
         NRecErroneos = NRecErroneos + 1
         
         
      End If
      
NextRec:
   Loop

   Close #Fd
   
   '3071158
   Call CalcTot
   '3071158
   
   Grid.FlxGrid.Redraw = True
   
   Me.MousePointer = vbDefault
   
   If NRecErroneos = 0 Then
      If r = 1 Then
         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregó " & r & " documento.", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      ElseIf r > 1 And Mayor3000reg = False Then
         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos.", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      ElseIf r > 1 And Mayor3000reg Then
         MsgBox1 "Importación finalizada con éxito. Resultado:" & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos, Si desea importar una mayor cantidad debera hacer una captura mediante Registro de Ventas SII (CSV)", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
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
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      ElseIf r > 1 And Mayor3000reg = False Then
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos.", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      ElseIf r > 1 And Mayor3000reg Then
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- Se agregaron " & r & " documentos, Si desea importar una mayor cantidad debera hacer una captura mediante Registro de Ventas SII (CSV)", vbInformation + vbOKOnly
         MsgBox1 "Si desea cancelar la importación, presione el botón Cancelar.", vbInformation + vbOKOnly
      Else  ' r=0
         MsgBox1 "Importación finalizada. Resultado:" & vbNewLine & vbNewLine & StrNRecErroneos & vbNewLine & vbNewLine & "- No se agregaron documentos.", vbInformation + vbOKOnly
      End If
      
      If MsgBox1("¿Desea revisar el log de importación " & lFNameLogImp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Me.hWnd, "open", lFNameLogImp, "", "", SW_SHOW)
      End If
   End If


   ImportFromFile = True
   
End Function
