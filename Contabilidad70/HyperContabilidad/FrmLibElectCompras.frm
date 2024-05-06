VERSION 5.00
Begin VB.Form FrmLibElectCompras 
   Caption         =   "Generar Libro de Electrónico de Compras"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6435
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Manual_IEC 
      Caption         =   "Manual IEC"
      Height          =   795
      Left            =   4740
      Picture         =   "FrmLibElectCompras.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Ver Manual Libro Electrónico de Compras"
      Top             =   1320
      Width           =   1275
   End
   Begin VB.TextBox Tx_Path 
      BackColor       =   &H8000000F&
      Height          =   675
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   4020
      Width           =   4815
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   4740
      TabIndex        =   3
      Top             =   780
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   360
      Picture         =   "FrmLibElectCompras.frx":06B6
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   9
      Top             =   420
      Width           =   585
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Index           =   0
      Left            =   1200
      TabIndex        =   8
      Top             =   300
      Width           =   3075
      Begin VB.OptionButton Op_Libros 
         Caption         =   "Libro de Compras"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   0
         Top             =   420
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame Fr_Periodo 
      Caption         =   "Período"
      Height          =   1095
      Left            =   1200
      TabIndex        =   4
      Top             =   2340
      Width           =   4815
      Begin VB.TextBox Tx_Ano 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3300
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   420
         Width           =   795
      End
      Begin VB.ComboBox Cb_Mes 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   7
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   6
         Top             =   480
         Width           =   345
      End
   End
   Begin VB.CommandButton Bt_Export 
      Caption         =   "Exportar"
      Height          =   315
      Left            =   4740
      TabIndex        =   2
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Exportar a:"
      Height          =   195
      Left            =   1200
      TabIndex        =   10
      Top             =   3780
      Width           =   1515
   End
End
Attribute VB_Name = "FrmLibElectCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lCamposLibElectCompras As String
Dim lCamposLibComprasAcepta As String

Dim lSep As String

Dim lTipoExport As String

Public Sub FGenLibComprasSII()
   lTipoExport = "SII"
   Me.Show vbModal
   
End Sub

Public Sub FGenLibComprasAcepta()
   lTipoExport = "ACEPTA"
   Me.Show vbModal
   
End Sub


Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Export_Click()
   
   If Op_Libros(LIB_COMPRAS) <> 0 Then
      If lTipoExport = "SII" Then
         Call ExportLibElectCompras
      ElseIf lTipoExport = "ACEPTA" Then
         Call ExportLibComprasAcepta
      End If
   
'   ElseIf Op_Libros(LIB_VENTAS) <> 0 Then
'      Call ExportLibSII(LIB_VENTAS)
      
   End If
   
End Sub

Private Sub Bt_Manual_IEC_Click()
   Dim Rc As Long
   Dim Buf As String
   
   MousePointer = vbHourglass
   DoEvents
   
   Buf = gAppPath & "\Manual_Importación_IEC.pdf"
   Rc = ExistFile(Buf)
      
   If Rc = 0 Then
      MsgBox1 "No se encontró el archivo que contiene el Manual del Libro Electrónico de Compras, por favor contáctese con su proveedor para obtenerlo.", vbExclamation
   Else

      Rc = ShellExecute(Me.hWnd, "open", Buf, "", "", 1)
      If Rc < 32 Then
         MsgBox1 "Error " & Rc & " al abrir el archivo '" & Buf & "' que contiene el Manual del Libro Electrónico de Compras." & vbLf & "Trate de abrir este archivo con otro programa.", vbExclamation
      End If
   End If

   MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

   MsgBox1 "Recuerde que a contar de Agosto 2017 entra en vigencia el nuevo Registro de Compras y Ventas. Para mayor información puede revisar las Res. Ex. SII N°s 61 y 68, ambas de 2017.", vbInformation
   
End Sub

Private Sub Form_Load()

   Call FillMes(Cb_Mes, GetMesActual())

   Tx_Ano = gEmpresa.Ano
   
   If lTipoExport = "SII" Then
      Tx_Path = gExportPath & "\SII\" & gEmpresa.Rut
   ElseIf lTipoExport = "ACEPTA" Then
      Tx_Path = gExportPath & "\Factura\" & gEmpresa.Rut
      Bt_Manual_IEC.Visible = False
   End If
   
   If lTipoExport = "SII" Then
      Me.Caption = "Generar Libro Electrónico de Compras"
   ElseIf lTipoExport = "ACEPTA" Then
      Me.Caption = "Exportar Libro de Compras para Facturación Electrónica"
   End If


End Sub

Private Function ExportLibElectCompras() As Boolean
   Dim FirstDay As Long, LastDay As Long
   Dim FPath As String
   Dim SFName As String
   Dim Fd As Long
   Dim MesAno As String, AnoMes As String
   Dim TipoOperacion As String
   Dim TxtFields As String
   Dim Q1 As String, Q2 As String
   Dim Rs As Recordset, RsImp As Recordset
   Dim TxtLine As String
   Dim nDocs As Integer
   Dim AddReg As Boolean
   Dim i As Integer, r As Integer
   Dim MsgTipoDoc As Boolean
   Dim TipoLib As Integer
   Dim TipoDocSII As String
   Dim RegFijo1 As String, RegFijo2
   Dim RegVar As String, RegVarVacio As String
   Dim IdDoc As Long
   Dim ImpAdicRec(MAX_IMPADICDOC) As ImpAdicDoc_t
   Dim IdxValLib As Integer, IdxTipoDoc As Integer
   Dim IVAActFijo As Double, IVARet As Double, IVANoRet As Double, IVA As Double, ImpNoRec As Double
   Dim HayImpAdic As Boolean
   Dim DimTipoDocAsoc As String
   Dim EmisorReceptor As String
   Dim ActFijo As Double
   Dim lFNameLogExp As String
   Dim MsgIdDoc As String
   Dim NErrores As Integer
   Dim EsImpCoDEspecial As Boolean
   Dim DimDoc As String
   Dim LogPath As String
   Dim RegFijo1Ext As String
   Dim RegFijo1ExtSinIVANoRec As String
   Dim Diminutivo As String
   Dim IVAUsoComun As Double
   
   TipoLib = LIB_COMPRAS
   Call FirstLastMonthDay(DateSerial(Val(Tx_Ano), CbItemData(Cb_Mes), 1), FirstDay, LastDay)
   MesAno = Right("0" & CbItemData(Cb_Mes), 2) & Tx_Ano
   AnoMes = Tx_Ano & Right("0" & CbItemData(Cb_Mes), 2)
   
   If DateSerial(Val(Tx_Ano), CbItemData(Cb_Mes), 1) >= DateSerial(2018, 2, 1) Then    'a partir de Febrero 2018 no se genera este libro
      MsgBox1 "A contar del periodo Febrero 2018 todos los contribuyentes deben utilizar el Registro de Compras y Ventas de acuerdo a las Res. Ex. N°s 61 y 68, ambas del 2017, por lo que no es posible generar el archivo para el mes y año seleccionado.", vbInformation
      Exit Function
   End If
   
   'vemos si hay documentos excluídos del libro
   Q1 = "SELECT Count(*) "
   Q1 = Q1 & " FROM Documento INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDOC = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib & " AND Documento.Estado <> " & ED_ANULADO & " AND  ( TipoDocs.CodDocSII IS NULL OR TipoDocs.CodDocSII = '') AND (TipoDocs.CodDocDTESII IS NULL OR TipoDocs.CodDocDTESII = '' ) AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      If vFld(Rs(0)) > 0 Then
         MsgBox1 "Recuerde que los siguientes documentos no están considerados en el Informe Electrónico de Compras:" & vbCrLf & vbCrLf & "1. Form. Importaciones Exento (IEX)" & vbCrLf & vbCrLf & "2. Otros Documentos (OTC)" & vbCrLf & vbCrLf & "3. Crédito Impto. Timbres y Estam. (CIT)", vbInformation
      End If
   End If
   Call CloseRs(Rs)


   lSep = ";"
   lCamposLibElectCompras = "Tipo Doc" & lSep & "Folio" & lSep & "Rut Contraparte" & lSep & "Tasa Impuesto" & lSep & "Razón Social Contraparte" & lSep & "Tipo Impuesto [1=IVA:2=LEY 18211]" & lSep & "Fecha Emisión" & lSep & "Anulado[A]" & lSep & "Monto Exento" & lSep & "Monto Neto" & lSep & "Monto IVA (Recuperable)" & lSep & "Cod IVA no Rec" & lSep & "Monto IVA no Rec" & lSep & "IVA Uso Común" & lSep & "Cod Otro Imp (Con Crédito)" & lSep & "Tasa Otro Imp (Con Crédito)" & lSep & "Monto Otro Imp (Con Crédito)" & lSep & "Monto Total" & lSep & "Monto Otro Imp Sin Crédito" & lSep & "Monto Activo Fijo" & lSep & "Monto IVA Activo Fijo" & lSep & "IVA No Retenido" & lSep & "Tabacos - Puros" & lSep & "Tabacos - Cigarrillos" & lSep & "Tabacos - Elaborados" & lSep & "Impuesto a Vehiculos Automóviles" & lSep & "Codigo sucursal SII" & lSep & "Numero Interno" & lSep & "Emisor/Receptor"
         
   TipoOperacion = "Com-"
   
   On Error Resume Next
   
   FPath = gExportPath & "\SII\"
   MkDir FPath
   FPath = gExportPath & "\SII\" & gEmpresa.Rut
   MkDir FPath
   
   LogPath = FPath & "\Log"
   MkDir LogPath
   
   On Error GoTo 0
      
   SFName = FPath & "\LE_" & TipoOperacion & AnoMes & ".csv"
   lFNameLogExp = LogPath & "\LE_Com-" & Format(Now, "yyyymmdd") & ".log"
         
   Fd = FreeFile()
   Open SFName For Output As #Fd
   If Err Then
      MsgErr "Error al abrir el archivo '" & SFName & "'.", vbExclamation
      Exit Function
   End If
      
   'ponemos códigos de campos en primera línea
   TxtFields = lCamposLibElectCompras
   Print #Fd, TxtFields
   

   'Consulta para obtener cada documento y su detalle
   Q1 = "SELECT IdDoc, Documento.TipoDoc, DTE, TipoDocs.CodDocSII, TipoDocs.CodDocDTESII, Documento.Giro, NumDoc, CorrInterno, NumDocHasta, Documento.MovEdited, "
   Q1 = Q1 & " Documento.IdEntidad, Documento.RutEntidad, Documento.NombreEntidad, Entidades.Rut, Entidades.NotValidRut, Entidades.Nombre,"
   Q1 = Q1 & " FEmision, FEmisionOri, FVenc, Exento, Afecto, IVA, IVAIrrecuperable, ValIVAIrrec, CodSIIDTEIVAIrrec, "
   Q1 = Q1 & " OtroImp, OtrosVal, Total, Descrip, Documento.Estado, Documento.TipoDocAsoc, Documento.IVAActFijo "
   Q1 = Q1 & " FROM ( Documento INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDoc = TipoDocs.TipoDoc) "
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad "
   Q1 = Q1 & " AND Entidades.IdEmpresa = Documento.IdEmpresa "
   Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib & " AND Documento.Estado <> " & ED_ANULADO
   Q1 = Q1 & " AND ( TipoDocs.CodDocSII <> '' OR TipoDocs.CodDocDTESII <> '' ) AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY IdDoc"
      
   
   Set Rs = OpenRs(DbMain, Q1)
   
   TxtLine = ""
   nDocs = 0
   NErrores = 0
   
   Do While Not Rs.EOF
   
      AddReg = True
      IdDoc = vFld(Rs("IdDoc"))
      Diminutivo = gTipoDoc(GetTipoDoc(LIB_COMPRAS, vFld(Rs("TipoDoc")))).Diminutivo
            
      If vFld(Rs("DTE")) Then
         TipoDocSII = vFld(Rs("CodDocDTESII"))
      Else
         TipoDocSII = vFld(Rs("CodDocSII"))
      End If
      
      'Tipo Doc,Folio
      TxtLine = TipoDocSII & lSep & vFld(Rs("NumDoc"))
      
      'Rut Contraparte
      TxtLine = TxtLine & lSep & vFld(Rs("Rut")) & "-" & DV_Rut(vFld(Rs("Rut")))
      
      MsgIdDoc = "DOC.: " & Diminutivo & "-" & vFld(Rs("NumDoc")) & " - RUT: " & FmtRut(gEmpresa.Rut)
      
      'Tasa IVA. En esta columna, siempre se registra 19.
      TxtLine = TxtLine & lSep & gIVA * 100
      
      'Razón Social Contraparte
      TxtLine = TxtLine & lSep & FilterChrSII(vFld(Rs("Nombre")))
      
      'Tipo Impuesto [1=IVA:2=LEY 18211]. Siempre debe ir en esta columna 1
      TxtLine = TxtLine & lSep & "1"
      
      'Fecha Emisión
      TxtLine = TxtLine & lSep & Format(vFld(Rs("FEmisionOri")), "dd-mm-yyyy")
      
      'Anulado[A]
      If vFld(Rs("Estado")) = ED_ANULADO Then
         TxtLine = TxtLine & lSep & "A"
      Else
         TxtLine = TxtLine & lSep
      End If
      
      'Monto Exento
      If vFld(Rs("Exento")) > 0 Then
         TxtLine = TxtLine & lSep & vFld(Rs("Exento"))
      Else
         TxtLine = TxtLine & lSep
      End If
      
      'Monto Neto
      If vFld(Rs("Afecto")) > 0 Then
         TxtLine = TxtLine & lSep & vFld(Rs("Afecto"))
      Else
         TxtLine = TxtLine & lSep
      End If
      
      
      
      '********  Obtenermos los impruestos adicionales e IVA activo fijo
      For r = 0 To UBound(ImpAdicRec)
         ImpAdicRec(r).IdTipoValLib = 0
         ImpAdicRec(r).Valor = 0
         ImpAdicRec(r).Tasa = 0
      Next r
      
      Q2 = "SELECT MovDocumento.IdTipoValLib, MovDocumento.Tasa, MovDocumento.EsRecuperable, Sum(Debe) as SumDebe, Sum(Haber) as SumHaber "
      Q2 = Q2 & " FROM MovDocumento INNER JOIN TipoValor ON MovDocumento.IdTipoValLib = TipoValor.Codigo "
      Q2 = Q2 & " WHERE IdDoc = " & IdDoc & " AND TipoValor.TipoLib = " & LIB_COMPRAS & " AND IdTipoValLib >= " & LIBCOMPRAS_OTROSIMP & " AND Atributo NOT IN ('IVAIRREC', 'SINUSO')"
      Q1 = Q1 & " AND MovDocumento.IdEmpresa = " & gEmpresa.id & " AND MovDocumento.Ano = " & gEmpresa.Ano
      Q2 = Q2 & " GROUP BY MovDocumento.IdTipoValLib, MovDocumento.Tasa, MovDocumento.EsRecuperable"
      
      Set RsImp = OpenRs(DbMain, Q2)
            
      r = 0
      IVAActFijo = 0
      IVARet = 0
      ImpNoRec = 0
      
      Do While Not RsImp.EOF
      
         IdxValLib = GetTipoValLib(LIB_COMPRAS, vFld(RsImp("IdTipoValLib")))
         
         If vFld(RsImp("IdTipoValLib")) = LIBCOMPRAS_IVAACTFIJO Then
            IVAActFijo = Abs(vFld(RsImp("SumDebe")) - vFld(RsImp("SumHaber")))

         Else
         
            If gTipoValLib(IdxValLib).CodSIIDTE = "" Then
               Call AddLogExp(lFNameLogExp, MsgIdDoc, "Impuesto: '" & gTipoValLib(IdxValLib).Nombre & "' Está descontinuado o no corresponde.")
               AddReg = False
            End If
            If vFld(RsImp("Tasa")) = 0 Then   'si esto está definido, no podemos continuar
               Call AddLogExp(lFNameLogExp, MsgIdDoc, "Impuesto: " & gTipoValLib(IdxValLib).Nombre & " No tiene definida la tasa del impuesto adicional")
               AddReg = False
            End If
            
            If vFld(RsImp("EsRecuperable")) <> 0 Then
               ImpAdicRec(r).IdTipoValLib = vFld(RsImp("IdTipoValLib"))
               ImpAdicRec(r).CodSIIDTE = gTipoValLib(IdxValLib).CodSIIDTE
               ImpAdicRec(r).Tasa = vFld(RsImp("Tasa"))
               ImpAdicRec(r).Valor = Abs(vFld(RsImp("SumDebe")) - vFld(RsImp("SumHaber")))
               
               EsImpCoDEspecial = ImpAdicRec(r).IdTipoValLib = LIBCOMPRAS_IVARETPARCLEGUMBRES
               EsImpCoDEspecial = EsImpCoDEspecial Or ImpAdicRec(r).IdTipoValLib = LIBCOMPRAS_IVARETPARCGANADO
               EsImpCoDEspecial = EsImpCoDEspecial Or ImpAdicRec(r).IdTipoValLib = LIBCOMPRAS_IVARETPARCMADERA
               EsImpCoDEspecial = EsImpCoDEspecial Or ImpAdicRec(r).IdTipoValLib = LIBCOMPRAS_IVARETPARCTRIGO
               EsImpCoDEspecial = EsImpCoDEspecial Or ImpAdicRec(r).IdTipoValLib = LIBCOMPRAS_IVARETPARCARROZ
               
               If ImpAdicRec(r).Tasa = 100 And EsImpCoDEspecial Then
                  ImpAdicRec(r).CodSIIDTE = ImpAdicRec(r).CodSIIDTE & "1"
               End If
               
               
               r = r + 1
            Else
               ImpNoRec = ImpNoRec + Abs(vFld(RsImp("SumDebe")) - vFld(RsImp("SumHaber")))
            End If
            
            If gTipoValLib(IdxValLib).Atributo = "IVARETPAR" Or gTipoValLib(IdxValLib).Atributo = "IVARETTOT" Then
               IVARet = IVARet + Abs(vFld(RsImp("SumDebe")) - vFld(RsImp("SumHaber")))
            End If
            
         End If
         
         RsImp.MoveNext
      Loop
      
      Call CloseRs(RsImp)
      
      HayImpAdic = False
      If r > 0 Then
         HayImpAdic = True
      End If
      

      
      'Monto IVA (Recuperable) corresponde a la suma de IVA Crédito Fiscal + IVA Activo Fijo, por lo tanto no se resta el IVA Activo Fijo
      IVA = vFld(Rs("IVA"))                                             ' - IVAActFijo
      If IVA > 0 And vFld(Rs("CodSIIDTEIVAIrrec")) <> 1 Then
         TxtLine = TxtLine & lSep & IVA
      Else
         TxtLine = TxtLine & lSep
      End If
      
      
      '******   Terminan campos fijos de la primera parte, hasta IVA Uso Común
      RegFijo1 = TxtLine
      Call AddDebug("ExpLibElectCompras: RegFijo1: " & RegFijo1)
      
      'Cod IVA no Rec y Monto IVA no Rec
      IVAUsoComun = 0
      If Val(vFld(Rs("CodSIIDTEIVAIrrec"))) > 0 And vFld(Rs("ValIVAIrrec")) > 0 Then
         TxtLine = TxtLine & lSep & vFld(Rs("CodSIIDTEIVAIrrec"))
         TxtLine = TxtLine & lSep & vFld(Rs("ValIVAIrrec"))
         IVAUsoComun = IVA - vFld(Rs("ValIVAIrrec"))
      Else
         TxtLine = TxtLine & lSep
         TxtLine = TxtLine & lSep
      End If

      RegFijo1ExtSinIVANoRec = RegFijo1 & lSep & lSep

      'IVA Uso Común
      If vFld(Rs("CodSIIDTEIVAIrrec")) = 1 Then
         TxtLine = TxtLine & lSep & IVAUsoComun
         RegFijo1ExtSinIVANoRec = RegFijo1ExtSinIVANoRec & lSep & IVAUsoComun
      Else
         TxtLine = TxtLine & lSep
         RegFijo1ExtSinIVANoRec = RegFijo1ExtSinIVANoRec & lSep
      End If
      
      RegFijo1Ext = TxtLine
      
      '******   Ahora vienen los impuestos adicionales que pueden ser más de uno

      
      'Llenamos la parte del registro variable en vacío, para cuando no hay impuestos adicionales, con y sin crédito (esta último va sumado)
      RegVarVacio = ""
      
      'Imp. Adicional Recuperable (con crédito)
      RegVarVacio = RegVarVacio & lSep
      RegVarVacio = RegVarVacio & lSep
      RegVarVacio = RegVarVacio & lSep
        
      
      'Armamos los campos fijos de la segunda parte (Entre medio están los impuestos adicionales que son variables y puede haber más de uno recuperable (con crédito) y la suma de no recuperable (sin crédito)
      
      'Monto Total
      RegFijo2 = lSep & vFld(Rs("Total")) & lSep
   
      'Imp. Adicional No Recuperable (sin crédito)
      If ImpNoRec > 0 Then
         RegFijo2 = RegFijo2 & ImpNoRec & lSep
      Else
         RegFijo2 = RegFijo2 & lSep
      End If
      
      'Monto Activo Fijo
      If IVAActFijo > 0 Then
         ActFijo = Round(IVAActFijo / gIVA, 0)
         RegFijo2 = RegFijo2 & ActFijo & lSep & IVAActFijo & lSep
      Else
         RegFijo2 = RegFijo2 & lSep
         RegFijo2 = RegFijo2 & lSep
      End If
      
      'IVA no Retenido
      
      If IVARet > 0 Then
         IVANoRet = Abs(IVA - IVARet)
         
         If IVANoRet <> 0 Then
            RegFijo2 = RegFijo2 & IVANoRet
         Else
            RegFijo2 = RegFijo2 & lSep
         End If
      Else
         RegFijo2 = RegFijo2 & lSep
      End If
      
      'Tabacos - Puros
      RegFijo2 = RegFijo2 & lSep
      'Tabacos - Cigarrillos
      RegFijo2 = RegFijo2 & lSep
      'Tabacos - Elaborados
      RegFijo2 = RegFijo2 & lSep
      'Imp. Vehiculos
      RegFijo2 = RegFijo2 & lSep
      'Cod. Sucursal SII
      RegFijo2 = RegFijo2 & lSep
      'Número interno
      RegFijo2 = RegFijo2 & lSep
      
      'Emisor Receptor: 1 en cualquiera de los siguientes casos
      ' es NCC o NDC asociada a auna Factura de Compra
      ' es NCC o NDC y contiene Impuestos adicionales o IVAs retenidos en su detalle
      
      IdxTipoDoc = GetTipoDoc(LIB_COMPRAS, vFld(Rs("TipoDoc")))
      EmisorReceptor = ""
      DimDoc = gTipoDoc(IdxTipoDoc).Diminutivo
      
      If DimDoc = "NCC" Or DimDoc = "NDC" Then     'es nota de crédito o débito
         
         If HayImpAdic Then
            EmisorReceptor = "1"
         
         ElseIf vFld(Rs("TipoDocAsoc")) > 0 Then
            DimTipoDocAsoc = gTipoDoc(GetTipoDoc(LIB_COMPRAS, vFld(Rs("TipoDocAsoc")))).Diminutivo
            If DimTipoDocAsoc = "FCC" Or HayImpAdic Then         'documento asociado es factura de compra
               EmisorReceptor = "1"
            End If
         End If
            
      End If
      
      RegFijo2 = RegFijo2 & lSep & EmisorReceptor

      r = 0

      'si hay impuestos adicionales, los metemos entre medio de RegFijo1 y RegFijo2, utilizando RegVar para armarlo
      If HayImpAdic = True Then
      
         Do While r <= UBound(ImpAdicRec)
         
            If ImpAdicRec(r).IdTipoValLib = 0 Then
               Exit Do
            End If
                        
            RegVar = ""
            
            'Imp. Adicional Recuperable
            If ImpAdicRec(r).IdTipoValLib > 0 And ImpAdicRec(r).Tasa > 0 And ImpAdicRec(r).Valor > 0 Then
         
               'Cod Otro Imp (Con Crédito), Tasa Otro Imp (Con Crédito), Monto Otro Imp (Con Crédito)
               RegVar = RegVar & lSep & ImpAdicRec(r).CodSIIDTE
               RegVar = RegVar & lSep & ReplaceStr(Format(ImpAdicRec(r).Tasa, "0.00"), ",", ".") '2 decimales obligatorio, separado por punto
               RegVar = RegVar & lSep & ImpAdicRec(r).Valor
         
            Else
               RegVar = RegVar & lSep
               RegVar = RegVar & lSep
               RegVar = RegVar & lSep
            
            End If
                                   
            'agregamos uno a uno los registros del documento con la parte variable de los impeustos adicionales
            If AddReg Then
               If r = 0 Then
                  TxtLine = RegFijo1Ext & RegVar & RegFijo2
               Else
                  TxtLine = RegFijo1ExtSinIVANoRec & RegVar & RegFijo2
               End If
                  
              Print #Fd, TxtLine
               Call AddDebug("ExpLibElectCompras: TxtLine: " & TxtLine)
            Else
               NErrores = NErrores + 1
            End If

            r = r + 1
         
         Loop
      
      Else     'no hay impuestos adicionales
      
         If AddReg Then
            TxtLine = RegFijo1Ext & RegVarVacio & RegFijo2
            Print #Fd, TxtLine
            Call AddDebug("ExpLibElectCompras: TxtLine: " & TxtLine)
         Else
            NErrores = NErrores + 1
        End If
         
      End If
      
      TxtLine = ""
      
      Rs.MoveNext
      
      nDocs = nDocs + 1
      
   Loop
      
   Call CloseRs(Rs)
   Close #Fd
   
   If NErrores > 0 Then
      MsgBox1 "Proceso de generación del Libro Electrónico de Compras finalizado con Errores." & vbCrLf & vbCrLf & "Se encontraron " & NErrores & " registros con errores.", vbExclamation + vbOKOnly
      MsgBox1 "Archivo generado con Errores:" & vbCrLf & vbCrLf & SFName, vbInformation + vbOKOnly
      If MsgBox1("¿Desea revisar el log de exportación " & lFNameLogExp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Me.hWnd, "open", lFNameLogExp, "", "", SW_SHOW)
      End If
   Else
      MsgBox1 "Proceso de generación del Libro Electrónico de Compras finalizado.", vbInformation + vbOKOnly
      MsgBox1 "Archivo generado:" & vbCrLf & vbCrLf & SFName, vbInformation + vbOKOnly
   End If
   
   

End Function
Private Function FilterChrSII(ByVal InpTxt As String) As String
   Dim ChrFilter As String
   Dim AuxStr As String
   Dim i As Integer
   
   AuxStr = InpTxt
   ChrFilter = "¡!”#$%&/()=’\¿?¨´*+{[^}]"
   
   For i = 1 To Len(ChrFilter)
   
      AuxStr = ReplaceStr(AuxStr, Mid(ChrFilter, i, 1), "_")
      
   Next i
   
   FilterChrSII = AuxStr

End Function

Private Function ExportLibComprasAcepta() As Boolean
   Dim FirstDay As Long, LastDay As Long
   Dim FPath As String
   Dim SFName As String
   Dim Fd As Long
   Dim MesAno As String, AnoMes As String
   Dim TipoOperacion As String
   Dim TxtFields As String
   Dim Q1 As String, Q2 As String
   Dim Rs As Recordset, RsImp As Recordset
   Dim TxtLine As String
   Dim nDocs As Integer
   Dim AddReg As Boolean
   Dim i As Integer, r As Integer
   Dim MsgTipoDoc As Boolean
   Dim TipoLib As Integer
   Dim TipoDocSII As String
   Dim RegVar As String
   Dim IdDoc As Long
   Dim ImpAdicRec(MAX_IMPADICDOC) As ImpAdicDoc_t
   Dim IdxValLib As Integer
   Dim IVAActFijo As Double, IVARet As Double, IVANoRet As Double, IVA As Double, ImpNoRec As Double
   Dim HayImpAdic As Boolean
   Dim DimTipoDocAsoc As String
   Dim EmisorReceptor As String
   Dim ActFijo As Double
   Dim lFNameLogExp As String
   Dim MsgIdDoc As String
   Dim NErrores As Integer
   Dim EsImpCoDEspecial As Boolean
   Dim LogPath As String
   Dim RegFijo1Ext As String
   Dim RegFijo1ExtSinIVANoRec As String
   Dim Diminutivo As String
   Dim CodDocSIIAsoc As String, NumDocAsoc As String
   Dim CodIVAIrrec As Integer
   
   TipoLib = LIB_COMPRAS
   Call FirstLastMonthDay(DateSerial(Val(Tx_Ano), CbItemData(Cb_Mes), 1), FirstDay, LastDay)
   MesAno = Right("0" & CbItemData(Cb_Mes), 2) & Tx_Ano
   AnoMes = Tx_Ano & Right("0" & CbItemData(Cb_Mes), 2)
   
   'vemos si hay documentos excluídos del libro
   Q1 = "SELECT Count(*) "
   Q1 = Q1 & " FROM Documento INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDOC = TipoDocs.TipoDoc "
   Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib & " AND Documento.Estado <> " & ED_ANULADO & " AND  ( TipoDocs.CodDocSII IS NULL OR TipoDocs.CodDocSII = '') AND (TipoDocs.CodDocDTESII IS NULL OR TipoDocs.CodDocDTESII = '' ) AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      If vFld(Rs(0)) > 0 Then
         MsgBox1 "Recuerde que los siguientes documentos no están considerados en el Informe Electrónico de Compras:" & vbCrLf & vbCrLf & "1. Form. Importaciones Exento (IEX)" & vbCrLf & vbCrLf & "2. Otros Documentos (OTC)" & vbCrLf & vbCrLf & "3. Crédito Impto. Timbres y Estam. (CIT)", vbInformation
      End If
   End If
   Call CloseRs(Rs)


   lSep = ";"
      
   lCamposLibComprasAcepta = "Tipo Documento" & lSep & "Folio" & lSep & "Folio Anulado [A]" & lSep & "Fecha Emision [AAAA-MM-DD]" & lSep & "EXCEPCION EMISOR/RECEPTOR" & lSep & "Codigo Sucursal Emisora" & lSep & "RUT Contraparte"
   lCamposLibComprasAcepta = lCamposLibComprasAcepta & lSep & "Razon Social Contraparte" & lSep & "Tipo de documento referencia" & lSep & "Folio documento referencia" & lSep & "Tasa Impuesto" & lSep & "Monto Exento" & lSep & "Monto Neto" & lSep & "Monto IVA" & lSep & "Monto Total" & lSep & "Monto Neto Activo Fijo"
   lCamposLibComprasAcepta = lCamposLibComprasAcepta & lSep & "IVA Activo Fijo" & lSep & "IVA Uso Comun" & lSep & "Impuestos sin derecho a cridito" & lSep & "IVA no retenido" & lSep & " IVA No Recuperable [1] Op. No grabada o exenta " & lSep & "IVA No Recuperable [2] Reg.Fuera Plazo" & lSep & "IVA No Recuperable [3] Gastos rechazados" & lSep & "IVA No Recuperable [4] Entregas Gratuitas"
   lCamposLibComprasAcepta = lCamposLibComprasAcepta & lSep & "IVA No Recuperable [9] Otros" & lSep & "TABACOS (Cigarros puros)" & lSep & "TABACOS (Cigarrillos)" & lSep & "TABACOS (Tabaco elaborado)" & lSep & "Impuesto  Vehmculos Automsviles" & lSep & "Csdigo del Impuesto o Recargo 1" & lSep & "Tasa del Impuesto o Recargo 1" & lSep & "Valor del Impuesto o Recargo 1"
   lCamposLibComprasAcepta = lCamposLibComprasAcepta & lSep & "Csdigo del Impuesto o Recargo 2" & lSep & "Tasa del Impuesto o Recargo 2" & lSep & "Valor del Impuesto o Recargo 2" & lSep & "Csdigo del Impuesto o Recargo 3" & lSep & "Tasa del Impuesto o Recargo 3" & lSep & "Valor del Impuesto o Recargo 3" & lSep & "Csdigo del Impuesto o Recargo 4"
   lCamposLibComprasAcepta = lCamposLibComprasAcepta & lSep & "Tasa del Impuesto o Recargo 4" & lSep & "Valor del Impuesto o Recargo 4" & lSep & "Csdigo del Impuesto o Recargo 5" & lSep & "Tasa del Impuesto o Recargo 5" & lSep & "Valor del Impuesto o Recargo 5"
   lCamposLibComprasAcepta = lCamposLibComprasAcepta & lSep & "Csdigo del Impuesto o Recargo 6" & lSep & "Tasa del Impuesto o Recargo 6" & lSep & "Valor del Impuesto o Recargo 6" & lSep & "Csdigo del Impuesto o Recargo 7" & lSep & "Tasa del Impuesto o Recargo 7" & lSep & "Valor del Impuesto o Recargo 7"
   lCamposLibComprasAcepta = lCamposLibComprasAcepta & lSep & "Csdigo del Impuesto o Recargo 8" & lSep & "Tasa del Impuesto o Recargo 8" & lSep & "Valor del Impuesto o Recargo 8" & lSep & "Csdigo del Impuesto o Recargo 9" & lSep & "Tasa del Impuesto o Recargo 9" & lSep & "Valor del Impuesto o Recargo 9"
   lCamposLibComprasAcepta = lCamposLibComprasAcepta & lSep & "Csigo del Impuesto o Recargo 10" & lSep & "Tasa del Impuesto o Recargo 10" & lSep & "Valor del Impuesto o Recargo 10"
      
   TipoOperacion = "Com-"
   
   On Error Resume Next
   
   FPath = gExportPath & "\Factura\"
   MkDir FPath
   FPath = gExportPath & "\Factura\" & gEmpresa.Rut
   MkDir FPath
   
   LogPath = FPath & "\Log"
   MkDir LogPath
   
   On Error GoTo 0
      
   SFName = FPath & "\LACP_" & TipoOperacion & AnoMes & ".csv"
   lFNameLogExp = LogPath & "\LACP_Com-" & Format(Now, "yyyymmdd") & ".log"
         
   Fd = FreeFile()
   Open SFName For Output As #Fd
   If Err Then
      MsgErr "Error al abrir el archivo '" & SFName & "'.", vbExclamation
      Exit Function
   End If
      
   'ponemos códigos de campos en primera línea
   TxtFields = lCamposLibComprasAcepta
   Print #Fd, TxtFields
   

   'Consulta para obtener cada documento y su detalle
   Q1 = "SELECT IdDoc, Documento.TipoDoc, DTE, TipoDocs.CodDocSII, TipoDocs.CodDocDTESII, Documento.Giro, NumDoc, CorrInterno, NumDocHasta, Documento.MovEdited, "
   Q1 = Q1 & " Documento.IdEntidad, Documento.RutEntidad, Documento.NombreEntidad, Entidades.Rut, Entidades.NotValidRut, Entidades.Nombre,"
   Q1 = Q1 & " FEmision, FEmisionOri, FVenc, Exento, Afecto, IVA, IVAIrrecuperable, ValIVAIrrec, CodSIIDTEIVAIrrec, "
   Q1 = Q1 & " OtroImp, OtrosVal, Total, Descrip, Documento.Estado, Documento.TipoDocAsoc, Documento.IdDocAsoc, Documento.IVAActFijo "
   Q1 = Q1 & " FROM ( Documento INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDOC = TipoDocs.TipoDoc) "
   Q1 = Q1 & " LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad   AND Entidades.IdEmpresa = Documento.IdEmpresa "
   Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib & " AND Documento.Estado <> " & ED_ANULADO & " AND ( TipoDocs.CodDocSII <> '' OR TipoDocs.CodDocDTESII <> '' ) AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " AND Documento.IdEmpresa = " & gEmpresa.id & " AND Documento.Ano = " & gEmpresa.Ano
   Q1 = Q1 & " ORDER BY IdDoc"
      
   
   Set Rs = OpenRs(DbMain, Q1)
   
   TxtLine = ""
   nDocs = 0
   NErrores = 0
   
   Do While Not Rs.EOF
   
      AddReg = True
      IdDoc = vFld(Rs("IdDoc"))
      Diminutivo = gTipoDoc(GetTipoDoc(LIB_COMPRAS, vFld(Rs("TipoDoc")))).Diminutivo
            
      If vFld(Rs("DTE")) Then
         TipoDocSII = vFld(Rs("CodDocDTESII"))
      Else
         TipoDocSII = vFld(Rs("CodDocSII"))
      End If
      
      'Tipo Doc,Folio
      TxtLine = TipoDocSII & lSep & vFld(Rs("NumDoc"))
      
      'Anulado[A]
      If vFld(Rs("Estado")) = ED_ANULADO Then
         TxtLine = TxtLine & lSep & "A"
      Else
         TxtLine = TxtLine & lSep
      End If
      
      'Fecha Emisión
      TxtLine = TxtLine & lSep & Format(vFld(Rs("FEmisionOri")), "yyyy-mm-dd")
      
      
      
      
      '********  Obtenermos los impruestos adicionales e IVA activo fijo
      For r = 0 To UBound(ImpAdicRec)
         ImpAdicRec(r).IdTipoValLib = 0
         ImpAdicRec(r).Valor = 0
         ImpAdicRec(r).Tasa = 0
      Next r
      
      Q2 = "SELECT MovDocumento.IdTipoValLib, MovDocumento.Tasa, MovDocumento.EsRecuperable, Sum(Debe) as SumDebe, Sum(Haber) as SumHaber "
      Q2 = Q2 & " FROM MovDocumento INNER JOIN TipoValor ON MovDocumento.IdTipoValLib = TipoValor.Codigo "
      Q2 = Q2 & " WHERE IdDoc = " & IdDoc & " AND TipoValor.TipoLib = " & LIB_COMPRAS & " AND IdTipoValLib >= " & LIBCOMPRAS_OTROSIMP & " AND Atributo NOT IN ('IVAIRREC', 'SINUSO')"
      Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Q2 = Q2 & " GROUP BY MovDocumento.IdTipoValLib, MovDocumento.Tasa, MovDocumento.EsRecuperable"
      
      Set RsImp = OpenRs(DbMain, Q2)
            
      r = 0
      IVAActFijo = 0
      IVARet = 0
      ImpNoRec = 0
      
      Do While Not RsImp.EOF
      
         IdxValLib = GetTipoValLib(LIB_COMPRAS, vFld(RsImp("IdTipoValLib")))
         
         If vFld(RsImp("IdTipoValLib")) = LIBCOMPRAS_IVAACTFIJO Then
            IVAActFijo = Abs(vFld(RsImp("SumDebe")) - vFld(RsImp("SumHaber")))

         Else
         
            If gTipoValLib(IdxValLib).CodSIIDTE = "" Then
               Call AddLogExp(lFNameLogExp, MsgIdDoc, "Impuesto: '" & gTipoValLib(IdxValLib).Nombre & "' Está descontinuado o no corresponde.")
               AddReg = False
            End If
            If vFld(RsImp("Tasa")) = 0 Then   'si esto está definido, no podemos continuar
               Call AddLogExp(lFNameLogExp, MsgIdDoc, "Impuesto: " & gTipoValLib(IdxValLib).Nombre & " No tiene definida la tasa del impuesto adicional")
               AddReg = False
            End If
            
            If vFld(RsImp("EsRecuperable")) <> 0 Then
               ImpAdicRec(r).IdTipoValLib = vFld(RsImp("IdTipoValLib"))
               ImpAdicRec(r).CodSIIDTE = gTipoValLib(IdxValLib).CodSIIDTE
               ImpAdicRec(r).Tasa = vFld(RsImp("Tasa"))
               ImpAdicRec(r).Valor = Abs(vFld(RsImp("SumDebe")) - vFld(RsImp("SumHaber")))
               
               EsImpCoDEspecial = ImpAdicRec(r).IdTipoValLib = LIBCOMPRAS_IVARETPARCLEGUMBRES
               EsImpCoDEspecial = EsImpCoDEspecial Or ImpAdicRec(r).IdTipoValLib = LIBCOMPRAS_IVARETPARCGANADO
               EsImpCoDEspecial = EsImpCoDEspecial Or ImpAdicRec(r).IdTipoValLib = LIBCOMPRAS_IVARETPARCMADERA
               EsImpCoDEspecial = EsImpCoDEspecial Or ImpAdicRec(r).IdTipoValLib = LIBCOMPRAS_IVARETPARCTRIGO
               EsImpCoDEspecial = EsImpCoDEspecial Or ImpAdicRec(r).IdTipoValLib = LIBCOMPRAS_IVARETPARCARROZ
               
               If ImpAdicRec(r).Tasa = 100 And EsImpCoDEspecial Then
                  ImpAdicRec(r).CodSIIDTE = ImpAdicRec(r).CodSIIDTE & "1"
               End If
               
               
               r = r + 1
            Else
               ImpNoRec = ImpNoRec + Abs(vFld(RsImp("SumDebe")) - vFld(RsImp("SumHaber")))
            End If
            
            If gTipoValLib(IdxValLib).Atributo = "IVARETPAR" Or gTipoValLib(IdxValLib).Atributo = "IVARETTOT" Then
               IVARet = IVARet + Abs(vFld(RsImp("SumDebe")) - vFld(RsImp("SumHaber")))
            End If
            
         End If
         
         RsImp.MoveNext
      Loop
      
      Call CloseRs(RsImp)
      
      HayImpAdic = False
      If r > 0 Then
         HayImpAdic = True
      End If
      


      'Emisor Receptor: 1 en cualquiera de los siguientes casos
      ' es NCC o NDC asociada a auna Factura de Compra
      ' es NCC o NDC y contiene Impuestos adicionales o IVAs retenidos en su detalle
      EmisorReceptor = ""
      CodDocSIIAsoc = ""
      NumDocAsoc = ""
      If Diminutivo = "NCC" Or Diminutivo = "NDC" Then     'es nota de crédito o débito
         
         CodDocSIIAsoc = GetCodSIIDoc(vFld(Rs("IdDocAsoc")), NumDocAsoc)
         If CodDocSIIAsoc <> "" Then
            If HayImpAdic Then
               EmisorReceptor = "1"
            
            ElseIf vFld(Rs("TipoDocAsoc")) > 0 Then
               DimTipoDocAsoc = gTipoDoc(GetTipoDoc(LIB_COMPRAS, vFld(Rs("TipoDocAsoc")))).Diminutivo
               If DimTipoDocAsoc = "FCC" Or HayImpAdic Then         'documento asociado es factura de compra
                  EmisorReceptor = "1"
               End If
            End If
         End If
         
      End If
      TxtLine = TxtLine & lSep & EmisorReceptor
      
      'Codigo Sucursal Emisora
      TxtLine = TxtLine & lSep
      
      MsgIdDoc = "DOC.: " & Diminutivo & "-" & vFld(Rs("NumDoc")) & " - RUT: " & FmtRut(gEmpresa.Rut)
      
      'Rut Contraparte
      TxtLine = TxtLine & lSep & vFld(Rs("Rut")) & "-" & DV_Rut(vFld(Rs("Rut")))
      
      'Razón Social Contraparte
      TxtLine = TxtLine & lSep & FilterChrSII(vFld(Rs("Nombre")))
      
      'Tipo de documento referencia
      TxtLine = TxtLine & lSep & CodDocSIIAsoc
      
      'Folio documento referencia
      TxtLine = TxtLine & lSep & NumDocAsoc
      
      'Tasa Impuesto IVA. En esta columna, siempre se registra 19.
      TxtLine = TxtLine & lSep & gIVA * 100
      
      'Monto Exento
      If vFld(Rs("Exento")) > 0 Then
         TxtLine = TxtLine & lSep & vFld(Rs("Exento"))
      Else
         TxtLine = TxtLine & lSep
      End If
      
      'Monto Neto
      If vFld(Rs("Afecto")) > 0 Then
         TxtLine = TxtLine & lSep & vFld(Rs("Afecto"))
      Else
         TxtLine = TxtLine & lSep
      End If
      
      'Monto IVA (Recuperable) corresponde a la suma de IVA Crédito Fiscal + IVA Activo Fijo, por lo tanto no se resta el IVA Activo Fijo
      IVA = vFld(Rs("IVA"))                                             ' - IVAActFijo
      If IVA > 0 And vFld(Rs("CodSIIDTEIVAIrrec")) <> 1 Then
         TxtLine = TxtLine & lSep & IVA
      Else
         TxtLine = TxtLine & lSep
      End If
      
      'Monto Total
      TxtLine = TxtLine & lSep & vFld(Rs("Total"))
      
      'Monto Activo Fijo, IVA Activo Fijo
      If IVAActFijo > 0 Then
         ActFijo = Round(IVAActFijo / gIVA, 0)
         If ActFijo > vFld(Rs("Afecto")) Then
            ActFijo = vFld(Rs("Afecto"))
         End If
         TxtLine = TxtLine & lSep & ActFijo & lSep & IVAActFijo & lSep
      Else
         TxtLine = TxtLine & lSep & lSep & lSep
      End If
      
      'IVA Uso Común
      If vFld(Rs("CodSIIDTEIVAIrrec")) = 1 Then
         TxtLine = TxtLine & lSep & IVA & lSep
      Else
         TxtLine = TxtLine & lSep
      End If
      
      'Imp. Adicional No Recuperable (sin crédito)
      If ImpNoRec > 0 Then
         TxtLine = TxtLine & ImpNoRec & lSep
      Else
         TxtLine = TxtLine & lSep
      End If
     
      'IVA no Retenido
      
      If IVARet > 0 Then
         IVANoRet = Abs(IVA - IVARet)
         
         If IVANoRet <> 0 Then
            TxtLine = TxtLine & IVANoRet
         Else
            TxtLine = TxtLine & lSep
         End If
      Else
         TxtLine = TxtLine & lSep
      End If
      
            
      'Cod IVA no Rec y Monto IVA no Rec
      CodIVAIrrec = Val(vFld(Rs("CodSIIDTEIVAIrrec")))
      
      If CodIVAIrrec > 0 Then
         If CodIVAIrrec = 9 Then
            CodIVAIrrec = 5
         End If
         
         For r = 1 To CodIVAIrrec    'IVAIrrec 1, 2, 3, 4, 9
            TxtLine = TxtLine & lSep
         Next r
         TxtLine = TxtLine & vFld(Rs("ValIVAIrrec"))
         
         For r = CodIVAIrrec + 1 To 5   'IVAIrrec 1, 2, 3, 4, 9
            TxtLine = TxtLine & lSep
         Next r
      Else
         For r = 1 To 5  'IVAIrrec 1, 2, 3, 4, 9
            TxtLine = TxtLine & lSep
         Next r
      End If
            
      'Tabacos - Puros
      TxtLine = TxtLine & lSep
      'Tabacos - Cigarrillos
      TxtLine = TxtLine & lSep
      'Tabacos - Elaborados
      TxtLine = TxtLine & lSep
      'Imp. Vehiculos
      TxtLine = TxtLine & lSep

      
      '******   Ahora vienen los impuestos adicionales que pueden ser más de uno

      

      r = 0
      RegVar = ""

      'si hay impuestos adicionales
      If HayImpAdic = True Then
      
         Do While r <= UBound(ImpAdicRec)
         
            If ImpAdicRec(r).IdTipoValLib = 0 Then
               Exit Do
            End If
                                    
            'Imp. Adicional Recuperable
            If ImpAdicRec(r).IdTipoValLib > 0 And ImpAdicRec(r).Tasa > 0 And ImpAdicRec(r).Valor > 0 Then
         
               'Cod Otro Imp (Con Crédito), Tasa Otro Imp (Con Crédito), Monto Otro Imp (Con Crédito)
               RegVar = RegVar & ImpAdicRec(r).CodSIIDTE & lSep
               RegVar = RegVar & Str(Format(ImpAdicRec(r).Tasa, "#.#")) & lSep   '1 decimal a lo más y sepeardo por punto
               RegVar = RegVar & ImpAdicRec(r).Valor & lSep
         
            Else
               RegVar = RegVar & lSep
               RegVar = RegVar & lSep
               RegVar = RegVar & lSep
            
            End If
                                   
            r = r + 1
         
            If r > 10 Then
               Exit Do
            End If
            
         Loop
         
      
      End If
      
      'agregamos columnas de impuestos adicionales en blanco hasta llegar a 10
      r = r + 1
      Do While r <= 10
      
         RegVar = RegVar & lSep
         RegVar = RegVar & lSep
         RegVar = RegVar & lSep
         r = r + 1
         
      Loop
     
      If AddReg Then
         TxtLine = TxtLine & Left(RegVar, Len(RegVar) - 1)
         Print #Fd, TxtLine
         Call AddDebug("ExpLibElectCompras: TxtLine: " & TxtLine)
      Else
         NErrores = NErrores + 1
      End If
         
      
      TxtLine = ""
      
      Rs.MoveNext
      
      nDocs = nDocs + 1
      
   Loop
      
   Call CloseRs(Rs)
   Close #Fd
   
   If NErrores > 0 Then
      MsgBox1 "Proceso de generación del Libro Electrónico de Compras finalizado con Errores." & vbCrLf & vbCrLf & "Se encontraron " & NErrores & " registros con errores.", vbExclamation + vbOKOnly
      MsgBox1 "Archivo generado con Errores:" & vbCrLf & vbCrLf & SFName, vbInformation + vbOKOnly
      If MsgBox1("¿Desea revisar el log de exportación " & lFNameLogExp & vbCrLf & vbCrLf & "para ver el detalle de los errores?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         Call ShellExecute(Me.hWnd, "open", lFNameLogExp, "", "", SW_SHOW)
      End If
   Else
      MsgBox1 "Proceso de exportación del Libro de Compras para Facturación Electrónica finalizado.", vbInformation + vbOKOnly
      MsgBox1 "Archivo generado:" & vbCrLf & vbCrLf & SFName, vbInformation + vbOKOnly
   End If
   
   

End Function

Private Function ExportLibSII_Old(ByVal TipoLib As Integer) As Boolean
   Dim RegLib(MAX_CODFLDSII_COMPRAS) As String
   Dim FirstDay As Long, LastDay As Long
   Dim FPath As String
   Dim SFName As String
   Dim Fd As Long
   Dim MesAno As String, AnoMes As String
   Dim TipoOperacion As String
   Dim FldSII() As String
   Dim TxtFields As String
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TxtLine As String
   Dim nDocs As Integer
   Dim AddReg As Boolean
   Dim i As Integer
   Dim MsgTipoDoc As Boolean
   

   Call FirstLastMonthDay(DateSerial(Val(Tx_Ano), CbItemData(Cb_Mes), 1), FirstDay, LastDay)
   MesAno = Right("0" & CbItemData(Cb_Mes), 2) & Tx_Ano
   AnoMes = Tx_Ano & Right("0" & CbItemData(Cb_Mes), 2)
   
   If TipoLib = LIB_VENTAS Then
      FldSII = gFldSIIVentas
      TipoOperacion = "V"
   Else
      FldSII = gFldSIICompras
      TipoOperacion = "C"
   End If

      
   On Error Resume Next
   
   FPath = gExportPath & "\SII\"
   MkDir FPath
   FPath = gExportPath & "\SII\" & gEmpresa.Rut
   MkDir FPath
   
   On Error GoTo 0
      
   SFName = FPath & "\" & TipoOperacion & AnoMes & ".txt"
         
   Fd = FreeFile()
   Open SFName For Output As #Fd
   If Err Then
      MsgErr "Error al abrir el archivo '" & SFName & "'.", vbExclamation
      Exit Function
   End If
      
   'ponemos códigos de campos en primera línea
   TxtFields = GenRegLargoFijo(FldSII, FldSII)
   Print #Fd, TxtFields

   Q1 = "SELECT IdDoc, Documento.TipoDoc, DTE, TipoDocs.CodDocSII, TipoDocs.CodDocDTESii, Documento.Giro, NumDoc, CorrInterno, NumDocHasta, Documento.MovEdited, "
   Q1 = Q1 & " Documento.IdEntidad, Documento.RutEntidad, Documento.NombreEntidad, Entidades.Rut, Entidades.NotValidRut, Entidades.Nombre,"
   Q1 = Q1 & " FEmision, FEmisionOri, FVenc, Exento, Afecto, IVA,  "
   Q1 = Q1 & " OtroImp, OtrosVal, Total, Descrip  "
   Q1 = Q1 & " FROM ( Documento LEFT JOIN Entidades ON Documento.IdEntidad = Entidades.IdEntidad )"
   Q1 = Q1 & " INNER JOIN TipoDocs ON Documento.TipoLib = TipoDocs.TipoLib AND Documento.TipoDOC = TipoDocs.TipoDoc"
   Q1 = Q1 & " WHERE Documento.TipoLib = " & TipoLib & " AND (FEmision BETWEEN " & FirstDay & " AND " & LastDay & ")"
   Q1 = Q1 & " ORDER BY FEmision, NumDoc"
      
   
   Set Rs = OpenRs(DbMain, Q1)
   
   TxtLine = ""
   nDocs = 0
     
   Do While Not Rs.EOF
   
      AddReg = True
      
      'RUT, Fecha Registro, Numero Interno, Tipo de Operación
      RegLib(1) = gEmpresa.Rut & DV_Rut(gEmpresa.Rut)      'con dígito y sin guión
      RegLib(2) = MesAno
      RegLib(3) = vFld(Rs("CorrInterno"))
      RegLib(4) = TipoOperacion
      
      'Tipo de Documento
      If vFld(Rs("DTE")) <> 0 Then
         RegLib(5) = vFld(Rs("CodDocDTESII"))
      Else
         RegLib(5) = vFld(Rs("CodDocSII"))
      End If
      
      If RegLib(5) = "" Then    'no existe equivalente en el SII de este tipo de documento
         AddReg = False
         If Not MsgTipoDoc Then
            MsgBox1 "Atención:" & vbCrLf & vbCrLf & "Este libro incluye algunos documentos cuyo tipo no está definido en el formato del archivo del SII.", vbInformation + vbOKOnly
            MsgTipoDoc = True
         End If
      End If
      
      'Numero documento, Fecha Documento
      RegLib(6) = vFld(Rs("NumDoc"))
      RegLib(7) = Format(vFld(Rs("FEmisionOri")), "ddmmyyyy")
      
      'RUT asociado, nombre o razón social
      If vFld(Rs("IdEntidad")) = 0 Then
         If vFld(Rs("RutEntidad")) <> "" And vFld(Rs("RutEntidad")) <> "0" Then
            RegLib(8) = vFld(Rs("RutEntidad")) & DV_Rut(vFld(Rs("RutEntidad")))     'con dígito y sin guión
            RegLib(9) = vFld(Rs("NombreEntidad"))
         End If
      Else
         If vFld(Rs("Rut")) <> "" And vFld(Rs("Rut")) <> "0" Then
            RegLib(8) = vFld(Rs("Rut")) & DV_Rut(vFld(Rs("Rut")))     'con dígito y sin guión
            RegLib(9) = vFld(Rs("Nombre"))
         End If
      End If

      'Valores documento
      RegLib(10) = vFld(Rs("Exento"))
      RegLib(11) = vFld(Rs("Afecto"))
      RegLib(12) = vFld(Rs("IVA"))
      RegLib(13) = vFld(Rs("Total"))
      
      If AddReg Then
         Call GetOtrosImp(RegLib, TipoLib, vFld(Rs("IdDoc")), FldSII)
      End If
         
      If AddReg Then
         TxtLine = GenRegLargoFijo(FldSII, RegLib)
         Print #Fd, TxtLine
      End If
      
      TxtLine = ""
      
      For i = 1 To UBound(RegLib)
         RegLib(i) = ""
      Next i

      Rs.MoveNext
      
      nDocs = nDocs + 1
      
   Loop
      
   Call CloseRs(Rs)
   Close #Fd
   
   MsgBox1 "Proceso de exportación del libro finalizado.", vbInformation + vbOKOnly
   MsgBox1 "Archivo generado:" & vbCrLf & vbCrLf & SFName, vbInformation + vbOKOnly


End Function


Private Function GetOtrosImp(RegLib() As String, ByVal TipoLib As Integer, ByVal IdDoc As Long, FldSII) As Boolean
   Dim Q1 As String
   Dim Rs As Recordset
   Dim TotOtrosImp As Double
   Dim Col As Integer
   Dim Valor As Double
   Dim i As Integer
   
   GetOtrosImp = False
      
   If IdDoc <= 0 Then
      Exit Function
   End If
         
'   If IdDoc = 1712 Then
'      MsgBeep vbExclamation
'   End If
   
   Q1 = "SELECT CodImpSII, Sum(Debe) as SumDebe, Sum(Haber) as SumHaber "
   Q1 = Q1 & " FROM MovDocumento INNER JOIN TipoValor ON MovDocumento.IdTipoValLib = TipoValor.Codigo "
   Q1 = Q1 & " WHERE TipoLib = " & TipoLib & " AND IdDoc = " & IdDoc & " AND NOT " & SQLNull("CodImpSII")
   Q1 = Q1 & " AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
   Q1 = Q1 & " GROUP BY CodImpSII"
   
   Set Rs = OpenRs(DbMain, Q1)
      
   Do While Not Rs.EOF
   
      Valor = Abs(vFld(Rs("SumDebe")) - vFld(Rs("SumHaber")))
      
      For i = 14 To UBound(FldSII)
      
         If Trim(FldSII(i)) = vFld(Rs("CodImpSII")) Then
            RegLib(i) = Valor
         End If
      Next i
      
      
      Rs.MoveNext
   
   Loop
   
   Call CloseRs(Rs)
   
   GetOtrosImp = True
   
End Function


