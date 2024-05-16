VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmTrapasoODToODF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso de OD a ODF"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9735
      Left            =   -960
      TabIndex        =   0
      Top             =   -720
      Width           =   13335
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   9960
         TabIndex        =   7
         Top             =   1680
         Width           =   1635
      End
      Begin VB.CommandButton Bt_Traspaso 
         Caption         =   "Traspasar"
         Height          =   315
         Left            =   9960
         TabIndex        =   6
         Top             =   1200
         Width           =   1635
      End
      Begin VB.Frame Frame2 
         Caption         =   "Estado de Traspaso"
         Height          =   1215
         Left            =   1200
         TabIndex        =   1
         Top             =   1080
         Width           =   8655
         Begin VB.CommandButton Bt_SelecAll 
            Caption         =   "&Seleccionar Todo"
            Height          =   675
            Left            =   6960
            Picture         =   "FrmTrapasoODToODF.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox Cb_Tratamiento 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   480
            Width           =   1545
         End
         Begin VB.Label Lbl_Tratamiento 
            Caption         =   "Tratamiento"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   520
            Width           =   975
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   4875
         Left            =   1200
         TabIndex        =   4
         Top             =   2520
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8599
         _Version        =   393216
         Rows            =   3
         Cols            =   15
         SelectionMode   =   1
      End
   End
End
Attribute VB_Name = "FrmTrapasoODToODF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_IDDOC = 0
Const C_TIPOLIB = 1
Const C_IDTIPOLIB = 2
Const C_TIPODOC = 3
Const C_NUMDOC = 4
Const C_VALOR = 5
Const C_DESC = 6
Const C_CHECK = 7
'2855046 ffv
Const C_IDCTABANCO = 8
'2855046 ffv

Const NCOLS = C_IDCTABANCO

Dim lOper As Integer
Dim lTogCheck As Boolean




Private Sub Bt_Close_Click()
Unload Me
End Sub

Private Sub Bt_SelecAll_Click()
Call LoadGrid(True)
End Sub

Private Sub Bt_Traspaso_Click()

    If MsgBox1("¿Está seguro que desea Traspasar los Otros Documentos a Otros Documentos Full? (No se puede Volver Atras)", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes Then
      MousePointer = vbDefault
      Call SaveAll
    End If


End Sub

Private Sub Form_Load()
Dim i As Integer

For i = 1 To UBound(gTratamiento)
      Cb_Tratamiento.AddItem ReplaceStr(gTratamiento(i).Nombre, "Libro de ", "")
      Cb_Tratamiento.ItemData(Cb_Tratamiento.NewIndex) = gTratamiento(i).id 'i
   Next i
   Cb_Tratamiento.ListIndex = 1
   
   Call SetUpGrid
   Call LoadGrid

End Sub

Private Sub SaveAll()
Dim Q1 As String
Dim i As Long
Dim Cant As Long

'2855046 ffv
Dim ctaODf As Long
Dim Q2 As String
Dim Rs As Recordset

ctaODf = 0
'2855046 ffv

Cant = 0

    For i = Grid.FixedRows To Grid.rows - 1
        If Grid.TextMatrix(i, C_CHECK) = "X" Then
            
            '2855046 ffv
            If Grid.TextMatrix(i, C_IDCTABANCO) = 0 Then
            
            Q2 = ""
            Q2 = " SELECT Valor FROM ParamEmpresa "
            
                If ItemData(Cb_Tratamiento) = 1 Then
                 Q2 = Q2 & " WHERE Tipo = '" & "CTAODFACTI" & "'"
                Else
                 Q2 = Q2 & " WHERE Tipo = '" & "CTAODFPASI" & "'"
                End If
             
             Q2 = Q2 & " AND IdEmpresa = " & gEmpresa.id
             Q2 = Q2 & " AND Ano =" & gEmpresa.Ano
            
             Set Rs = OpenRs(DbMain, Q2)
   
                If Rs.EOF = False Then
                   
                ctaODf = vFld(Rs("Valor"))
                Else
                ctaODf = 0
                
                End If
            
            Else
            
            ctaODf = Grid.TextMatrix(i, C_IDCTABANCO)
            End If
            '2855046 ffv
            
            Q1 = " UPDATE DOCUMENTO "
            Q1 = Q1 & " SET TipoLib = 8, "
            Q1 = Q1 & "     TipoDoc = 1, "
            Q1 = Q1 & "     Tratamiento = " & ItemData(Cb_Tratamiento)
            '2855046 ffv
            Q1 = Q1 & "     ,IdCtaBanco = " & ctaODf
            '2855046 ffv
            Q1 = Q1 & " WHERE TIPOLIB = 5 "
            Q1 = Q1 & " AND IDDOC = " & Grid.TextMatrix(i, C_IDDOC)
            Call ExecSQL(DbMain, Q1)
            
            'Tracking 3227543
            Call SeguimientoDocumento(Grid.TextMatrix(i, C_IDDOC), gEmpresa.id, gEmpresa.Ano, "FrmTraspasoODToODF.SaveAll", Q1, 1, "", gUsuario.IdUsuario, 1, 2)
            ' fin 3227543
            
            Cant = Cant + 1
        End If
    Next i

'    Q1 = " INSERT INTO DOCUMENTOFULL "
'    Q1 = Q1 & " ([IdEmpresa],[Ano],[IdCompCent],[IdCompPago],[TipoLib],[TipoDoc],[NumDoc],[NumDocHasta],[IdEntidad],[TipoEntidad],[RutEntidad],[NombreEntidad],[FEmision],[FVenc],[Descrip],[Estado],[Exento],[IdCuentaExento],[Afecto],[IdCuentaAfecto],[IVA],[IdCuentaIVA],[OtroImp],[IdCuentaOtroImp],[Total],[IdCuentaTotal],[IdUsuario],[FechaCreacion] "
'    Q1 = Q1 & " ,[FEmisionOri],[CorrInterno],[SaldoDoc],[FExported],[OldIdDoc],[DTE],[PorcentRetencion],[TipoRetencion],[MovEdited],[OtrosVal],[FImporF29],[NumDocRef],[IdCtaBanco],[TipoRelEnt],[IdSucursal],[TotPagadoAnoAnt],[FImportSuc],[Giro],[FacCompraRetParcial],[IVAIrrecuperable],[DocOtrosEnAnalitico],[OldIdDocTmp],[NumFiscImpr],[NumInformeZ] "
'    Q1 = Q1 & " ,[CantBoletas],[VentasAcumInfZ],[IdDocAsoc],[PropIVA],[ValIVAIrrec],[IVAInmueble],[FImpFacturacion],[CodSIIDTEIVAIrrec],[TipoDocAsoc],[IVAActFijo],[EntRelacionada],[NumCuotas],[CompraBienRaiz],[NumDocAsoc],[DTEDocAsoc],[IdANegCCosto],[UrlDTE],[CodCtaAfectoOld],[CodCtaExentoOld],[CodCtaTotalOld],[DocOtroEsCargo],[ValRet3Porc],[IdCuentaRet3Porc],[Tratamiento]) "
'    Q1 = Q1 & " SELECT DOC.IdEmpresa, DOC.Ano,IdCompCent, IdCompPago, 8 as TipoLib, 1 as TipoDoc, NumDoc, NumDocHasta, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, DOC.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, "
'    Q1 = Q1 & " IdCuentaOtroImp , Total, IdCuentaTotal, DOC.IdUsuario, DOC.FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal, TotPagadoAnoAnt, "
'    Q1 = Q1 & " FImportSuc, Giro, FacCompraRetParcial, IVAIrrecuperable, DocOtrosEnAnalitico, OldIdDocTmp, NumFiscImpr, NumInformeZ, CantBoletas, VentasAcumInfZ, IdDocAsoc, PropIVA, ValIVAIrrec, IVAInmueble, FImpFacturacion, CodSIIDTEIvaIrrec, TipoDocAsoc, IVAActFijo, "
'    Q1 = Q1 & " EntRelacionada, NumCuotas, CompraBienRaiz, NumDocAsoc, DTEDocAsoc, IdANegCCosto, UrlDTE, CodCtaAfectoOld, CodCtaExentoOld, CodCtaTotalOld, DocOtroEsCargo, ValRet3Porc, IdCuentaRet3Porc, " & ItemData(Cb_Tratamiento) & " as Tratamiento "
'    Q1 = Q1 & " FROM DOCUMENTO DOC "
'    Q1 = Q1 & " Where TIPOLIB = 5 "
'    Call ExecSQL(DbMain, Q1)

'    Q1 = " INSERT INTO COMPROBANTEFULL "
'    Q1 = Q1 & " ([IdEmpresa],[Ano],[Correlativo],[Fecha],[Tipo],[Estado],[Glosa],[TotalDebe],[TotalHaber],[IdUsuario],[FechaCreacion],[ImpResumido],[EsCCMM],[FechaImport],[TipoAjuste],[OtrosIngEg14TER]) "
'    Q1 = Q1 & " SELECT COM.IdEmpresa, COM.Ano, COM.Correlativo, COM.Fecha, COM.Tipo, COM.Estado, COM.Glosa, COM.TotalDebe, COM.TotalHaber, COM.IdUsuario, COM.FechaCreacion, COM.ImpResumido, COM.EsCCMM, COM.FechaImport, COM.TipoAjuste, COM.OtrosIngEg14TER "
'    Q1 = Q1 & " FROM COMPROBANTE COM "
'    Q1 = Q1 & " WHERE IDCOMP IN  (SELECT IDCOMP FROM MOVCOMPROBANTE WHERE IDDOC IN (SELECT IDDOC FROM DOCUMENTO DOC WHERE TIPOLIB = 5 ) GROUP BY IDCOMP) "
'    Call ExecSQL(DbMain, Q1)


'    Q1 = " INSERT INTO MOVCOMPROBANTEFULL "
'    Q1 = Q1 & " SELECT MCOM.IdMov, MCOM.IdComp, MCOM.IdDoc, MCOM.Orden, MCOM.IdCuenta, MCOM.Debe, MCOM.Haber, MCOM.Glosa, MCOM.idCCosto, MCOM.idAreaNeg, MCOM.IdCartola, MCOM.DeCentraliz, MCOM.DePago, MCOM.DeRemu, MCOM.Nota, MCOM.IdDocCuota, MCOM.IdEmpresa, MCOM.Ano "
'    Q1 = Q1 & " FROM DOCUMENTO DOC, MOVCOMPROBANTE MCOM, COMPROBANTE COM "
'    Q1 = Q1 & " Where DOC.IdDoc = MCOM.IdDoc "
'    Q1 = Q1 & " AND COM.IDCOMP = MCOM.IDCOMP "
'    Q1 = Q1 & " AND TIPOLIB = 5 "
    
'    Q1 = " INSERT INTO MOVCOMPROBANTEFULL "
'    Q1 = Q1 & " ([IdEmpresa],[Ano],[IdComp],[IdDoc],[Orden],[IdCuenta],[Debe],[Haber],[Glosa],[idCCosto],[idAreaNeg],[IdCartola],[DeCentraliz],[DePago],[DeRemu],[Nota],[IdDocCuota]) "
'    Q1 = Q1 & " SELECT MCOM.IdEmpresa, MCOM.Ano, COMF.IdComp, DOCF.IdDoc, MCOM.Orden, MCOM.IdCuenta, MCOM.Debe, MCOM.Haber, MCOM.Glosa, MCOM.idCCosto, MCOM.idAreaNeg, MCOM.IdCartola, MCOM.DeCentraliz, MCOM.DePago, MCOM.DeRemu, MCOM.Nota, MCOM.IdDocCuota "
'    Q1 = Q1 & " FROM (((MOVCOMPROBANTE MCOM INNER JOIN COMPROBANTE COM ON COM.IDCOMP = MCOM.IDCOMP) "
'    Q1 = Q1 & " LEFT JOIN COMPROBANTEFULL COMF ON COMF.IdEmpresa = COM.IdEmpresa AND COMF.Ano = COM.Ano AND COMF.Correlativo = COM.Correlativo AND COMF.TotalDebe = COM.TotalDebe AND COMF.TotalHaber =COM.TotalHaber) "
'    Q1 = Q1 & " LEFT JOIN DOCUMENTO DOC ON DOC.IdDoc = MCOM.IdDoc) "
'    Q1 = Q1 & " LEFT JOIN DOCUMENTOFULL DOCF ON DOCF.NUMDOC = DOC.NUMDOC "
'    Q1 = Q1 & " WHERE MCOM.IdComp IN (SELECT SCOM.IdComp FROM DOCUMENTO SDOC INNER JOIN MOVCOMPROBANTE SCOM ON  SDOC.IDDOC = SCOM.IDDOC WHERE TIPOLIB = 5)  "
'    Call ExecSQL(DbMain, Q1)
    
'    Q1 = " INSERT INTO MOVCOMPROBANTEFULL "
'    Q1 = Q1 & " ([IdEmpresa],[Ano],[IdComp],[IdDoc],[Orden],[IdCuenta],[Debe],[Haber],[Glosa],[idCCosto],[idAreaNeg],[IdCartola],[DeCentraliz],[DePago],[DeRemu],[Nota],[IdDocCuota]) "
'    Q1 = Q1 & " SELECT MCOM.IdEmpresa, MCOM.Ano, MCOM.IdComp, MCOM.IdDoc, MCOM.Orden, MCOM.IdCuenta, MCOM.Debe, MCOM.Haber, MCOM.Glosa, MCOM.idCCosto, MCOM.idAreaNeg, MCOM.IdCartola, MCOM.DeCentraliz, MCOM.DePago, MCOM.DeRemu, MCOM.Nota, MCOM.IdDocCuota "
'    Q1 = Q1 & " FROM MOVCOMPROBANTE MCOM "
'    Q1 = Q1 & " WHERE IDDOC IN (SELECT IDDOC FROM DOCUMENTO DOC  WHERE TIPOLIB = 5) "
'    Call ExecSQL(DbMain, Q1)
    
'    Q1 = " INSERT INTO MOVCOMPROBANTEFULL "
'    Q1 = Q1 & " ([IdEmpresa],[Ano],[IdComp],[IdDoc],[Orden],[IdCuenta],[Debe],[Haber],[Glosa],[idCCosto],[idAreaNeg],[IdCartola],[DeCentraliz],[DePago],[DeRemu],[Nota],[IdDocCuota]) "
'    Q1 = Q1 & " SELECT MCOM.IdEmpresa, MCOM.Ano, MCOM.IdComp, MCOM.IdDoc, MCOM.Orden, MCOM.IdCuenta, MCOM.Debe, MCOM.Haber, MCOM.Glosa, MCOM.idCCosto, MCOM.idAreaNeg, MCOM.IdCartola, MCOM.DeCentraliz, MCOM.DePago, MCOM.DeRemu, MCOM.Nota, MCOM.IdDocCuota "
'    Q1 = Q1 & " FROM MOVCOMPROBANTE MCOM "
'    Q1 = Q1 & " WHERE MCOM.IDCOMP IN (SELECT COM.IDCOMP "
'    Q1 = Q1 & " FROM COMPROBANTE COM, COMPROBANTEFULL COMF "
'    Q1 = Q1 & " WHERE COM.IdEmpresa = COMF.IdEmpresa AND COM.Ano = COMF.Ano AND COM.Correlativo = COMF.Correlativo AND COM.Fecha = COMF.Fecha AND COM.Tipo = COMF.Tipo AND COM.Estado = COMF.Estado AND COM.Glosa = COMF.Glosa AND COM.TotalDebe = COMF.TotalDebe  AND COM.TotalHaber = COMF.TotalHaber AND COM.IdUsuario = COMF.IdUsuario AND COM.FechaCreacion = COMF.FechaCreacion AND COM.ImpResumido = COMF.ImpResumido AND COM.EsCCMM = COMF.EsCCMM  AND COM.TipoAjuste = COMF.TipoAjuste AND COM.OtrosIngEg14TER = COMF.OtrosIngEg14TER ) "
'    Call ExecSQL(DbMain, Q1)
    
'    Q1 = " INSERT INTO DOCUMENTOFULL "
'    Q1 = Q1 & " SELECT DOC.IdDoc, IdCompCent, IdCompPago, 8 as TipoLib, TipoDoc, NumDoc, NumDocHasta, IdEntidad, TipoEntidad, RutEntidad, NombreEntidad, FEmision, FVenc, Descrip, DOC.Estado, Exento, IdCuentaExento, Afecto, IdCuentaAfecto, IVA, IdCuentaIVA, OtroImp, "
'    Q1 = Q1 & " IdCuentaOtroImp , Total, IdCuentaTotal, DOC.IdUsuario, DOC.FechaCreacion, FEmisionOri, CorrInterno, SaldoDoc, FExported, OldIdDoc, DTE, PorcentRetencion, TipoRetencion, MovEdited, OtrosVal, FImporF29, NumDocRef, IdCtaBanco, TipoRelEnt, IdSucursal, TotPagadoAnoAnt, "
'    Q1 = Q1 & " FImportSuc, Giro, FacCompraRetParcial, IVAIrrecuperable, DocOtrosEnAnalitico, OldIdDocTmp, NumFiscImpr, NumInformeZ, CantBoletas, VentasAcumInfZ, IdDocAsoc, PropIVA, ValIVAIrrec, IVAInmueble, DOC.IdEmpresa, FImpFacturacion, CodSIIDTEIvaIrrec, TipoDocAsoc, IVAActFijo, "
'    Q1 = Q1 & " EntRelacionada, NumCuotas, CompraBienRaiz, NumDocAsoc, DTEDocAsoc, IdANegCCosto, UrlDTE, DOC.Ano, CodCtaAfectoOld, CodCtaExentoOld, CodCtaTotalOld, DocOtroEsCargo, ValRet3Porc, IdCuentaRet3Porc, " & ItemData(Cb_Tratamiento) & " as Tratamiento "
'    Q1 = Q1 & " FROM DOCUMENTO DOC, MOVCOMPROBANTE MCOM, COMPROBANTE COM "
'    Q1 = Q1 & " Where DOC.IdDoc = MCOM.IdDoc "
'    Q1 = Q1 & " AND COM.IDCOMP = MCOM.IDCOMP "
'    Q1 = Q1 & " AND TIPOLIB = 5 "
   

    
'    Q1 = " DELETE FROM COMPROBANTE "
'    Q1 = Q1 & " WHERE IDCOMP IN ( "
'    Q1 = Q1 & " SELECT MCOM.IdComp "
'    Q1 = Q1 & " FROM DOCUMENTO DOC, MOVCOMPROBANTE MCOM, COMPROBANTE COM "
'    Q1 = Q1 & " Where DOC.IdDoc = MCOM.IdDoc "
'    Q1 = Q1 & " AND COM.IDCOMP = MCOM.IDCOMP "
'    Q1 = Q1 & " AND TIPOLIB = 5) "
'    Call ExecSQL(DbMain, Q1)
'
'    Q1 = " DELETE FROM MOVCOMPROBANTE "
'    Q1 = Q1 & " WHERE IDCOMP IN ( "
'    Q1 = Q1 & " SELECT MCOM.IdComp "
'    Q1 = Q1 & " FROM DOCUMENTO DOC, MOVCOMPROBANTE MCOM "
'    Q1 = Q1 & " Where DOC.IdDoc = MCOM.IdDoc "
'    Q1 = Q1 & " AND TIPOLIB = 5) "
'    Call ExecSQL(DbMain, Q1)
'
'    Q1 = " DELETE FROM DOCUMENTO "
'    Q1 = Q1 & " Where TipoLib = 5 "
'    Call ExecSQL(DbMain, Q1)
    If Cant > 0 Then
        MsgBox1 "El Traspaso Fue Exitoso.", vbInformation
    Else
        MsgBox1 "No Existen Documentos Seleccionados.", vbInformation
    End If
    Call LoadGrid
    
End Sub

Private Sub SetUpGrid()
   Dim i As Integer
   
   Grid.Cols = NCOLS + 1
    
   Grid.ColWidth(C_IDDOC) = 450
   Grid.ColWidth(C_TIPOLIB) = 830
   Grid.ColWidth(C_IDTIPOLIB) = 0
   Grid.ColWidth(C_TIPODOC) = 450
   Grid.ColWidth(C_NUMDOC) = 950
   'Grid.ColWidth(C_RUT) = 1100
   'Grid.ColWidth(C_ENTIDAD) = 1800
   Grid.ColWidth(C_VALOR) = 1200
'   Grid.ColWidth(C_SALDO) = 1200
'   Grid.ColWidth(C_FEMISION) = 800
'   Grid.ColWidth(C_NUMCUOTAS) = 0
'   Grid.ColWidth(C_IDDOCCUOTA) = 0
'   Grid.ColWidth(C_CUOTA) = 700
'   Grid.ColWidth(C_NUMCUOTA) = 0
'   Grid.ColWidth(C_MONTOCUOTA) = 1200
'   Grid.ColWidth(C_FVENC) = 800
'   Grid.ColWidth(C_ESTADO) = 900
'   Grid.ColWidth(C_IDESTADO) = 0
'   Grid.ColWidth(C_DOCASOC) = 1400
   Grid.ColWidth(C_DESC) = 2800
   Grid.ColWidth(C_CHECK) = 1000
   Grid.ColWidth(C_IDCTABANCO) = 0
   
   
         
   Grid.ColAlignment(C_IDDOC) = flexAlignCenterCenter
   Grid.ColAlignment(C_TIPOLIB) = flexAlignLeftCenter
   Grid.ColAlignment(C_TIPODOC) = flexAlignLeftCenter
   Grid.ColAlignment(C_NUMDOC) = flexAlignRightCenter
   'Grid.ColAlignment(C_RUT) = flexAlignRightCenter
   Grid.ColAlignment(C_CHECK) = flexAlignCenterCenter
   Grid.ColAlignment(C_DESC) = flexAlignLeftCenter
   Grid.ColAlignment(C_VALOR) = flexAlignRightCenter
'   Grid.ColAlignment(C_SALDO) = flexAlignRightCenter
'   Grid.ColAlignment(C_CUOTA) = flexAlignRightCenter
'   Grid.ColAlignment(C_MONTOCUOTA) = flexAlignRightCenter
'   Grid.ColAlignment(C_FEMISION) = flexAlignRightCenter
'   Grid.ColAlignment(C_FVENC) = flexAlignRightCenter
'   Grid.ColAlignment(C_ESTADO) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_IDDOC) = "ID"
   Grid.TextMatrix(0, C_TIPOLIB) = "Libro"
   Grid.TextMatrix(0, C_TIPODOC) = "TD"
   Grid.TextMatrix(0, C_NUMDOC) = "N° Doc."
   'Grid.TextMatrix(0, C_ESTADO) = "Est. Doc."
   'Grid.TextMatrix(0, C_RUT) = "RUT"
   'Grid.TextMatrix(0, C_ENTIDAD) = "Razón Social"
   Grid.TextMatrix(0, C_VALOR) = "Total"
'   Grid.TextMatrix(0, C_SALDO) = "Saldo Doc."
'   Grid.TextMatrix(0, C_FEMISION) = "Emisión"
'   Grid.TextMatrix(0, C_CUOTA) = "Cuota"
'   Grid.TextMatrix(0, C_MONTOCUOTA) = "Monto Cuota"
'   Grid.TextMatrix(0, C_FVENC) = "Vencim."
'   Grid.TextMatrix(0, C_DOCASOC) = "Doc. Asoc."
   Grid.TextMatrix(0, C_DESC) = "Descripción"
   Grid.TextMatrix(0, C_CHECK) = "Seleccionar"
   Grid.TextMatrix(0, C_IDCTABANCO) = "ID Cuenta Contab."
'   Grid.TextMatrix(0, C_TRATAMIENTO) = "Tratamiento"
    
'   If lOper = O_SELECT Then
'      Grid.ColWidth(C_CHECK) = 300
'      Grid.Row = 0
'      Grid.Col = C_CHECK
'      Set Grid.CellPicture = Pc_HdCheck
'      Grid.CellPictureAlignment = flexAlignCenterCenter
'   End If
    
   Call FGrSetup(Grid)
   'Call FGrTotales(Grid, GridTot)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
End Sub


Private Sub LoadGrid(Optional ByVal SelAll As Boolean = False)
   Dim Q1 As String
   Dim Q2 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim Where As String
   Dim IdEnt As Long
   Dim NombEnt As String
   Dim Row As Integer
   Dim Total As Double
   Dim TotSaldo As Double
   Dim NotValidRut As Boolean
   Dim TmpTbl As String
         
   Grid.Redraw = False
   

   Q1 = "Select Doc.IdDoc, Doc.TipoLib, 'OD' as TipoDocDesc, Tip.TipoDoc, Tip.Nombre, Doc.NumDoc, Doc.Total, Doc.Descrip,IdCtaBanco "
   Q1 = Q1 & "From Documento Doc INNER JOIN TipoDocs Tip ON Doc.tipoLib =  Tip.tipoLib AND Doc.TipoDoc = Tip.TipoDoc "
   Q1 = Q1 & " Where Doc.TipoLib = 5 "
   Q1 = Q1 & " and Doc.IdEmpresa = " & gEmpresa.id
   Q1 = Q1 & " and Doc.Ano = " & gEmpresa.Ano
   
   Set Rs = OpenRs(DbMain, Q1)

   Grid.rows = Grid.FixedRows
   i = Grid.FixedRows
   
   Do While Rs.EOF = False
      
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_IDDOC) = vFld(Rs("IdDoc"))
      
'      If IdDoc > 0 And vFld(Rs("IdDoc")) = IdDoc Then
'         Row = i
'      End If
      
      Grid.TextMatrix(i, C_TIPOLIB) = Left(ReplaceStr(gTipoLibNew(IIf(vFld(Rs("TipoLib")) = 8, 6, vFld(Rs("TipoLib")))).Nombre, "Libro de ", ""), 9)
      Grid.TextMatrix(i, C_IDTIPOLIB) = vFld(Rs("TipoLib"))
      Grid.TextMatrix(i, C_TIPODOC) = GetDiminutivoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))
      Grid.TextMatrix(i, C_NUMDOC) = vFld(Rs("NumDoc"))
      
      
      Grid.TextMatrix(i, C_VALOR) = Format(vFld(Rs("Total")), NUMFMT)
      Total = Total + vFld(Rs("Total"))
      
                  
      Grid.TextMatrix(i, C_DESC) = vFld(Rs("Descrip"), True)
      
      If SelAll Then
        Grid.TextMatrix(i, C_CHECK) = "X"
      End If
            
       '2855046 ffv
       Grid.TextMatrix(i, C_IDCTABANCO) = vFld(Rs("IdCtaBanco"))
       '2855046 ffv
        
      Rs.MoveNext
      i = i + 1
   Loop

   Call CloseRs(Rs)
   
   Call FGrVRows(Grid)
   Grid.TopRow = Grid.FixedRows
   
   
   If Row = 0 Then
      Row = Grid.FixedRows
   End If
   
   Call FGrSelRow(Grid, Row)
      
   Grid.Redraw = True
   'Call EnableFrm(False)
   
End Sub



Private Sub Grid_DblClick()
   Dim Row As Integer
   Dim Col As Integer
   
   Row = Grid.MouseRow
   Col = Grid.MouseCol
   
   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Col = C_CHECK Then
        If Grid.TextMatrix(Row, C_CHECK) = "" Then
            Grid.TextMatrix(Row, C_CHECK) = "X"
        Else
            Grid.TextMatrix(Row, C_CHECK) = ""
        End If
   End If
   
End Sub
