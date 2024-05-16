Attribute VB_Name = "MCorrigeBaseSQLServer"
Option Explicit
Private lDbVer As Integer
Dim lUpdOK As Boolean
Public Const MAX_COL = 37 + MAX_COLOTROIMP

Public Sub CorrigeBaseSQLServer()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Rc As Long
   
   lDbVer = 0
   lUpdOK = True
   
   On Error Resume Next
   
   Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO ParamEmpresa (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      lDbVer = 0
   Else
      lDbVer = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)
                  
   If Not CorrigeBaseSQLServer_V701() Then     'agregada 27 dic 2019
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V702() Then     'agregada 6 ene 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V710() Then     'agregada 28 ene 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V711() Then     'agregada 2 mar 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V712() Then     'agregada 6 may 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V713() Then     'agregada 10 jun 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V714() Then     'agregada 5 ago 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V715() Then     'agregada 9 sep 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V716() Then     'agregada 16 sep 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V717() Then     'agregada 23 sep 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V718() Then     'agregada 27 oct 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V719() Then     'agregada 2 nov 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V720() Then     'agregada 16 nov 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V721() Then     'agregada 14 dic 2020
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V722() Then     'agregada 18 ene 2021
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V723() Then     'agregada 22 jun 2021
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V724() Then     'agregada 9 sep 2021
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V725() Then     'agregada 29 sep 2021
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V726() Then     'agregada 18 oct 2021
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V727() Then     'agregada 21 oct 2021
      Exit Sub
   End If
   
   If Not CorrigeBaseSQLServer_V728() Then     'agregada 19 nov 2021 FPG ADO 2678539
      Exit Sub
   End If
   
   If Not CorrigeBaseSQLServer_V729() Then
      Exit Sub
   End If
   
   If Not CorrigeBaseSQLServer_V730() Then     'Percepciones 2699584
      Exit Sub
   End If
               
   If Not CorrigeBaseSQLServer_V731() Then     '3% ret 2864171
      Exit Sub
   End If
   
   '2860036
   If Not CorrigeBaseSQLServer_V732() Then     'Membrete 2860036 FFV
      Exit Sub
   End If
    'fin 2860036
    
    '2861570
   If Not CorrigeBaseSQLServer_V733() Then     'Firmas 2861570 FFV
      Exit Sub
   End If
    'fin 2861570
        
   If Not CorrigeBaseSQLServer_V734() Then     'Firmas 2861570 FFV
      Exit Sub
   End If
   
   If Not CorrigeBaseSQLServer_V735() Then     'ADO 2913643 FPG
      Exit Sub
   End If
   
   If Not CorrigeBaseSQLServer_V736() Then     'agregada 25 ENE 2023 FPR ADO 2862611
      Exit Sub
   End If
   
    If Not CorrigeBaseSQLServer_V737() Then     'agregada 03 ABR 2023 FFV ADO
      Exit Sub
   End If
   
    If Not CorrigeBaseSQLServer_V738() Then     'agregada 03 ABR 2023 FFV ADO 2861733
      Exit Sub
   End If
   
'   '3269719
   If Not CorrigeBaseSQLServer_V739() Then     'agregada 18 OCT 2023 FFV ADO 3269719
      Exit Sub
   End If
   '3269719
   
   If Not CorrigeBaseSQLServer_V740() Then     'agregada 17 OCT 2023 FPR ADO 3217833
      Exit Sub
   End If
               
   If lDbVer > 741 Then
      MsgBox1 "¡ ATENCION !" & vbCrLf & vbCrLf & "La base de datos corresponde a una versión posterior de este programa." & vbCrLf & "Debe actualizar el programa antes de continuar, de lo contrario podría dañar la información..", vbCritical
      Call CloseDb(DbMain)
      End
   End If
   
  
End Sub

'3217833
Public Function CorrigeBaseSQLServer_V740() As Boolean   '3217833 FPR
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 740-----------------------------------
   If lDbVer = 740 And lUpdOK = True Then
   
      Q1 = "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'Tracking_Documento')"
      Q1 = Q1 & "  BEGIN"
      Q1 = Q1 & " CREATE TABLE [dbo].[Tracking_Documento]("
      Q1 = Q1 & "       [IdDoc] [int] NOT NULL,"
      Q1 = Q1 & "       [FechaHora] [DateTime] NOT NULL,"
      Q1 = Q1 & "       [IdEmpresa] [int] NULL,"
      Q1 = Q1 & "       [Ano] [smallint] NULL,"
      Q1 = Q1 & "       [IdCompCent] [int] NULL,"
      Q1 = Q1 & "       [IdCompPago] [int] NULL,"
      Q1 = Q1 & "       [TipoLib] [tinyint] NULL,"
      Q1 = Q1 & "       [TipoDoc] [tinyint] NULL,"
      Q1 = Q1 & "       [NumDoc] [varchar](20) NULL,"
      Q1 = Q1 & "       [NumDocHasta] [varchar](20) NULL,"
      Q1 = Q1 & "       [IdEntidad] [int] NULL,"
      Q1 = Q1 & "       [TipoEntidad] [smallint] NULL,"
      Q1 = Q1 & "       [RutEntidad] [varchar](12) NULL,"
      Q1 = Q1 & "       [NombreEntidad] [varchar](50) NULL,"
      Q1 = Q1 & "       [FEmision] [int] NULL,"
      Q1 = Q1 & "       [FVenc] [int] NULL,"
      Q1 = Q1 & "       [Descrip] [varchar](100) NULL,"
      Q1 = Q1 & "       [Estado] [smallint] NULL,"
      Q1 = Q1 & "       [Exento] [float] NULL,"
      Q1 = Q1 & "       [IdCuentaExento] [int] NULL,"
      Q1 = Q1 & "       [Afecto] [float] NULL,"
      Q1 = Q1 & "       [IdCuentaAfecto] [int] NULL,"
      Q1 = Q1 & "       [IVA] [float] NULL,"
      Q1 = Q1 & "       [IdCuentaIVA] [int] NULL,"
      Q1 = Q1 & "       [OtroImp] [float] NULL,"
      Q1 = Q1 & "       [IdCuentaOtroImp] [int] NULL,"
      Q1 = Q1 & "       [Total] [float] NULL,"
      Q1 = Q1 & "       [IdCuentaTotal] [int] NULL,"
      Q1 = Q1 & "       [IdUsuario] [int] NULL,"
      Q1 = Q1 & "       [FechaCreacion] [int] NULL,"
      Q1 = Q1 & "       [FEmisionOri] [int] NULL,"
      Q1 = Q1 & "       [CorrInterno] [int] NULL,"
      Q1 = Q1 & "       [SaldoDoc] [float] NULL,"
      Q1 = Q1 & "       [FExported] [int] NULL,"
      Q1 = Q1 & "       [OldIdDoc] [int] NULL,"
      Q1 = Q1 & "       [DTE] [bit] NULL,"
      Q1 = Q1 & "       [PorcentRetencion] [tinyint] NULL,"
      Q1 = Q1 & "       [TipoRetencion] [tinyint] NULL,"
      Q1 = Q1 & "       [MovEdited] [bit] NULL,"
      Q1 = Q1 & "       [OtrosVal] [float] NULL,"
      Q1 = Q1 & "       [FImporF29] [int] NULL,"
      Q1 = Q1 & "       [NumDocRef] [varchar](20) NULL,"
      Q1 = Q1 & "       [IdCtaBanco] [int] NULL,"
      Q1 = Q1 & "       [TipoRelEnt] [smallint] NULL,"
      Q1 = Q1 & "       [IdSucursal] [int] NULL,"
      Q1 = Q1 & "       [TotPagadoAnoAnt] [float] NULL,"
      Q1 = Q1 & "       [FImportSuc] [int] NULL,"
      Q1 = Q1 & "       [Giro] [bit] NULL,"
      Q1 = Q1 & "       [FacCompraRetParcial] [bit] NULL,"
      Q1 = Q1 & "       [IVAIrrecuperable] [smallint] NULL,"
      Q1 = Q1 & "       [DocOtrosEnAnalitico] [bit] NULL,"
      Q1 = Q1 & "       [OldIdDocTmp] [int] NULL,"
      Q1 = Q1 & "       [NumFiscImpr] [varchar](20) NULL,"
      Q1 = Q1 & "       [NumInformeZ] [varchar](20) NULL,"
      Q1 = Q1 & "       [CantBoletas] [int] NULL,"
      Q1 = Q1 & "       [VentasAcumInfZ] [float] NULL,"
      Q1 = Q1 & "       [IdDocAsoc] [int] NULL,"
      Q1 = Q1 & "       [PropIVA] [smallint] NULL,"
      Q1 = Q1 & "       [ValIVAIrrec] [float] NULL,"
      Q1 = Q1 & "       [IVAInmueble] [bit] NULL,"
      Q1 = Q1 & "       [FImpFacturacion] [int] NULL,"
      Q1 = Q1 & "       [CodSIIDTEIVAIrrec] [smallint] NULL,"
      Q1 = Q1 & "       [TipoDocAsoc] [smallint] NULL,"
      Q1 = Q1 & "       [IVAActFijo] [float] NULL,"
      Q1 = Q1 & "       [EntRelacionada] [bit] NULL,"
      Q1 = Q1 & "       [NumCuotas] [smallint] NULL,"
      Q1 = Q1 & "       [CompraBienRaiz] [bit] NULL,"
      Q1 = Q1 & "       [NumDocAsoc] [varchar](20) NULL,"
      Q1 = Q1 & "       [DTEDocAsoc] [bit] NULL,"
      Q1 = Q1 & "       [IdANegCCosto] [varchar](20) NULL,"
      Q1 = Q1 & "       [UrlDTE] [varchar](250) NULL,"
      Q1 = Q1 & "       [CodCtaAfectoOld] [varchar](15) NULL,"
      Q1 = Q1 & "       [CodCtaExentoOld] [varchar](15) NULL,"
      Q1 = Q1 & "       [CodCtaTotalOld] [varchar](15) NULL,"
      Q1 = Q1 & "       [DocOtroEsCargo] [tinyint] NULL,"
      Q1 = Q1 & "       [ValRet3Porc] [float] NULL,"
      Q1 = Q1 & "       [IdCuentaRet3Porc] [int] NULL,"
      Q1 = Q1 & "       [Tratamiento] [float] NULL,"
      'Q1 = Q1 & "       [IdTras] [int] NULL,"
      Q1 = Q1 & "       [Origen] [varchar](250) NULL,"
      Q1 = Q1 & "       [Query] [varchar](MAX) NULL,"
      Q1 = Q1 & "       [Vigente] [int] NULL," 'Para saber si el cambio lo hizo el cliente o el sistema
      Q1 = Q1 & "       [FormaIngreso] [int] NULL,"
      Q1 = Q1 & "       [Ajuste] [int] NULL,"
      Q1 = Q1 & "    CONSTRAINT [Tracking_Documento_IdDoc_PK] PRIMARY KEY CLUSTERED"
      Q1 = Q1 & "   ("
      Q1 = Q1 & "       [IdDoc] ASC,"
      Q1 = Q1 & "       [FechaHora]"
      Q1 = Q1 & "   )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]"
      Q1 = Q1 & "   ) ON [PRIMARY]"
      Q1 = Q1 & "   End"
      'Q1 = Q1 & "   GO"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'Tracking_MovDocumento')"
      Q1 = Q1 & "  BEGIN"
      Q1 = Q1 & "  CREATE TABLE [dbo].[Tracking_MovDocumento]("
      Q1 = Q1 & "      [IdMovDoc] [int] NOT NULL,"
      Q1 = Q1 & "      [FechaHora] [DateTime] NOT NULL,"
      Q1 = Q1 & "      [IdEmpresa] [int] NULL,"
      Q1 = Q1 & "      [Ano] [smallint] NULL,"
      Q1 = Q1 & "      [IdDoc] [int] NULL,"
      Q1 = Q1 & "      [IdCompCent] [int] NULL,"
      Q1 = Q1 & "      [IdCompPago] [int] NULL,"
      Q1 = Q1 & "      [Orden] [tinyint] NULL,"
      Q1 = Q1 & "      [IdCuenta] [int] NULL,"
      Q1 = Q1 & "      [Debe] [float] NULL,"
      Q1 = Q1 & "      [Haber] [float] NULL,"
      Q1 = Q1 & "      [Glosa] [varchar](50) NULL,"
      Q1 = Q1 & "      [IdTipoValLib] [smallint] NULL,"
      Q1 = Q1 & "      [EsTotalDoc] [bit] NULL,"
      Q1 = Q1 & "      [IdCCosto] [int] NULL,"
      Q1 = Q1 & "      [IdAreaNeg] [int] NULL,"
      Q1 = Q1 & "      [Tasa] [real] NULL,"
      Q1 = Q1 & "      [EsRecuperable] [bit] NULL,"
      Q1 = Q1 & "      [CodSIIDTE] [char](2) NULL,"
      Q1 = Q1 & "      [CodCuentaOld] [varchar](15) NULL,"
      'Q1 = Q1 & "      [IdTras] [int] NULL,"
      Q1 = Q1 & "      [Origen] [varchar](250) NULL,"
      Q1 = Q1 & "      [Query] [varchar](MAX) NULL,"
      Q1 = Q1 & "      [Vigente] [int] NULL,"
      Q1 = Q1 & "      [FormaIngreso] [int] NULL,"
      Q1 = Q1 & "      [Ajuste] [int] NULL,"
      Q1 = Q1 & "   CONSTRAINT [Tracking_MovDocumento_IdMovDoc_PK] PRIMARY KEY CLUSTERED"
      Q1 = Q1 & "  ([IdMovDoc] ASC,[FechaHora]"
      Q1 = Q1 & "  )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]"
      Q1 = Q1 & "  ) ON [PRIMARY] END "
      Call ExecSQL(DbMain, Q1)

    
      Q1 = "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'Tracking_Comprobante')"
      Q1 = Q1 & "  BEGIN"
      Q1 = Q1 & "  CREATE TABLE [dbo].[Tracking_Comprobante]("
      Q1 = Q1 & "  [IdComp] [int] NOT NULL,"
      Q1 = Q1 & "  [FechaHora] [datetime] NOT NULL,"
      Q1 = Q1 & "  [IdEmpresa] [int] NULL,"
      Q1 = Q1 & "  [Ano] [smallint] NULL,"
      Q1 = Q1 & "  [Correlativo] [int] NULL,"
      Q1 = Q1 & "  [Fecha] [int] NULL,"
      Q1 = Q1 & "  [Tipo] [tinyint] NULL,"
      Q1 = Q1 & "  [Estado] [tinyint] NULL,"
      Q1 = Q1 & "  [Glosa] [varchar](100) NULL,"
      Q1 = Q1 & "  [TotalDebe] [float] NULL,"
      Q1 = Q1 & "  [TotalHaber] [float] NULL,"
      Q1 = Q1 & "  [IdUsuario] [int] NULL,"
      Q1 = Q1 & "  [FechaCreacion] [int] NULL,"
      Q1 = Q1 & "  [ImpResumido] [bit] NULL,"
      Q1 = Q1 & "  [EsCCMM] [bit] NULL,"
      Q1 = Q1 & "  [FechaImport] [int] NULL,"
      Q1 = Q1 & "  [TipoAjuste] [tinyint] NULL,"
      Q1 = Q1 & "  [OtrosIngEg14TER] [bit] NULL,"
      Q1 = Q1 & "  [Origen] [varchar](250) NULL,"
      Q1 = Q1 & "  [Query] [varchar](MAX) NULL,"
      Q1 = Q1 & "  [Vigente] [int] NULL,"
      Q1 = Q1 & "  [FormaIngreso] [int] NULL,"
      Q1 = Q1 & "  [Ajuste] [int] NULL,"
      Q1 = Q1 & "  CONSTRAINT [Tracking_Comprobante_IdComp_PK] PRIMARY KEY CLUSTERED"
      Q1 = Q1 & "  ("
      Q1 = Q1 & "  [IdComp] ASC,"
      Q1 = Q1 & "  [FechaHora] Asc"
      Q1 = Q1 & "  )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]"
      Q1 = Q1 & "  ) ON [PRIMARY]"
      Q1 = Q1 & "   END"
      'Q1 = Q1 & "   GO"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'Tracking_MovComprobante')"
      Q1 = Q1 & "  BEGIN"
      Q1 = Q1 & "  CREATE TABLE [dbo].[Tracking_MovComprobante]("
      Q1 = Q1 & "  [IdMov] [int] NOT NULL,"
      Q1 = Q1 & "  [FechaHora] [datetime] NOT NULL,"
      Q1 = Q1 & "  [IdEmpresa] [int] NULL,"
      Q1 = Q1 & "  [Ano] [smallint] NULL,"
      Q1 = Q1 & "  [IdComp] [int] NULL,"
      Q1 = Q1 & "  [IdDoc] [int] NULL,"
      Q1 = Q1 & "  [Orden] [int] NULL,"
      Q1 = Q1 & "  [IdCuenta] [int] NULL,"
      Q1 = Q1 & "  [Debe] [float] NULL,"
      Q1 = Q1 & "  [Haber] [float] NULL,"
      Q1 = Q1 & "  [Glosa] [varchar](50) NULL,"
      Q1 = Q1 & "  [idCCosto] [int] NULL,"
      Q1 = Q1 & "  [idAreaNeg] [int] NULL,"
      Q1 = Q1 & "  [IdCartola] [int] NULL,"
      Q1 = Q1 & "  [DeCentraliz] [bit] NULL,"
      Q1 = Q1 & "  [DePago] [bit] NULL,"
      Q1 = Q1 & "  [DeRemu] [bit] NULL,"
      Q1 = Q1 & "  [Nota] [varchar](120) NULL,"
      Q1 = Q1 & "  [IdDocCuota] [int] NULL,"
      Q1 = Q1 & "  [Origen] [varchar](250) NULL,"
      Q1 = Q1 & "  [Query] [varchar](MAX) NULL,"
      Q1 = Q1 & "  [Vigente] [int] NULL,"
      Q1 = Q1 & "  [FormaIngreso] [int] NULL,"
      Q1 = Q1 & "  [Ajuste] [int] NULL,"
      Q1 = Q1 & "  CONSTRAINT [Tracking_MovComprobante_IdMov_PK] PRIMARY KEY CLUSTERED"
      Q1 = Q1 & "  ("
      Q1 = Q1 & "  [IdMov] ASC,"
      Q1 = Q1 & "  [FechaHora] Asc"
      Q1 = Q1 & "  )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]"
      Q1 = Q1 & "  ) ON [PRIMARY] END "
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVer = 741
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V740 = lUpdOK

End Function
'fin 3217833

'3269719
Public Function CorrigeBaseSQLServer_V739() As Boolean   '2861570 FFV
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 739-----------------------------------
   If lDbVer = 739 And lUpdOK = True Then

      Q1 = "CREATE TABLE [dbo].[tbl_Comp_Centra_Full]( "
      Q1 = Q1 & "[IdComp] [int] NULL, "
      Q1 = Q1 & "[IdEmpresa] [int] NULL, "
      Q1 = Q1 & "[Tipo] [varchar](50) NULL, "
      Q1 = Q1 & "[Fecha] [int] NULL, "
      Q1 = Q1 & "[ano] [varchar](4) NULL );"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVer = 740
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V739 = lUpdOK

End Function
'3269719



'2861733 tema 2

Public Function CorrigeBaseSQLServer_V738() As Boolean   'Agregada 14 dic 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double

   On Error Resume Next
   
   '--------------------- Versión 738 -----------------------------------
   If lDbVer = 738 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo idCCosto a MovActivoFijo
      'Set Tbl = DbMain.TableDefs("MovActivoFijo")
     
      ERR.Clear
   
       Q1 = "ALTER TABLE MovActivoFijo ADD idCCosto INT  NULL;"
      Call ExecSQL(DbMain, Q1)
   
       'Agregamos campo IdAreaNeg a MovActivoFijo
       ERR.Clear
      'Tbl.Fields.Append Tbl.CreateField("IdAreaNeg", dbLong)
      
      Q1 = "ALTER TABLE MovActivoFijo ADD IdAreaNeg INT  NULL;"
      Call ExecSQL(DbMain, Q1)
          
      Q1 = "DROP INDEX idx_IdCCosto ON MovActivoFijo "
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX idx_IdCCosto ON MovActivoFijo (IdCCosto) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "DROP INDEX idx_IdAreaNeg ON MovActivoFijo "
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX idx_IdAreaNeg ON MovActivoFijo (IdAreaNeg) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      If lUpdOK Then
         lDbVer = 739
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V738 = lUpdOK

End Function
'2861733 tema 2

'3043065 SF 14006128
Public Function CorrigeBaseSQLServer_V737() As Boolean   'Agregada 14 dic 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim CapPropio As Double
   
    Dim idxDocumento As Index

   On Error Resume Next
   
   '--------------------- Versión 737 -----------------------------------
   If lDbVer = 737 And lUpdOK = True Then
   
       ERR.Clear
       
    'Call CloseDb(DbMain)
   ' Call OpenDbEmp(gEmpresa.Rut, gEmpresa.Ano)
    
       Q1 = "DROP INDEX idx_TipoLibDoc ON Documento "
      Call ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE INDEX idx_TipoLibDoc ON Documento (TipoLib) "
      Rc = ExecSQL(DbMain, Q1, True)
      
       Q1 = "DROP INDEX idx_FEmision ON Documento "
      Call ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE INDEX idx_FEmision ON Documento (FEmision) "
      Rc = ExecSQL(DbMain, Q1, False)
      
       Q1 = "DROP INDEX idx_Exento ON Documento "
      Call ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE INDEX idx_Exento ON Documento (Exento) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "DROP INDEX idx_TipoRetencion ON Documento "
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX idx_TipoRetencion ON Documento (TipoRetencion) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "DROP INDEX idx_PorcentRetencion ON Documento "
      Call ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE INDEX idx_PorcentRetencion ON Documento (PorcentRetencion) "
      Rc = ExecSQL(DbMain, Q1, False)
      
      Q1 = "DROP INDEX idx_IdEmpresa ON Documento "
      Call ExecSQL(DbMain, Q1, False)
      
       Q1 = "CREATE INDEX idx_IdEmpresa ON Documento (IdEmpresa) "
      Rc = ExecSQL(DbMain, Q1, False)
      
       Q1 = "DROP INDEX idx_AnoDOC ON Documento "
      Call ExecSQL(DbMain, Q1, False)
      
      Q1 = "CREATE INDEX idx_AnoDOC ON Documento (Ano) "
      Rc = ExecSQL(DbMain, Q1, False)
      
       If ERR <> 0 Then
         MsgBeep vbExclamation
         MsgBox "Error " & ERR & ", " & Error, vbExclamation
         lUpdOK = False
      End If
      
      If lUpdOK Then
         lDbVer = 738
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V737 = lUpdOK

End Function
'3043065 SF 14006128


Public Function CorrigeBaseSQLServer_V736() As Boolean   'Agregada 25 ENE 2023 FPG
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 736 -----------------------------------
   If lDbVer <= 736 And lUpdOK = True Then
   
      ERR.Clear
      
        Q1 = " INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) VALUES (" & gEmpresa.id & "," & gEmpresa.Ano & ",'DATOSII',1,'1') "
        Call ExecSQL(DbMain, Q1)
        
        Q1 = " INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) VALUES (" & gEmpresa.id & "," & gEmpresa.Ano & ",'DATOSII',2,'0') "
        Call ExecSQL(DbMain, Q1)
        
        Q1 = " INSERT INTO ParamEmpresa (IdEmpresa, Ano, Tipo, Codigo, Valor) VALUES (" & gEmpresa.id & "," & gEmpresa.Ano & ",'DATOSII',3,'0') "
        Call ExecSQL(DbMain, Q1)
                                  
      If lUpdOK Then
         lDbVer = 737
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V736 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V735() As Boolean   'Agregada 18 oct 2022
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 735 -----------------------------------
   If lDbVer <= 735 And lUpdOK = True Then
   
      ERR.Clear
      
      Q1 = "ALTER TABLE Empresa ADD CodArea NUMERIC  NULL;"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE Empresa ADD Celular NUMERIC  NULL;"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE Empresa ADD Villa VARCHAR (80)  NULL;"
      Call ExecSQL(DbMain, Q1)
                                  
      If lUpdOK Then
         lDbVer = 736
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V735 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V734() As Boolean   'Se agrega para llevar un orden
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 734 -----------------------------------
   If lDbVer = 734 And lUpdOK = True Then
        
        
        ERR.Clear

        Q1 = "INSERT INTO Param "
        Q1 = Q1 & " (Tipo, Codigo, Valor)"
        Q1 = Q1 & " VALUES( 'TIPOLIB', 8, 'Otros Documentos Full')"
                                                                                         
                                                                                        
        Call ExecSQL(DbMain, Q1)

        Q1 = "INSERT INTO Param "
        Q1 = Q1 & " (Tipo, Codigo, Valor)"
        Q1 = Q1 & " VALUES( 'TIPOLIBCOD', 8, 'LIBOTROFULL')"
        Call ExecSQL(DbMain, Q1)

        Q1 = "INSERT INTO TipoDocs (TipoLib, TipoDoc, Nombre, Diminutivo, Atributo, TipoDocFijo, TieneAfecto, TieneExento, ExigeRUT, EsRebaja, DocImpExp, DocBoletas, CodDocSII, CodDocDTESII)"
         Q1 = Q1 & " VALUES(8,1, 'Otros Documentos Full', 'ODF', 'ACTIVO', -1, 0, 1, 0, 0, 0, 0, '', '')"
         Call ExecSQL(DbMain, Q1)

        Q1 = "INSERT INTO Param "
        Q1 = Q1 & " (Tipo, Codigo, Valor)"
        Q1 = Q1 & " VALUES( 'TRATAMIENTO', 1, 'Activo')"
        Call ExecSQL(DbMain, Q1)

        Q1 = "INSERT INTO Param "
        Q1 = Q1 & " (Tipo, Codigo, Valor)"
        Q1 = Q1 & " VALUES( 'TRATAMIENTO', 2, 'Pasivo')"
        Call ExecSQL(DbMain, Q1)
        
        Q1 = "ALTER TABLE Documento ADD Tratamiento FLOAT NULL;"
        Call ExecSQL(DbMain, Q1)
        
        'Call CreateTablasODF

      If lUpdOK Then
         lDbVer = 735
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If

   CorrigeBaseSQLServer_V734 = lUpdOK

End Function

'2861570
Public Function CorrigeBaseSQLServer_V733() As Boolean   '2861570 FFV
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 733-----------------------------------
   If lDbVer = 733 And lUpdOK = True Then

      Q1 = "CREATE TABLE [dbo].[Firmas]( "
      Q1 = Q1 & "[Patch] [varchar](255) NULL, "
      Q1 = Q1 & "[IdEmpresa] [numeric](18) NULL, "
      Q1 = Q1 & "[Tipo] [varchar](50) NULL, "
      Q1 = Q1 & "[ano] [varchar](4) NULL );"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVer = 734
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V733 = lUpdOK

End Function
'fin 2861570


'2860036
Public Function CorrigeBaseSQLServer_V732() As Boolean   '2860036 FFV
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 732-----------------------------------
   If lDbVer = 732 And lUpdOK = True Then

      Q1 = "CREATE TABLE [dbo].[Membrete]( "
      Q1 = Q1 & "[TituloMembrete1] [varchar](50) NULL, "
      Q1 = Q1 & "[TituloMembrete2] [varchar](50) NULL, "
      Q1 = Q1 & "[Texto1] [varchar](50) NULL, "
      Q1 = Q1 & "[Texto2] [varchar](50) NULL, "
      Q1 = Q1 & "[IdEmpresa] [numeric](18) NULL );"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVer = 733
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V732 = lUpdOK

End Function
'fin 2860036


'2864171
Public Function CorrigeBaseSQLServer_V731() As Boolean   'agregada 8 ago 2022 ffv
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String


   On Error Resume Next

   '--------------------- Versión 731 -----------------------------------

   If lDbVer = 731 And lUpdOK = True Then
           
      'Insertamos Retención 3% préstamo solidario a libro de retenciones
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
      Q1 = Q1 & " VALUES(" & LIB_RETEN & "," & LIBRETEN_RET3PORC & ", 'Retención 3%', ' ', ' ', 0, ' ', ' ', 'Retención 3%',' ',' ', 5, 0, 0, ' ', 'Retención 3% Prést. Sol.')"
      Call ExecSQL(DbMain, Q1)
      
      
           
     If lUpdOK Then
         lDbVer = 732
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseSQLServer_V731 = lUpdOK
   
End Function
'2864171



Public Function CorrigeBaseSQLServer_V730() As Boolean   'Se agrega para llevar un orden
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 730 -----------------------------------
   If lDbVer = 730 And lUpdOK = True Then

      Q1 = "CREATE TABLE [dbo].[Percepciones]( "
      Q1 = Q1 & "[IDPerc] [numeric](18, 0) PRIMARY KEY, "
      Q1 = Q1 & "[IdComp] [numeric](18, 0) NULL, "
      Q1 = Q1 & "[Orden] [numeric](18, 0) NULL, "
      Q1 = Q1 & "[IdCuenta] [numeric](18, 0) NULL, "
      Q1 = Q1 & "[IdEmpresa] [numeric](18, 0) NULL, "
      Q1 = Q1 & "[Ano] [int] NULL, "
      Q1 = Q1 & "[Fecha] [int] NULL, "
      Q1 = Q1 & "[NumCertificado] [numeric](18, 0) NULL, "
      Q1 = Q1 & "[RutEmpresa] [numeric](18, 0) NULL, "
      Q1 = Q1 & "[Regimen] [numeric](18, 0) NULL, "
      Q1 = Q1 & "[Contabilizacion] [numeric](18, 0) NULL, "
      Q1 = Q1 & "[TasaTef] [decimal](18, 6) NULL, "
      Q1 = Q1 & "[TasaTex] [decimal](18, 6) NULL, "
      Q1 = Q1 & "[Percepciones] [numeric](18, 0) NULL );"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE TABLE [dbo].[DetPercepciones]( "
      Q1 = Q1 & "[IDPerc] [numeric](18, 0) NOT NULL, "
      Q1 = Q1 & "[CodDet] [numeric](18, 0) NOT NULL, "
      Q1 = Q1 & "[Valor] [numeric](18, 0) NULL , "
      Q1 = Q1 & "PRIMARY KEY (IDPerc, CodDet)) "
      Call ExecSQL(DbMain, Q1)
      
      
      Q1 = "ALTER TABLE cuentas ADD Percepcion tinyint NULL;"
      Call ExecSQL(DbMain, Q1)



      If lUpdOK Then
         lDbVer = 731
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V730 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V729() As Boolean   'Se agrega para llevar un orden
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 729 -----------------------------------
   If lDbVer = 729 And lUpdOK = True Then

      If lUpdOK Then
         lDbVer = 730
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V729 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V728() As Boolean   'Agregada 19 nov 2021 FPG Solicitado por Victor Morales
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 728 -----------------------------------
   If lDbVer = 728 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo FDesde3Porc a tabla Entidades
      Q1 = "ALTER TABLE Entidades ADD FDesde3Porc INT NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo FHasta3Porc a tabla Entidades
      ERR.Clear
      Q1 = "ALTER TABLE Entidades ADD FHasta3Porc INT NULL;"
      Call ExecSQL(DbMain, Q1)
     

      If lUpdOK Then
         lDbVer = 729
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V728 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V727() As Boolean   'Agregada 21 oct 2021
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 727 -----------------------------------
   If lDbVer = 727 And lUpdOK = True Then
   
      ERR.Clear
            
      'Actualizamos EsTotalDoc por si acaso  FCA 9 sep 2021
      'Para solucionar tema de cliente que importó documentos desde Facturación
      'No es necesario en este caso (SQL SERVER) dado que no tenemos clientes de Facturación para SQL Server

      If lUpdOK Then
         lDbVer = 728
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V727 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V726() As Boolean   'Agregada 18 oct 2021
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 726 -----------------------------------
   If lDbVer = 726 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo Ret3Porc a tabla Entidades
      Q1 = "ALTER TABLE Entidades ADD Ret3Porc BIT NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo ValRet3Porc a tabla Documento
      ERR.Clear
      Q1 = "ALTER TABLE Documento ADD ValRet3Porc FLOAT NULL;"
      Call ExecSQL(DbMain, Q1)
     

      'Agregamos campo IdCuentaRet3Porc a tabla Documento
      ERR.Clear
      Q1 = "ALTER TABLE Documento ADD IdCuentaRet3Porc INT NULL;"
      Call ExecSQL(DbMain, Q1)


      If lUpdOK Then
         lDbVer = 727
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V726 = lUpdOK

End Function



Public Function CorrigeBaseSQLServer_V725() As Boolean   'Agregada 9 sep 2021
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 724 -----------------------------------
   If lDbVer = 725 And lUpdOK = True Then
   
      ERR.Clear

      Q1 = "CREATE TABLE AjusteIVAMensual ("
      Q1 = Q1 & " IdEmpresa int,"
      Q1 = Q1 & " Ano smallint,"
      Q1 = Q1 & " Mes tinyint, "
      Q1 = Q1 & " Valor float "
      Q1 = Q1 & ");"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVer = 726
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V725 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V724() As Boolean   'Agregada 9 sep 2021
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 724 -----------------------------------
   If lDbVer = 724 And lUpdOK = True Then
   
      ERR.Clear
            
      'Actualizamos EsTotalDoc por si acaso  FCA 9 sep 2021
      'Para solucionar tema de cliente que importó documentos desde Facturación
      'No es necesario en este caso (SQL SERVER) dado que no tenemos clientes de Facturación para SQL Server

      If lUpdOK Then
         lDbVer = 725
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V724 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V723() As Boolean   'Agregada 22 jun 2021
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 723 -----------------------------------
   If lDbVer <= 723 And lUpdOK = True Then
   
      ERR.Clear
            
      'Agregamos campo MontoAfectaBaseImp a tabla LibroCaja
      Q1 = "ALTER TABLE LibroCaja ADD MontoAfectaBaseImp FLOAT NULL;"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVer = 724
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V723 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V722() As Boolean   'agregada 18 ene 2021
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 722 -----------------------------------

   If lDbVer = 722 And lUpdOK = True Then
   
      'Agregamos tabla CtasAjustesExContRLI
      Q1 = "CREATE TABLE BaseImponible14D ("
      Q1 = Q1 & " IdBaseImponible14D int IDENTITY (1,1) NOT NULL,"
      Q1 = Q1 & " IdEmpresa int,"
      Q1 = Q1 & " Ano smallint, "
      Q1 = Q1 & " Tipo tinyint, "
      Q1 = Q1 & " Nivel tinyint, "
      Q1 = Q1 & " Codigo smallint,"
      Q1 = Q1 & " Fecha int,"
      Q1 = Q1 & " Valor float,"
      Q1 = Q1 & " CONSTRAINT IdxBaseImponible14D PRIMARY KEY (IdBaseImponible14D) "
      Q1 = Q1 & ");"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE INDEX IdxEmpAno ON BaseImponible14D (IdEmpresa, Ano, Codigo, Fecha)"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVer = 723
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseSQLServer_V722 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V721() As Boolean   'Agregada 14 dic 2020
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 721 -----------------------------------
   If lDbVer <= 721 And lUpdOK = True Then
   
      ERR.Clear
            
      'Agregamos campo DepLey21256 a MovActivoFijo
      Q1 = "ALTER TABLE Entidades ADD FranqTribEnt TINYINT NULL;"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVer = 722
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V721 = lUpdOK

End Function


Public Function CorrigeBaseSQLServer_V720() As Boolean   'Agregada 16 nov 2020
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 720 -----------------------------------
   If lDbVer <= 720 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agrandamos campo Orden en tabla CT_MovComprobante
      Q1 = "ALTER TABLE CT_MovComprobante DROP CONSTRAINT DF_CT_MovComprobante_Orden;"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE CT_MovComprobante ALTER COLUMN Orden SMALLINT;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo DepLey21256 a MovActivoFijo
      Q1 = "ALTER TABLE MovActivoFijo ADD DepLey21256 TINYINT NULL;"
      Call ExecSQL(DbMain, Q1)

      'Agregamos campo DepLey21256 a MovActivoFijo
      Q1 = "ALTER TABLE MovActivoFijo ADD DepLey21256Hist TINYINT NULL;"
      Call ExecSQL(DbMain, Q1)

                                  
      If lUpdOK Then
         lDbVer = 721
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V720 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V719() As Boolean   'Agregada 2 nov 2020
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
   Dim CapPropio As Double
  

   On Error Resume Next
   
   '--------------------- Versión 719 -----------------------------------
   If lDbVer <= 719 And lUpdOK = True Then
   
      ERR.Clear
      
      'copiamos Capital Ptopio Tributario desde tabla ParamEmpresa a EmpresasAno dado que ahor ase requiere revisar años antriores en el informe de CPS
      Q1 = "SELECT Valor FROM ParamEmpresa WHERE Tipo = 'CAPPROPIO' AND IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
      Set Rs = OpenRs(DbMain, Q1)
      If Not Rs.EOF Then
         CapPropio = Val(vFld(Rs("Valor")))
      End If
      Call CloseRs(Rs)
      
      If CapPropio <> 0 Then
         Q1 = "UPDATE EmpresasAno SET CPS_CapPropioTrib = " & CapPropio
         Q1 = Q1 & " WHERE IdEmpresa = " & gEmpresa.id & " AND Ano = " & gEmpresa.Ano
         Call ExecSQL(DbMain, Q1)
      End If
                                  
      If lUpdOK Then
         lDbVer = 720
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V719 = lUpdOK

End Function


Public Function CorrigeBaseSQLServer_V718() As Boolean   'Agregada 27 oct 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 718 -----------------------------------
   If lDbVer <= 718 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo Descrip a DetCapPropioSimpl
      Q1 = "ALTER TABLE DetCapPropioSimpl ADD Descrip VARCHAR (80)  NULL;"
      Call ExecSQL(DbMain, Q1)
                                  
      If lUpdOK Then
         lDbVer = 719
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V718 = lUpdOK

End Function
            

Public Function CorrigeBaseSQLServer_V717() As Boolean   'agregada 23 sept 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 717 -----------------------------------

   If lDbVer = 717 And lUpdOK = True Then
   
      'Agregamos tabla CtasAjustesExContRLI
      Q1 = "CREATE TABLE DetCapPropioSimpl ("
      Q1 = Q1 & " IdDetCapPropioSimpl int IDENTITY (1,1) NOT NULL,"
      Q1 = Q1 & " IdEmpresa int,"
      Q1 = Q1 & " Ano smallint, "
      Q1 = Q1 & " TipoDetCPS tinyint, "
      Q1 = Q1 & " IngresoManual bit, "
      Q1 = Q1 & " IdCuenta int,"
      Q1 = Q1 & " CodCuenta varchar(15),"
      Q1 = Q1 & " Fecha int,"
      Q1 = Q1 & " IdMovComp Int,"
      Q1 = Q1 & " Valor float,"
      Q1 = Q1 & " CONSTRAINT IdxDetCapPropioSimpl PRIMARY KEY (IdDetCapPropioSimpl) "
      Q1 = Q1 & ");"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE INDEX IdxEmpAno ON DetCapPropioSimpl (IdEmpresa, Ano, TipoDetCPS, IngresoManual, Fecha, CodCuenta)"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVer = 718
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseSQLServer_V717 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V716() As Boolean   'Agregada 16 sep 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 716 -----------------------------------
   If lDbVer <= 716 And lUpdOK = True Then
   
      ERR.Clear

      'agregamos campo MontoIngresadoUsuario a tabla Socios
      Q1 = "ALTER TABLE Socios ADD MontoIngresadoUsuario FLOAT NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'agregamos campo MontoTraspasado a tabla Socios
      Q1 = "ALTER TABLE Socios ADD MontoATraspasar FLOAT NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'sólo actualiza la versión para meparejar con versión Access
                       
      If lUpdOK Then
         lDbVer = 717
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V716 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V715() As Boolean   'Agregada 9 sep 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 715 -----------------------------------
   If lDbVer <= 715 And lUpdOK = True Then
   
      ERR.Clear
      
      'sólo actualiza la versión para meparejar con versión Access
                       
      If lUpdOK Then
         lDbVer = 716
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V715 = lUpdOK

End Function
Public Function CorrigeBaseSQLServer_V714() As Boolean   'Agregada 5 ago 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 714 -----------------------------------
   If lDbVer <= 714 And lUpdOK = True Then
   
      ERR.Clear
      
      'sólo actualiza la versión para meparejar con versión Access
                       
      If lUpdOK Then
         lDbVer = 715
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V714 = lUpdOK

End Function
Public Function CorrigeBaseSQLServer_V713() As Boolean   'Agregada 1 jun 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 713 -----------------------------------
   If lDbVer <= 713 And lUpdOK = True Then
   
      ERR.Clear
      
      'agregamos nuevas franquicias a Tabla Empresa
      Q1 = "ALTER TABLE Empresa ADD Franq14ASemiIntegrado BIT NULL, FranqProPymeGeneral BIT NULL, "
      Q1 = Q1 & "FranqProPymeTransp BIT NULL, FranqRentasPresuntas BIT NULL, "
      Q1 = Q1 & "FranqRentaEfectiva BIT NULL, FranqOtro BIT NULL, "
      Q1 = Q1 & "FranqNoSujetoArt14 BIT NULL ; "
      Call ExecSQL(DbMain, Q1)
      
                       
      If lUpdOK Then
         lDbVer = 714
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V713 = lUpdOK

End Function

Public Function CorrigeBaseSQLServer_V712() As Boolean   'Agregada 6 mayo 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 712 -----------------------------------
   If lDbVer <= 712 And lUpdOK = True Then
   
      ERR.Clear
      
      'Agregamos campo TipoDepLey21210 a MovActivoFijo
      Q1 = "ALTER TABLE MovActivoFijo ADD TipoDepLey21210 TINYINT NULL;"
      Call ExecSQL(DbMain, Q1)
                 
      'Agregamos campo TipoDepLey21210 a MovActivoFijo
      Q1 = "ALTER TABLE MovActivoFijo ADD DepDecimaParte2 SMALLINT NULL;"
      Call ExecSQL(DbMain, Q1)
                 
      'Agregamos campo TipoDepLey21210 a MovActivoFijo
      Q1 = "ALTER TABLE MovActivoFijo ADD DepDecimaParte2Hist SMALLINT  NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo PatenteRol a MovActivoFijo
      Q1 = "ALTER TABLE MovActivoFijo ADD PatenteRol VARCHAR (30) NULL;"
      Call ExecSQL(DbMain, Q1)
      
      'Agregamos campo NombreProy a MovActivoFijo
      Q1 = "ALTER TABLE MovActivoFijo ADD NombreProy VARCHAR (60)  NULL;"
      Call ExecSQL(DbMain, Q1)
                 
      'Agregamos campo FechaProy a MovActivoFijo
      Q1 = "ALTER TABLE MovActivoFijo ADD FechaProy INT  NULL;"
      Call ExecSQL(DbMain, Q1)
                 
      'Agregamos campo TipoDepLey21210 a MovActivoFijo
      Q1 = "ALTER TABLE MovActivoFijo ADD TipoDepLey21210Hist TINYINT NULL;"
      Call ExecSQL(DbMain, Q1)
                 
      If lUpdOK Then
         lDbVer = 713
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V712 = lUpdOK

End Function
            

Public Function CorrigeBaseSQLServer_V711() As Boolean   'agregada 2 mar 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 711 -----------------------------------

   If lDbVer = 711 And lUpdOK = True Then
   
      'Agregamos tabla CtasAjustesExContRLI
      Q1 = "CREATE TABLE CtasAjustesExContRLI ("
      Q1 = Q1 & " IdCtaAjustesRLI int IDENTITY (1,1) NOT NULL,"
      Q1 = Q1 & " IdEmpresa int,"
      Q1 = Q1 & " Ano smallint, "
      Q1 = Q1 & " TipoAjuste tinyint,"
      Q1 = Q1 & " IdGrupo smallint,"
      Q1 = Q1 & " IdItem smallint,"
      Q1 = Q1 & " IdCuenta int,"
      Q1 = Q1 & " CodCuenta varchar(15),"
      Q1 = Q1 & " CONSTRAINT Idx PRIMARY KEY (IdCtaAjustesRLI) "
      Q1 = Q1 & ");"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE INDEX IdxItem ON CtasAjustesExContRLI (IdEmpresa, Ano, TipoAjuste, IdGrupo, IdItem)"
      Call ExecSQL(DbMain, Q1)

      If lUpdOK Then
         lDbVer = 712
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseSQLServer_V711 = lUpdOK

End Function
Public Function CorrigeBaseSQLServer_V710() As Boolean   'Agregada 28 ene 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim QBase As String
  

   On Error Resume Next
   
   '--------------------- Versión 710 -----------------------------------
   If lDbVer <= 710 And lUpdOK = True Then
   
      ERR.Clear
      
       'Agregamos campo DocOtroEsCargo a tabla Documento (esto para traer los OtrosDocs del año anterior)
      Q1 = "ALTER TABLE Documento ADD DocOtroEsCargo TINYINT NULL;"
      Call ExecSQL(DbMain, Q1)
                 
      If lUpdOK Then
         lDbVer = 711
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseSQLServer_V710 = lUpdOK

End Function
Public Function CorrigeBaseSQLServer_V702() As Boolean   'agregada 6 ene 2020
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim QBase As String, QEnd As String


   On Error Resume Next

   '--------------------- Versión 702 -----------------------------------

   If lDbVer = 702 And lUpdOK = True Then
      
      Q1 = "ALTER TABLE Socios ADD CantAcciones INT NULL;"
      Call ExecSQL(DbMain, Q1)
           
      If lUpdOK Then
         lDbVer = 703
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseSQLServer_V702 = lUpdOK
   
End Function

Public Function CorrigeBaseSQLServer_V701() As Boolean   'agregada 27 dic 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim Fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim QBase As String, QEnd As String


   On Error Resume Next

   '--------------------- Versión 701 -----------------------------------

   If lDbVer = 701 And lUpdOK = True Then
      
      Q1 = "ALTER TABLE Empresa ADD FranqSocProfPrimCat BIT NULL, FranqSocProfSegCat BIT NULL ;"
      Call ExecSQL(DbMain, Q1)
      
      'cambiamos código plan SII para cuenta Deudores incobrables en la tabla Cuentas
      QBase = "UPDATE Cuentas "

      Q1 = "SET CodCtaPlanSII ='3030700' WHERE Codigo = '3010644' AND " & GenLike(DbMain, "Deudores Incobrables", "Descripcion")
      Call ExecSQL(DbMain, QBase & Q1)
     
      If lUpdOK Then
         lDbVer = 702
         Q1 = "UPDATE ParamEmpresa SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseSQLServer_V701 = lUpdOK
   
End Function

Private Sub CreateTablasODF()
Dim Q1 As String

      Q1 = "CREATE TABLE [dbo].[DocumentoFull]( "
      Q1 = Q1 & " [IdDoc] [int] IDENTITY(1,1) NOT NULL,"
      Q1 = Q1 & " [IdEmpresa] [int] NULL,"
      Q1 = Q1 & " [Ano] [smallint] NULL,"
      Q1 = Q1 & " [IdCompCent] [int] NULL,"
      Q1 = Q1 & " [IdCompPago] [int] NULL,"
      Q1 = Q1 & " [TipoLib] [tinyint] NULL,"
      Q1 = Q1 & " [TipoDoc] [tinyint] NULL,"
      Q1 = Q1 & " [NumDoc] [varchar](20) NULL,"
      Q1 = Q1 & " [NumDocHasta] [varchar](20) NULL,"
      Q1 = Q1 & " [IdEntidad] [int] NULL,"
      Q1 = Q1 & " [TipoEntidad] [smallint] NULL,"
      Q1 = Q1 & " [RutEntidad] [varchar](12) NULL,"
      Q1 = Q1 & " [NombreEntidad] [varchar](50) NULL,"
      Q1 = Q1 & " [FEmision] [int] NULL,"
      Q1 = Q1 & " [FVenc] [int] NULL,"
      Q1 = Q1 & " [Descrip] [varchar](100) NULL,"
      Q1 = Q1 & " [Estado] [smallint] NULL,"
      Q1 = Q1 & " [Exento] [float] NULL,"
      Q1 = Q1 & " [IdCuentaExento] [int] NULL,"
      Q1 = Q1 & " [Afecto] [float] NULL,"
      Q1 = Q1 & " [IdCuentaAfecto] [int] NULL,"
      Q1 = Q1 & " [IVA] [float] NULL,"
      Q1 = Q1 & " [IdCuentaIVA] [int] NULL,"
      Q1 = Q1 & " [OtroImp] [float] NULL,"
      Q1 = Q1 & " [IdCuentaOtroImp] [int] NULL,"
      Q1 = Q1 & " [Total] [float] NULL,"
      Q1 = Q1 & " [IdCuentaTotal] [int] NULL,"
      Q1 = Q1 & " [IdUsuario] [int] NULL,"
      Q1 = Q1 & " [FechaCreacion] [int] NULL,"
      Q1 = Q1 & " [FEmisionOri] [int] NULL,"
      Q1 = Q1 & " [CorrInterno] [int] NULL,"
      Q1 = Q1 & " [SaldoDoc] [float] NULL,"
      Q1 = Q1 & " [FExported] [int] NULL,"
      Q1 = Q1 & " [OldIdDoc] [int] NULL,"
      Q1 = Q1 & " [DTE] [bit] NULL,"
      Q1 = Q1 & " [PorcentRetencion] [tinyint] NULL,"
      Q1 = Q1 & " [TipoRetencion] [tinyint] NULL,"
      Q1 = Q1 & " [MovEdited] [bit] NULL,"
      Q1 = Q1 & " [OtrosVal] [float] NULL,"
      Q1 = Q1 & " [FImporF29] [int] NULL,"
      Q1 = Q1 & " [NumDocRef] [varchar](20) NULL,"
      Q1 = Q1 & " [IdCtaBanco] [int] NULL,"
      Q1 = Q1 & " [TipoRelEnt] [smallint] NULL,"
      Q1 = Q1 & " [IdSucursal] [int] NULL,"
      Q1 = Q1 & " [TotPagadoAnoAnt] [float] NULL,"
      Q1 = Q1 & " [FImportSuc] [int] NULL,"
      Q1 = Q1 & " [Giro] [bit] NULL,"
      Q1 = Q1 & " [FacCompraRetParcial] [bit] NULL,"
      Q1 = Q1 & " [IVAIrrecuperable] [smallint] NULL,"
      Q1 = Q1 & " [DocOtrosEnAnalitico] [bit] NULL,"
      Q1 = Q1 & " [OldIdDocTmp] [int] NULL,"
      Q1 = Q1 & " [NumFiscImpr] [varchar](20) NULL,"
      Q1 = Q1 & " [NumInformeZ] [varchar](20) NULL,"
      Q1 = Q1 & " [CantBoletas] [int] NULL,"
      Q1 = Q1 & " [VentasAcumInfZ] [float] NULL,"
      Q1 = Q1 & " [IdDocAsoc] [int] NULL,"
      Q1 = Q1 & " [PropIVA] [smallint] NULL,"
      Q1 = Q1 & " [ValIVAIrrec] [float] NULL,"
      Q1 = Q1 & " [IVAInmueble] [bit] NULL,"
      Q1 = Q1 & " [FImpFacturacion] [int] NULL,"
      Q1 = Q1 & " [CodSIIDTEIVAIrrec] [smallint] NULL,"
      Q1 = Q1 & " [TipoDocAsoc] [smallint] NULL,"
      Q1 = Q1 & " [IVAActFijo] [float] NULL,"
      Q1 = Q1 & " [EntRelacionada] [bit] NULL,"
      Q1 = Q1 & " [NumCuotas] [smallint] NULL,"
      Q1 = Q1 & " [CompraBienRaiz] [bit] NULL,"
      Q1 = Q1 & " [NumDocAsoc] [varchar](20) NULL,"
      Q1 = Q1 & " [DTEDocAsoc] [bit] NULL,"
      Q1 = Q1 & " [IdANegCCosto] [varchar](20) NULL,"
      Q1 = Q1 & " [UrlDTE] [varchar](250) NULL,"
      Q1 = Q1 & " [CodCtaAfectoOld] [varchar](15) NULL,"
      Q1 = Q1 & " [CodCtaExentoOld] [varchar](15) NULL,"
      Q1 = Q1 & " [CodCtaTotalOld] [varchar](15) NULL,"
      Q1 = Q1 & " [DocOtroEsCargo] [tinyint] NULL,"
      Q1 = Q1 & " [ValRet3Porc] [float] NULL,"
      Q1 = Q1 & " [IdCuentaRet3Porc] [int] NULL,"
      Q1 = Q1 & " [Tratamiento] [int] NULL,"
      Q1 = Q1 & " CONSTRAINT [DocumentoFull_IdDoc_PK] PRIMARY KEY CLUSTERED"
      Q1 = Q1 & " ("
      Q1 = Q1 & " [IdDoc] Asc"
      '3401961
      'Q1 = Q1 & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]"
      Q1 = Q1 & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
      '3401961
      Q1 = Q1 & " ) ON [PRIMARY]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_Ano]  DEFAULT ((0)) FOR [Ano]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_IdCompCent]  DEFAULT ((0)) FOR [IdCompCent]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_IdCompPago]  DEFAULT ((0)) FOR [IdCompPago]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_TipoLib]  DEFAULT ((0)) FOR [TipoLib]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_TipoDoc]  DEFAULT ((0)) FOR [TipoDoc]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_NumDoc]  DEFAULT ((0)) FOR [NumDoc]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_NumDocHasta]  DEFAULT ((0)) FOR [NumDocHasta]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_IdEntidad]  DEFAULT ((0)) FOR [IdEntidad]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_TipoEntidad]  DEFAULT ((0)) FOR [TipoEntidad]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_FEmision]  DEFAULT ((0)) FOR [FEmision]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_FVenc]  DEFAULT ((0)) FOR [FVenc]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_Estado]  DEFAULT ((0)) FOR [Estado]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_Exento]  DEFAULT ((0)) FOR [Exento]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_IdCuentaExento]  DEFAULT ((0)) FOR [IdCuentaExento]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_Afecto]  DEFAULT ((0)) FOR [Afecto]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_IdCuentaAfecto]  DEFAULT ((0)) FOR [IdCuentaAfecto]"
      Call ExecSQL(DbMain, Q1)

      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_IVA]  DEFAULT ((0)) FOR [IVA]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_IdCuentaIVA]  DEFAULT ((0)) FOR [IdCuentaIVA]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_OtroImp]  DEFAULT ((0)) FOR [OtroImp]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_IdCuentaOtroImp]  DEFAULT ((0)) FOR [IdCuentaOtroImp]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_Total]  DEFAULT ((0)) FOR [Total]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_IdCuentaTotal]  DEFAULT ((0)) FOR [IdCuentaTotal]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_IdUsuario]  DEFAULT ((0)) FOR [IdUsuario]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[DocumentoFull] ADD  CONSTRAINT [DF_DocumentoFull_FechaCreacion]  DEFAULT ((0)) FOR [FechaCreacion]"
      Call ExecSQL(DbMain, Q1)
      
      
      
        Q1 = "CREATE TABLE [dbo].[ComprobanteFull]("
        Q1 = Q1 & " [IdComp] [int] IDENTITY(1,1) NOT NULL,"
        Q1 = Q1 & " [IdEmpresa] [int] NULL,"
        Q1 = Q1 & " [Ano] [smallint] NULL,"
        Q1 = Q1 & " [Correlativo] [int] NULL,"
        Q1 = Q1 & " [Fecha] [int] NULL,"
        Q1 = Q1 & " [Tipo] [tinyint] NULL,"
        Q1 = Q1 & " [Estado] [tinyint] NULL,"
        Q1 = Q1 & " [Glosa] [varchar](100) NULL,"
        Q1 = Q1 & " [TotalDebe] [float] NULL,"
        Q1 = Q1 & " [TotalHaber] [float] NULL,"
        Q1 = Q1 & " [IdUsuario] [int] NULL,"
        Q1 = Q1 & " [FechaCreacion] [int] NULL,"
        Q1 = Q1 & " [ImpResumido] [bit] NULL,"
        Q1 = Q1 & " [EsCCMM] [bit] NULL,"
        Q1 = Q1 & " [FechaImport] [int] NULL,"
        Q1 = Q1 & " [TipoAjuste] [tinyint] NULL,"
        Q1 = Q1 & " [OtrosIngEg14TER] [bit] NULL,"
        Q1 = Q1 & " CONSTRAINT [ComprobanteFull_IdComp_PK] PRIMARY KEY CLUSTERED"
        Q1 = Q1 & " ("
        Q1 = Q1 & " [idcomp] Asc"
        '3401961
        'Q1 = Q1 & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]"
        Q1 = Q1 & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
        '3401961
        Q1 = Q1 & " ) ON [PRIMARY]"
        Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[ComprobanteFull] ADD  CONSTRAINT [DF_ComprobanteFull_IdEmpresa]  DEFAULT ((0)) FOR [IdEmpresa]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[ComprobanteFull] ADD  CONSTRAINT [DF_ComprobanteFull_Ano]  DEFAULT ((0)) FOR [Ano]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[ComprobanteFull] ADD  CONSTRAINT [DF_ComprobanteFull_Correlativo]  DEFAULT ((0)) FOR [Correlativo]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[ComprobanteFull] ADD  CONSTRAINT [DF_ComprobanteFull_Fecha]  DEFAULT ((0)) FOR [Fecha]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[ComprobanteFull] ADD  CONSTRAINT [DF_ComprobanteFull_Tipo]  DEFAULT ((0)) FOR [Tipo]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[ComprobanteFull] ADD  CONSTRAINT [DF_ComprobanteFull_Estado]  DEFAULT ((0)) FOR [Estado]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[ComprobanteFull] ADD  CONSTRAINT [DF_ComprobanteFull_TotalDebe]  DEFAULT ((0)) FOR [TotalDebe]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[ComprobanteFull] ADD  CONSTRAINT [DF_ComprobanteFull_TotalHaber]  DEFAULT ((0)) FOR [TotalHaber]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[ComprobanteFull] ADD  CONSTRAINT [DF_ComprobanteFull_IdUsuario]  DEFAULT ((0)) FOR [IdUsuario]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[ComprobanteFull] ADD  CONSTRAINT [DF_ComprobanteFull_FechaCreacion]  DEFAULT ((0)) FOR [FechaCreacion]"
      Call ExecSQL(DbMain, Q1)
      
      
      
        Q1 = "CREATE TABLE [dbo].[MovComprobanteFull]("
        Q1 = Q1 & " [IdMov] [int] IDENTITY(1,1) NOT NULL,"
        Q1 = Q1 & " [IdEmpresa] [int] NULL,"
        Q1 = Q1 & " [Ano] [smallint] NULL,"
        Q1 = Q1 & " [IdComp] [int] NULL,"
        Q1 = Q1 & " [IdDoc] [int] NULL,"
        Q1 = Q1 & " [Orden] [int] NULL,"
        Q1 = Q1 & " [IdCuenta] [int] NULL,"
        Q1 = Q1 & " [Debe] [float] NULL,"
        Q1 = Q1 & " [Haber] [float] NULL,"
        Q1 = Q1 & " [Glosa] [varchar](50) NULL,"
        Q1 = Q1 & " [idCCosto] [int] NULL,"
        Q1 = Q1 & " [idAreaNeg] [int] NULL,"
        Q1 = Q1 & " [IdCartola] [int] NULL,"
        Q1 = Q1 & " [DeCentraliz] [bit] NULL,"
        Q1 = Q1 & " [DePago] [bit] NULL,"
        Q1 = Q1 & " [DeRemu] [bit] NULL,"
        Q1 = Q1 & " [Nota] [varchar](120) NULL,"
        Q1 = Q1 & " [IdDocCuota] [int] NULL,"
        Q1 = Q1 & " CONSTRAINT [MovComprobanteFull_IdMov_PK] PRIMARY KEY CLUSTERED"
        Q1 = Q1 & " ("
        Q1 = Q1 & " [idMov] Asc"
        '3401961
        'Q1 = Q1 & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]"
        Q1 = Q1 & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
        '3401961
        Q1 = Q1 & " ) ON [PRIMARY]"
        Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[MovComprobanteFull] ADD  CONSTRAINT [DF_MovComprobanteFull_IdEmpresa]  DEFAULT ((0)) FOR [IdEmpresa]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[MovComprobanteFull] ADD  CONSTRAINT [DF_MovComprobanteFull_Ano]  DEFAULT ((0)) FOR [Ano]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[MovComprobanteFull] ADD  CONSTRAINT [DF_MovComprobanteFull_IdComp]  DEFAULT ((0)) FOR [IdComp]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[MovComprobanteFull] ADD  CONSTRAINT [DF_MovComprobanteFull_IdDoc]  DEFAULT ((0)) FOR [IdDoc]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[MovComprobanteFull] ADD  CONSTRAINT [DF_MovComprobanteFull_Orden]  DEFAULT ((0)) FOR [Orden]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[MovComprobanteFull] ADD  CONSTRAINT [DF_MovComprobanteFull_IdCuenta]  DEFAULT ((0)) FOR [IdCuenta]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[MovComprobanteFull] ADD  CONSTRAINT [DF_MovComprobanteFull_Debe]  DEFAULT ((0)) FOR [Debe]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[MovComprobanteFull] ADD  CONSTRAINT [DF_MovComprobanteFull_Haber]  DEFAULT ((0)) FOR [Haber]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[MovComprobanteFull] ADD  CONSTRAINT [DF_MovComprobanteFull_idCCosto]  DEFAULT ((0)) FOR [idCCosto]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[MovComprobanteFull] ADD  CONSTRAINT [DF_MovComprobanteFull_idAreaNeg]  DEFAULT ((0)) FOR [idAreaNeg]"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "ALTER TABLE [dbo].[MovComprobanteFull] ADD  CONSTRAINT [DF_MovComprobanteFull_IdCartola]  DEFAULT ((0)) FOR [IdCartola]"
      Call ExecSQL(DbMain, Q1)
      

    

End Sub
